"""
Turf Manager API Routes - Financial dashboard for Turf Supply.

Endpoints:
- GET /manager-stats - Aggregated financial statistics for Turf Manager dashboard
- GET /weeks - Available week tabs
"""

import logging
import re
from datetime import date, timedelta
from typing import Dict, List, Optional, Any, Literal

from fastapi import APIRouter, HTTPException, Query

from app.services.google.sheets_client import GoogleSheetsClient
from app.config import settings
from app.models.turf_manager import (
    TurfManagerResponse,
    VarietyStats,
    VarietyTotals,
    DeliveryFees,
    LayingStats,
    FinancialTotals,
)

logger = logging.getLogger(__name__)

router = APIRouter(prefix="/turf", tags=["Turf Manager"])


class TurfManagerService:
    """Service for Turf Manager financial data operations."""

    # Column mapping (0-indexed)
    COLUMNS = {
        "day_slot": 0,        # A - Day/Slot
        "variety": 1,         # B - Turf variety
        "suburb": 2,          # C - Suburb
        "service_type": 3,    # D - SL/SD/P
        "sqm": 4,             # E - SQM
        "pallets": 5,         # F - Pallets (auto-calc)
        "sell_per_sqm": 6,    # G - Sell $/SQM
        "cost_per_sqm": 7,    # H - Cost $/SQM
        "delivery_fee": 8,    # I - Delivery Fee
        "laying_fee": 9,      # J - Laying Fee
        "turf_revenue": 10,   # K - Turf Revenue (formula)
        "delivery_t1": 11,    # L - Delivery Total T1
        "delivery_t2": 12,    # M - Delivery Total T2
        "laying_total": 13,   # N - Laying Total
        "total": 14,          # O - Total
        "cost": 15,           # P - Cost
        "margin": 16,         # Q - Margin %
    }

    # Laying cost per SQM
    LAYING_COST_PER_SQM = 2.20

    def __init__(self):
        self._client: Optional[GoogleSheetsClient] = None

    def _get_client(self) -> GoogleSheetsClient:
        """Lazy initialization of sheets client."""
        if self._client is None:
            self._client = GoogleSheetsClient()
        return self._client

    def get_week_tab_name(self, target_date: date) -> str:
        """Get the weekly tab name for any given date."""
        days_since_monday = target_date.weekday()
        monday_of_week = target_date - timedelta(days=days_since_monday)
        return monday_of_week.strftime("%b-%d")

    def get_available_weeks(self) -> List[str]:
        """Get list of available week tabs."""
        client = self._get_client()
        sheets = client.get_available_sheets()

        # Filter to only weekly sheets (format like "Dec-28", "Jan-05")
        week_pattern = re.compile(r"^[A-Z][a-z]{2}-\d{2}$")
        return [s for s in sheets if week_pattern.match(s)]

    def _parse_float(self, value: Any) -> float:
        """Parse a value to float, handling currency formatting."""
        if value is None or value == "":
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        # Remove currency symbols and commas
        cleaned = str(value).replace("$", "").replace(",", "").strip()
        try:
            return float(cleaned)
        except ValueError:
            return 0.0

    def _parse_percentage(self, value: Any) -> float:
        """Parse a percentage value."""
        if value is None or value == "":
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        # Remove % symbol
        cleaned = str(value).replace("%", "").strip()
        try:
            return float(cleaned)
        except ValueError:
            return 0.0

    def _get_week_data(self, week_tab: str) -> List[List[Any]]:
        """Fetch all data from a week tab (columns A through Q)."""
        client = self._get_client()
        try:
            return client.get_worksheet_data(week_tab, range_cols="A:Q")
        except Exception as e:
            logger.error(f"Failed to fetch week data for '{week_tab}': {e}")
            return []

    def _aggregate_week_data(self, rows: List[List[Any]]) -> Dict[str, Any]:
        """Aggregate data from a week's rows."""
        variety_stats: Dict[str, VarietyStats] = {}
        delivery_t1_total = 0.0
        delivery_t2_total = 0.0
        laying_fees_total = 0.0
        laying_costs_total = 0.0

        # Skip header row
        for row in rows[1:]:
            if len(row) < 5:  # Skip rows without essential data
                continue

            # Get variety
            variety = row[self.COLUMNS["variety"]] if len(row) > self.COLUMNS["variety"] else ""
            if not variety or variety.strip() == "":
                continue

            # Parse values
            sqm = self._parse_float(row[self.COLUMNS["sqm"]] if len(row) > self.COLUMNS["sqm"] else 0)
            sell_per_sqm = self._parse_float(row[self.COLUMNS["sell_per_sqm"]] if len(row) > self.COLUMNS["sell_per_sqm"] else 0)
            cost_per_sqm = self._parse_float(row[self.COLUMNS["cost_per_sqm"]] if len(row) > self.COLUMNS["cost_per_sqm"] else 0)
            delivery_fee = self._parse_float(row[self.COLUMNS["delivery_fee"]] if len(row) > self.COLUMNS["delivery_fee"] else 0)
            laying_fee = self._parse_float(row[self.COLUMNS["laying_fee"]] if len(row) > self.COLUMNS["laying_fee"] else 0)

            # Calculate turf revenue and cost
            turf_revenue = sqm * sell_per_sqm
            turf_cost = sqm * cost_per_sqm

            # Update variety stats
            if variety not in variety_stats:
                variety_stats[variety] = VarietyStats(variety=variety)

            variety_stats[variety].sqm_sold += sqm
            variety_stats[variety].sell_price += turf_revenue
            variety_stats[variety].cost += turf_cost

            # Check which truck (determine from day/slot or row position)
            day_slot = str(row[self.COLUMNS["day_slot"]]) if len(row) > 0 else ""
            # Assume T1 vs T2 based on delivery fee columns if present in data
            # For now, split delivery fees evenly or based on row patterns
            # Check if delivery_t1/t2 columns exist and have data
            if len(row) > self.COLUMNS["delivery_t1"]:
                dt1 = self._parse_float(row[self.COLUMNS["delivery_t1"]])
                delivery_t1_total += dt1
            if len(row) > self.COLUMNS["delivery_t2"]:
                dt2 = self._parse_float(row[self.COLUMNS["delivery_t2"]])
                delivery_t2_total += dt2

            # Accumulate laying fees and costs
            laying_fees_total += laying_fee
            # Laying cost is $2.20 per SQM for ALL deliveries
            laying_costs_total += sqm * self.LAYING_COST_PER_SQM

        # Calculate variety totals
        variety_totals = VarietyTotals(
            sqm_sold=sum(v.sqm_sold for v in variety_stats.values()),
            sell_price=sum(v.sell_price for v in variety_stats.values()),
            cost=sum(v.cost for v in variety_stats.values()),
        )

        # Calculate delivery fees
        delivery_fees = DeliveryFees(
            truck_1=delivery_t1_total,
            truck_2=delivery_t2_total,
            total=delivery_t1_total + delivery_t2_total,
        )

        # Calculate laying stats
        laying = LayingStats(
            sales=laying_fees_total,
            costs=laying_costs_total,
        )

        # Calculate financial totals
        total_sales = variety_totals.sell_price + delivery_fees.total + laying.sales
        total_costs = variety_totals.cost + laying.costs
        margin_percent = ((total_sales - total_costs) / total_sales * 100) if total_sales > 0 else 0

        totals = FinancialTotals(
            sales=total_sales,
            costs=total_costs,
            margin_percent=round(margin_percent, 1),
        )

        return {
            "by_variety": list(variety_stats.values()),
            "variety_totals": variety_totals,
            "delivery_fees": delivery_fees,
            "laying": laying,
            "totals": totals,
        }

    async def get_manager_stats(
        self,
        view: Literal["day", "week", "month", "annual"],
        target_date: Optional[date] = None,
        week: Optional[str] = None,
        month: Optional[str] = None,
        year: Optional[int] = None,
    ) -> TurfManagerResponse:
        """Get aggregated statistics for the Turf Manager dashboard."""

        if view == "day":
            if target_date is None:
                target_date = date.today()
            week_tab = self.get_week_tab_name(target_date)
            rows = self._get_week_data(week_tab)

            # Handle empty data
            if not rows:
                aggregated = self._aggregate_week_data([["Header"]])
            else:
                # Find day section in spreadsheet
                # Format: "Monday - Daily Turf Deliveries (6 slots per truck)"
                day_name = target_date.strftime("%A")
                day_rows = [["Header"]]  # Placeholder header
                in_day_section = False

                for row in rows:
                    if len(row) > 0:
                        cell_value = str(row[0])
                        # Check if this is a day section header
                        if " - Daily Turf Deliveries" in cell_value:
                            if cell_value.startswith(day_name):
                                in_day_section = True
                            elif in_day_section:
                                # We've reached the next day's section
                                break
                        elif in_day_section and len(row) >= 5:
                            # This is a data row within the day section
                            # Check if it has variety data (column B should have variety name)
                            if len(row) > 1 and row[1] and str(row[1]).strip() not in ["", "Variety", "TRUCK 1", "TRUCK 2", "Slot"]:
                                day_rows.append(row)

                aggregated = self._aggregate_week_data(day_rows)
            period = target_date.strftime("%A, %B %d, %Y")

        elif view == "week":
            if week is None:
                week = self.get_week_tab_name(date.today())
            rows = self._get_week_data(week)
            if not rows:
                aggregated = self._aggregate_week_data([["Header"]])
            else:
                aggregated = self._aggregate_week_data(rows)
            period = f"Week of {week}"

        elif view == "month":
            if month is None:
                month = date.today().strftime("%Y-%m")
            # Parse month string
            try:
                year_num, month_num = map(int, month.split("-"))
            except ValueError:
                raise ValueError(f"Invalid month format: {month}")

            # Get all weeks that fall in this month
            available_weeks = self.get_available_weeks()
            month_data = []

            for week_name in available_weeks:
                # Parse week name (e.g., "Jan-05")
                try:
                    week_date = date(year_num, 1, 1)  # Start of year
                    week_month_abbr = week_name[:3]
                    week_day = int(week_name[4:])

                    # Map month abbreviation to number
                    month_abbrs = {
                        "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
                        "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
                        "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
                    }
                    week_month_num = month_abbrs.get(week_month_abbr)
                    if week_month_num == month_num:
                        rows = self._get_week_data(week_name)
                        month_data.extend(rows[1:] if len(rows) > 1 else [])
                except (ValueError, IndexError):
                    continue

            # Aggregate all month data
            if month_data:
                aggregated = self._aggregate_week_data([["Header"]] + month_data)
            else:
                aggregated = self._aggregate_week_data([["Header"]])

            from calendar import month_name
            period = f"{month_name[month_num]} {year_num}"

        elif view == "annual":
            if year is None:
                year = date.today().year

            # Get all weeks for the year
            available_weeks = self.get_available_weeks()
            annual_data = []

            for week_name in available_weeks:
                rows = self._get_week_data(week_name)
                annual_data.extend(rows[1:] if len(rows) > 1 else [])

            if annual_data:
                aggregated = self._aggregate_week_data([["Header"]] + annual_data)
            else:
                aggregated = self._aggregate_week_data([["Header"]])

            period = f"Year {year}"

        else:
            raise ValueError(f"Invalid view: {view}")

        return TurfManagerResponse(
            success=True,
            view=view,
            period=period,
            by_variety=aggregated["by_variety"],
            variety_totals=aggregated["variety_totals"],
            delivery_fees=aggregated["delivery_fees"],
            laying=aggregated["laying"],
            totals=aggregated["totals"],
        )


# Global service instance
turf_manager_service = TurfManagerService()


# ============ Endpoints ============

@router.get("/manager-stats", response_model=TurfManagerResponse)
async def get_manager_stats(
    view: Literal["day", "week", "month", "annual"] = Query("week", description="View type"),
    date_param: Optional[str] = Query(None, alias="date", description="Date for day view (YYYY-MM-DD)"),
    week: Optional[str] = Query(None, description="Week tab name, e.g., 'Jan-05'"),
    month: Optional[str] = Query(None, description="Month for month view (YYYY-MM)"),
    year: Optional[int] = Query(None, description="Year for annual view"),
):
    """
    Get aggregated financial statistics for Turf Manager dashboard.

    Query Parameters:
    - view: "day", "week", "month", or "annual"
    - date: For day view, YYYY-MM-DD format. Defaults to today.
    - week: For week view, week tab name. Defaults to current week.
    - month: For month view, YYYY-MM format. Defaults to current month.
    - year: For annual view. Defaults to current year.

    Returns variety breakdown, delivery fees, laying costs, and totals with margin.
    """
    # Parse date parameter if provided
    target_date = None
    if date_param:
        try:
            target_date = date.fromisoformat(date_param)
        except ValueError:
            raise HTTPException(status_code=400, detail="Invalid date format. Use YYYY-MM-DD")

    try:
        result = await turf_manager_service.get_manager_stats(
            view=view,
            target_date=target_date,
            week=week,
            month=month,
            year=year,
        )
        return result

    except HTTPException:
        raise  # Re-raise HTTP exceptions as-is
    except ValueError as e:
        raise HTTPException(status_code=400, detail=str(e))
    except Exception as e:
        logger.error(f"Error fetching manager stats: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to fetch statistics: {str(e)}")


@router.get("/weeks")
async def get_available_weeks():
    """
    Get list of available week tabs for Turf Supply.

    Returns list of week tab names (e.g., ["Jan-05", "Jan-12", "Dec-29"]).
    Useful for populating week selector dropdowns.
    """
    try:
        weeks = turf_manager_service.get_available_weeks()
        return {"success": True, "weeks": weeks}
    except Exception as e:
        logger.error(f"Error fetching available weeks: {e}")
        raise HTTPException(status_code=500, detail=f"Failed to fetch weeks: {str(e)}")
