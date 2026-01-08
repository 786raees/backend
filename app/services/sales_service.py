"""
Sales Service - Business logic for sales dashboard.

This service:
1. Reads appointment data from Google Sheets
2. Parses the complex row structure
3. Aggregates statistics
4. Updates individual cells
"""

import logging
import re
from datetime import date, timedelta
from typing import Dict, List, Optional, Any

from app.services.google.sheets_client import GoogleSheetsClient
from app.config import settings

logger = logging.getLogger(__name__)


class SalesService:
    """Service for sales appointment data operations."""

    # Configuration constants
    SALES_REPS = ["GLEN", "GREAT REP", "ILAN"]
    DAYS = ["Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"]
    LEAD_SOURCES = ["Google", "Facebook", "Referral", "Vehicles", "Word of Mouth", "Other"]
    REASONS = ["Sold", "No Show", "Price", "Need Completed Sooner", "Not Ready", "Wrong Product", "Competitor", "Other"]

    # Row structure constants
    DAY_HEADER_ROWS = {
        "Monday": 1,
        "Tuesday": 31,
        "Wednesday": 61,
        "Thursday": 91,
        "Friday": 121,
        "Saturday": 151
    }

    # Slots per rep per day (Saturday has only 3 slots)
    SLOTS_PER_DAY = {
        "Monday": 5,
        "Tuesday": 5,
        "Wednesday": 5,
        "Thursday": 5,
        "Friday": 5,
        "Saturday": 3
    }

    REP_SLOT_OFFSETS = {
        "GLEN": 4,
        "GREAT REP": 13,
        "ILAN": 22
    }

    # Saturday has different rep offsets (3 slots per rep instead of 5)
    SATURDAY_REP_SLOT_OFFSETS = {
        "GLEN": 4,
        "GREAT REP": 10,
        "ILAN": 16
    }

    # Column mapping (0-indexed for API)
    COLUMNS = {
        "slot": 0,           # A
        "lead_name": 1,      # B
        "lead_source": 2,    # C
        "appointment_set": 3,     # D
        "appointment_confirmed": 4,  # E
        "appointment_attended": 5,   # F
        "job_sold": 6,       # G
        "reason": 7,         # H
        "conversion": 8,     # I
        "sell_price": 9,     # J - Sold Price ex GST
        "appointment_time": 10,  # K - Appointment Time
        "project_type": 11,      # L - Project Type
        "suburb": 12             # M - Suburb
    }

    # Editable columns (letter format)
    EDITABLE_COLUMNS = ["F", "G", "H", "J", "K", "L"]

    # Project type options
    PROJECT_TYPES = ["Turf Project", "Synthetic Turf Project", "Turf and Landscape project", "New Build Landscape Project"]

    def __init__(self):
        self._client: Optional[GoogleSheetsClient] = None
        self._service = None
        self._spreadsheet_id = settings.sales_spreadsheet_id if hasattr(settings, 'sales_spreadsheet_id') else None

    def _get_service(self):
        """Lazy initialization of sheets service."""
        if self._service is None:
            if self._client is None:
                self._client = GoogleSheetsClient()
            # Call _ensure_service() to initialize the service
            self._service = self._client._ensure_service()
        return self._service

    def _get_spreadsheet_id(self) -> str:
        """Get the sales spreadsheet ID."""
        if self._spreadsheet_id:
            return self._spreadsheet_id
        # Fall back to environment variable or default
        import os
        return os.getenv("SALES_SPREADSHEET_ID", settings.google_spreadsheet_id)

    def get_week_tab_name(self, target_date: date) -> str:
        """
        Get the weekly tab name for any given date.
        Tabs are named after the Monday of that week in "Mon-DD" format.

        Examples:
            2026-01-02 (Thu) -> "Dec-29" (Monday of that week was Dec 29, 2025)
            2026-01-05 (Mon) -> "Jan-05"
            2026-01-07 (Wed) -> "Jan-05"
        """
        days_since_monday = target_date.weekday()  # Monday=0, Sunday=6
        monday_of_week = target_date - timedelta(days=days_since_monday)
        return monday_of_week.strftime("%b-%d")

    def get_day_name(self, target_date: date) -> str:
        """Get weekday name for the target date (Mon-Sat, excludes Sunday)."""
        weekday = target_date.weekday()
        # Monday=0, Tuesday=1, ..., Saturday=5, Sunday=6
        if weekday <= 5:  # Monday through Saturday
            return self.DAYS[weekday]
        return None  # Sunday

    def calculate_row_number(self, day: str, rep: str, slot: int) -> int:
        """
        Calculate the 1-indexed row number for a specific appointment slot.

        Args:
            day: "Monday", "Tuesday", "Wednesday", "Thursday", or "Friday"
            rep: "GLEN", "GREAT REP", or "ILAN"
            slot: 1, 2, 3, 4, or 5

        Returns:
            Row number (1-indexed for Google Sheets API)
        """
        day_base = self.DAY_HEADER_ROWS.get(day, 1)
        # Use different offsets for Saturday
        if day == "Saturday":
            rep_offset = self.SATURDAY_REP_SLOT_OFFSETS.get(rep, 4)
        else:
            rep_offset = self.REP_SLOT_OFFSETS.get(rep, 4)
        return day_base + rep_offset + (slot - 1)

    def parse_boolean(self, value: Any) -> bool:
        """Convert 'Yes'/empty to boolean."""
        if value is None:
            return False
        if isinstance(value, bool):
            return value
        if isinstance(value, str):
            return value.strip().lower() == "yes"
        return False

    def parse_currency(self, value: Any) -> float:
        """Parse currency string to float (e.g., '$50,000' -> 50000.0)."""
        if value is None:
            return 0.0
        if isinstance(value, (int, float)):
            return float(value)
        if isinstance(value, str):
            # Remove currency symbols, commas, and whitespace
            cleaned = value.strip().replace('$', '').replace(',', '').replace(' ', '')
            if cleaned:
                try:
                    return float(cleaned)
                except ValueError:
                    return 0.0
        return 0.0

    def format_display_date(self, target_date: date) -> str:
        """Format date as 'Thu 2nd Jan'."""
        days = ['Mon', 'Tue', 'Wed', 'Thu', 'Fri', 'Sat', 'Sun']
        months = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun',
                  'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

        day_num = target_date.day
        suffix = self._get_ordinal_suffix(day_num)

        return f"{days[target_date.weekday()]} {day_num}{suffix} {months[target_date.month - 1]}"

    def _get_ordinal_suffix(self, n: int) -> str:
        """Get ordinal suffix for a number."""
        if 11 <= n <= 13:
            return 'th'
        return {1: 'st', 2: 'nd', 3: 'rd'}.get(n % 10, 'th')

    def _parse_row_to_appointment(self, row: List, slot: int, row_number: int) -> Dict:
        """Parse a single row into an appointment dict."""
        def safe_get(idx: int, default: str = "") -> str:
            if idx < len(row):
                val = row[idx]
                return str(val) if val is not None else default
            return default

        return {
            "slot": slot,
            "row_number": row_number,
            "lead_name": safe_get(self.COLUMNS["lead_name"]),
            "lead_source": safe_get(self.COLUMNS["lead_source"]),
            "appointment_set": self.parse_boolean(safe_get(self.COLUMNS["appointment_set"])),
            "appointment_confirmed": self.parse_boolean(safe_get(self.COLUMNS["appointment_confirmed"])),
            "appointment_attended": self.parse_boolean(safe_get(self.COLUMNS["appointment_attended"])),
            "job_sold": self.parse_boolean(safe_get(self.COLUMNS["job_sold"])),
            "reason": safe_get(self.COLUMNS["reason"]),
            "sell_price": self.parse_currency(safe_get(self.COLUMNS["sell_price"])),
            "appointment_time": safe_get(self.COLUMNS["appointment_time"]),
            "project_type": safe_get(self.COLUMNS["project_type"]),
            "suburb": safe_get(self.COLUMNS["suburb"])
        }

    async def get_daily_schedule(self, target_date: date) -> Dict:
        """Get all appointments for a specific date."""
        # Check if Sunday (only Sunday is excluded, Saturday has 3 slots)
        if target_date.weekday() == 6:  # Sunday
            return {
                "success": False,
                "error": "No appointments on Sundays"
            }

        week_tab = self.get_week_tab_name(target_date)
        day_name = self.get_day_name(target_date)

        if not day_name:
            return {
                "success": False,
                "error": "Invalid day"
            }

        try:
            service = self._get_service()
            spreadsheet_id = self._get_spreadsheet_id()

            # Check if the week tab exists
            available_weeks = await self.get_available_weeks()
            if week_tab not in available_weeks:
                return {
                    "success": False,
                    "error": f"Week '{week_tab}' not found. Available weeks: {', '.join(available_weeks[:5])}..."
                }

            # Read the entire day's data (all reps)
            # We need rows from the day header to cover all 3 reps
            day_start_row = self.DAY_HEADER_ROWS[day_name]
            # Saturday has fewer rows (3 slots per rep × 3 reps = ~20 rows)
            day_end_row = day_start_row + (20 if day_name == "Saturday" else 28)

            range_notation = f"'{week_tab}'!A{day_start_row}:M{day_end_row}"

            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_notation
            ).execute()

            all_rows = result.get('values', [])

            # Parse appointments for each rep
            reps_data = []
            totals = {
                "total_set": 0,
                "total_confirmed": 0,
                "total_attended": 0,
                "total_sold": 0
            }

            # Get number of slots for this day
            num_slots = self.SLOTS_PER_DAY.get(day_name, 5)

            for rep in self.SALES_REPS:
                rep_appointments = []

                for slot in range(1, num_slots + 1):
                    row_number = self.calculate_row_number(day_name, rep, slot)
                    # Convert to 0-indexed relative to our fetched range
                    relative_row = row_number - day_start_row

                    if 0 <= relative_row < len(all_rows):
                        row_data = all_rows[relative_row]
                        appointment = self._parse_row_to_appointment(row_data, slot, row_number)
                    else:
                        # Empty slot
                        appointment = {
                            "slot": slot,
                            "row_number": row_number,
                            "lead_name": "",
                            "lead_source": "",
                            "appointment_set": False,
                            "appointment_confirmed": False,
                            "appointment_attended": False,
                            "job_sold": False,
                            "reason": "",
                            "sell_price": 0.0,
                            "appointment_time": "",
                            "project_type": "",
                            "suburb": ""
                        }

                    rep_appointments.append(appointment)

                    # Update totals
                    if appointment["appointment_set"]:
                        totals["total_set"] += 1
                    if appointment["appointment_confirmed"]:
                        totals["total_confirmed"] += 1
                    if appointment["appointment_attended"]:
                        totals["total_attended"] += 1
                    if appointment["job_sold"]:
                        totals["total_sold"] += 1

                reps_data.append({
                    "name": rep,
                    "appointments": rep_appointments
                })

            # Calculate conversion rate
            conversion_rate = 0.0
            if totals["total_attended"] > 0:
                conversion_rate = round((totals["total_sold"] / totals["total_attended"]) * 100, 1)

            totals["conversion_rate"] = conversion_rate

            return {
                "success": True,
                "date": target_date.isoformat(),
                "day_name": day_name,
                "display_date": self.format_display_date(target_date),
                "week_tab": week_tab,
                "reps": reps_data,
                "day_totals": totals
            }

        except Exception as e:
            logger.error(f"Error fetching daily schedule: {e}")
            return {
                "success": False,
                "error": str(e)
            }

    async def update_appointment_field(
        self,
        week_tab: str,
        row_number: int,
        column: str,
        value: str
    ) -> Dict:
        """Update a single cell in the spreadsheet."""
        # Validate column
        if column.upper() not in self.EDITABLE_COLUMNS:
            return {
                "success": False,
                "error": f"Column {column} is not editable. Only F, G, H allowed."
            }

        try:
            service = self._get_service()
            spreadsheet_id = self._get_spreadsheet_id()

            # Build cell reference
            cell_ref = f"'{week_tab}'!{column.upper()}{row_number}"

            # Update the cell
            service.spreadsheets().values().update(
                spreadsheetId=spreadsheet_id,
                range=cell_ref,
                valueInputOption="USER_ENTERED",
                body={"values": [[value]]}
            ).execute()

            logger.info(f"Updated {cell_ref} to '{value}'")

            return {
                "success": True,
                "updated": {
                    "week_tab": week_tab,
                    "row_number": row_number,
                    "column": column.upper(),
                    "value": value
                }
            }

        except Exception as e:
            logger.error(f"Error updating appointment: {e}")
            return {
                "success": False,
                "error": str(e)
            }

    async def get_weekly_stats(self, week_tab: str) -> Dict:
        """Aggregate statistics for an entire week."""
        try:
            service = self._get_service()
            spreadsheet_id = self._get_spreadsheet_id()

            # Read entire week's data (rows 1-180 covers Mon-Sat with buffer)
            range_notation = f"'{week_tab}'!A1:M180"

            result = service.spreadsheets().values().get(
                spreadsheetId=spreadsheet_id,
                range=range_notation
            ).execute()

            all_rows = result.get('values', [])

            # Initialize aggregates
            totals = {
                "appointments_set": 0,
                "appointments_confirmed": 0,
                "in_homes_attended": 0,
                "jobs_sold": 0,
                "weekly_sales_total": 0.0
            }

            by_rep = {rep: {
                "name": rep,
                "appointments_set": 0,
                "appointments_confirmed": 0,
                "in_homes_attended": 0,
                "jobs_sold": 0,
                "sales_total": 0.0
            } for rep in self.SALES_REPS}

            by_day = {day: {"day": day, "attended": 0, "sold": 0} for day in self.DAYS}

            lead_source_counts = {source: 0 for source in self.LEAD_SOURCES}

            # Parse each day and rep
            for day in self.DAYS:
                # Get number of slots for this day (3 for Saturday, 5 for weekdays)
                num_slots = self.SLOTS_PER_DAY.get(day, 5)
                for rep in self.SALES_REPS:
                    for slot in range(1, num_slots + 1):
                        row_number = self.calculate_row_number(day, rep, slot)
                        row_idx = row_number - 1  # Convert to 0-indexed

                        if row_idx < len(all_rows):
                            row = all_rows[row_idx]

                            def safe_get(idx: int) -> str:
                                return str(row[idx]) if idx < len(row) and row[idx] else ""

                            is_set = self.parse_boolean(safe_get(self.COLUMNS["appointment_set"]))
                            is_confirmed = self.parse_boolean(safe_get(self.COLUMNS["appointment_confirmed"]))
                            is_attended = self.parse_boolean(safe_get(self.COLUMNS["appointment_attended"]))
                            is_sold = self.parse_boolean(safe_get(self.COLUMNS["job_sold"]))
                            lead_source = safe_get(self.COLUMNS["lead_source"])
                            sell_price = self.parse_currency(safe_get(self.COLUMNS["sell_price"]))

                            if is_set:
                                totals["appointments_set"] += 1
                                by_rep[rep]["appointments_set"] += 1

                                # Count lead source
                                if lead_source in lead_source_counts:
                                    lead_source_counts[lead_source] += 1
                                elif lead_source:
                                    lead_source_counts["Other"] += 1

                            if is_confirmed:
                                totals["appointments_confirmed"] += 1
                                by_rep[rep]["appointments_confirmed"] += 1

                            if is_attended:
                                totals["in_homes_attended"] += 1
                                by_rep[rep]["in_homes_attended"] += 1
                                by_day[day]["attended"] += 1

                            if is_sold:
                                totals["jobs_sold"] += 1
                                by_rep[rep]["jobs_sold"] += 1
                                by_day[day]["sold"] += 1
                                totals["weekly_sales_total"] += sell_price
                                by_rep[rep]["sales_total"] += sell_price

            # Calculate conversion rates
            if totals["in_homes_attended"] > 0:
                totals["conversion_rate"] = round(
                    (totals["jobs_sold"] / totals["in_homes_attended"]) * 100, 1
                )
            else:
                totals["conversion_rate"] = 0.0

            by_rep_list = []
            for rep in self.SALES_REPS:
                rep_data = by_rep[rep]
                if rep_data["in_homes_attended"] > 0:
                    rep_data["conversion_rate"] = round(
                        (rep_data["jobs_sold"] / rep_data["in_homes_attended"]) * 100, 1
                    )
                else:
                    rep_data["conversion_rate"] = 0.0
                by_rep_list.append(rep_data)

            # Format lead source data
            total_set = totals["appointments_set"]
            by_lead_source = []
            for source in self.LEAD_SOURCES:
                count = lead_source_counts[source]
                percentage = round((count / total_set) * 100, 1) if total_set > 0 else 0.0
                by_lead_source.append({
                    "source": source,
                    "count": count,
                    "percentage": percentage
                })

            # Sort by count descending
            by_lead_source.sort(key=lambda x: x["count"], reverse=True)

            # Get week display string
            week_display = self._format_week_display(week_tab)

            # Get available weeks
            available_weeks = await self.get_available_weeks()

            return {
                "success": True,
                "week": week_tab,
                "week_display": week_display,
                "totals": totals,
                "by_rep": by_rep_list,
                "by_lead_source": by_lead_source,
                "by_day": list(by_day.values()),
                "available_weeks": available_weeks
            }

        except Exception as e:
            logger.error(f"Error fetching weekly stats: {e}")
            return {
                "success": False,
                "error": str(e)
            }

    def _format_week_display(self, week_tab: str) -> str:
        """Format week tab as 'Jan 5 - Jan 9, 2026'."""
        try:
            # Parse the week tab (e.g., "Jan-05")
            month_str, day_str = week_tab.split("-")
            months = {
                "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
                "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
                "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
            }
            month = months.get(month_str, 1)
            day = int(day_str)

            # Assume current/next year
            from datetime import datetime
            year = datetime.now().year
            if month < datetime.now().month - 6:
                year += 1

            monday = date(year, month, day)
            friday = monday + timedelta(days=4)

            if monday.month == friday.month:
                return f"{month_str} {monday.day} - {friday.day}, {year}"
            else:
                friday_month = friday.strftime("%b")
                return f"{month_str} {monday.day} - {friday_month} {friday.day}, {year}"
        except Exception:
            return week_tab

    async def get_available_weeks(self) -> List[str]:
        """Get list of available week tabs."""
        try:
            service = self._get_service()
            spreadsheet_id = self._get_spreadsheet_id()

            # Get spreadsheet metadata
            spreadsheet = service.spreadsheets().get(
                spreadsheetId=spreadsheet_id
            ).execute()

            sheets = spreadsheet.get('sheets', [])
            week_pattern = re.compile(r'^[A-Z][a-z]{2}-\d{2}$')

            # Parse tab names into (date, tab_name) tuples
            week_dates = []
            for sheet in sheets:
                title = sheet.get('properties', {}).get('title', '')
                if week_pattern.match(title):
                    try:
                        # Parse tab name (e.g., "Jan-05")
                        month_str, day_str = title.split("-")
                        months = {
                            "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
                            "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
                            "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
                        }
                        month = months.get(month_str, 1)
                        day = int(day_str)

                        # Determine year (2026 for Jan onwards, 2025 for Dec)
                        year = 2026 if month >= 1 else 2025
                        if month == 12:  # December tabs are from 2025
                            year = 2025

                        tab_date = date(year, month, day)
                        week_dates.append((tab_date, title))
                    except (ValueError, KeyError):
                        # Skip invalid tab names
                        continue

            # Filter out dates before January 5, 2026
            cutoff_date = date(2026, 1, 5)
            week_dates = [(d, t) for d, t in week_dates if d >= cutoff_date]

            # Sort by date (most recent first)
            week_dates.sort(key=lambda x: x[0], reverse=True)

            # Return just the tab names
            available = [title for _, title in week_dates]

            return available

        except Exception as e:
            logger.error(f"Error fetching available weeks: {e}")
            return []

    async def get_daily_stats(self, target_date: date) -> Dict:
        """Get statistics for a single day."""
        if target_date.weekday() == 6:  # Sunday only
            return {
                "success": False,
                "error": "No appointments on Sundays"
            }

        week_tab = self.get_week_tab_name(target_date)
        day_name = self.get_day_name(target_date)

        try:
            # Get full week stats and filter to one day
            week_result = await self.get_weekly_stats(week_tab)

            if not week_result.get("success"):
                return week_result

            # Filter by_day to just this day
            by_day = week_result.get("by_day", [])
            day_data = next((d for d in by_day if d["day"] == day_name), {"attended": 0, "sold": 0})

            # Calculate day totals from day_data
            totals = {
                "appointments_set": 0,
                "appointments_confirmed": 0,
                "in_homes_attended": day_data.get("attended", 0),
                "jobs_sold": day_data.get("sold", 0),
                "conversion_rate": 0.0
            }

            if totals["in_homes_attended"] > 0:
                totals["conversion_rate"] = round(
                    (totals["jobs_sold"] / totals["in_homes_attended"]) * 100, 1
                )

            return {
                "success": True,
                "view": "day",
                "week": week_tab,
                "week_display": f"{day_name}, {self.format_display_date(target_date)}",
                "totals": totals,
                "by_rep": week_result.get("by_rep", []),
                "by_lead_source": week_result.get("by_lead_source", []),
                "by_day": [day_data]
            }

        except Exception as e:
            logger.error(f"Error fetching daily stats: {e}")
            return {"success": False, "error": str(e)}

    async def get_monthly_stats(self, month_str: str) -> Dict:
        """Get aggregated statistics for a month."""
        try:
            # Parse month string (e.g., "2026-01")
            year, month = map(int, month_str.split("-"))

            # Get all available weeks
            available_weeks = await self.get_available_weeks()

            # Filter weeks that fall within this month
            weeks_in_month = []
            for week_tab in available_weeks:
                try:
                    # Parse week tab (e.g., "Jan-05")
                    month_abbr, day_str = week_tab.split("-")
                    months = {
                        "Jan": 1, "Feb": 2, "Mar": 3, "Apr": 4,
                        "May": 5, "Jun": 6, "Jul": 7, "Aug": 8,
                        "Sep": 9, "Oct": 10, "Nov": 11, "Dec": 12
                    }
                    week_month = months.get(month_abbr, 0)
                    week_day = int(day_str)

                    # Check if this week's Monday falls in the target month
                    if week_month == month:
                        weeks_in_month.append(week_tab)
                    # Also include if week spans into target month
                    elif week_month == month - 1 or (month == 1 and week_month == 12):
                        # Check if Friday of this week is in target month
                        from datetime import datetime
                        try:
                            monday = date(year if week_month <= month else year - 1, week_month, week_day)
                            friday = monday + timedelta(days=4)
                            if friday.month == month:
                                weeks_in_month.append(week_tab)
                        except:
                            pass
                except:
                    continue

            # Aggregate stats from all weeks
            totals = {
                "appointments_set": 0,
                "appointments_confirmed": 0,
                "in_homes_attended": 0,
                "jobs_sold": 0,
                "weekly_sales_total": 0.0,
                "conversion_rate": 0.0
            }

            by_rep_agg = {rep: {
                "name": rep,
                "appointments_set": 0,
                "appointments_confirmed": 0,
                "in_homes_attended": 0,
                "jobs_sold": 0,
                "conversion_rate": 0.0
            } for rep in self.SALES_REPS}

            for week_tab in weeks_in_month:
                week_stats = await self.get_weekly_stats(week_tab)
                if week_stats.get("success"):
                    week_totals = week_stats.get("totals", {})
                    totals["appointments_set"] += week_totals.get("appointments_set", 0)
                    totals["appointments_confirmed"] += week_totals.get("appointments_confirmed", 0)
                    totals["in_homes_attended"] += week_totals.get("in_homes_attended", 0)
                    totals["jobs_sold"] += week_totals.get("jobs_sold", 0)
                    totals["weekly_sales_total"] += week_totals.get("weekly_sales_total", 0.0)

                    for rep_data in week_stats.get("by_rep", []):
                        rep_name = rep_data["name"]
                        if rep_name in by_rep_agg:
                            by_rep_agg[rep_name]["appointments_set"] += rep_data.get("appointments_set", 0)
                            by_rep_agg[rep_name]["appointments_confirmed"] += rep_data.get("appointments_confirmed", 0)
                            by_rep_agg[rep_name]["in_homes_attended"] += rep_data.get("in_homes_attended", 0)
                            by_rep_agg[rep_name]["jobs_sold"] += rep_data.get("jobs_sold", 0)

            # Calculate conversion rates
            if totals["in_homes_attended"] > 0:
                totals["conversion_rate"] = round(
                    (totals["jobs_sold"] / totals["in_homes_attended"]) * 100, 1
                )

            by_rep_list = []
            for rep in self.SALES_REPS:
                rep_data = by_rep_agg[rep]
                if rep_data["in_homes_attended"] > 0:
                    rep_data["conversion_rate"] = round(
                        (rep_data["jobs_sold"] / rep_data["in_homes_attended"]) * 100, 1
                    )
                by_rep_list.append(rep_data)

            # Format month display
            from calendar import month_name
            month_display = f"{month_name[month]} {year}"

            return {
                "success": True,
                "view": "month",
                "week_display": month_display,
                "totals": totals,
                "by_rep": by_rep_list,
                "by_lead_source": [],
                "by_day": [],
                "weeks_included": weeks_in_month
            }

        except Exception as e:
            logger.error(f"Error fetching monthly stats: {e}")
            return {"success": False, "error": str(e)}

    async def get_annual_stats(self, year_str: str) -> Dict:
        """Get aggregated statistics for an entire year."""
        try:
            year = int(year_str)

            # Get all available weeks
            available_weeks = await self.get_available_weeks()

            # We'll aggregate all weeks for simplicity (assuming all are in the target year)
            totals = {
                "appointments_set": 0,
                "appointments_confirmed": 0,
                "in_homes_attended": 0,
                "jobs_sold": 0,
                "weekly_sales_total": 0.0,
                "conversion_rate": 0.0
            }

            by_rep_agg = {rep: {
                "name": rep,
                "appointments_set": 0,
                "appointments_confirmed": 0,
                "in_homes_attended": 0,
                "jobs_sold": 0,
                "conversion_rate": 0.0
            } for rep in self.SALES_REPS}

            weeks_included = 0
            for week_tab in available_weeks:
                week_stats = await self.get_weekly_stats(week_tab)
                if week_stats.get("success"):
                    weeks_included += 1
                    week_totals = week_stats.get("totals", {})
                    totals["appointments_set"] += week_totals.get("appointments_set", 0)
                    totals["appointments_confirmed"] += week_totals.get("appointments_confirmed", 0)
                    totals["in_homes_attended"] += week_totals.get("in_homes_attended", 0)
                    totals["jobs_sold"] += week_totals.get("jobs_sold", 0)
                    totals["weekly_sales_total"] += week_totals.get("weekly_sales_total", 0.0)

                    for rep_data in week_stats.get("by_rep", []):
                        rep_name = rep_data["name"]
                        if rep_name in by_rep_agg:
                            by_rep_agg[rep_name]["appointments_set"] += rep_data.get("appointments_set", 0)
                            by_rep_agg[rep_name]["appointments_confirmed"] += rep_data.get("appointments_confirmed", 0)
                            by_rep_agg[rep_name]["in_homes_attended"] += rep_data.get("in_homes_attended", 0)
                            by_rep_agg[rep_name]["jobs_sold"] += rep_data.get("jobs_sold", 0)

            # Calculate conversion rates
            if totals["in_homes_attended"] > 0:
                totals["conversion_rate"] = round(
                    (totals["jobs_sold"] / totals["in_homes_attended"]) * 100, 1
                )

            by_rep_list = []
            for rep in self.SALES_REPS:
                rep_data = by_rep_agg[rep]
                if rep_data["in_homes_attended"] > 0:
                    rep_data["conversion_rate"] = round(
                        (rep_data["jobs_sold"] / rep_data["in_homes_attended"]) * 100, 1
                    )
                by_rep_list.append(rep_data)

            return {
                "success": True,
                "view": "annual",
                "week_display": f"Year {year} ({weeks_included} weeks)",
                "totals": totals,
                "by_rep": by_rep_list,
                "by_lead_source": [],
                "by_day": []
            }

        except Exception as e:
            logger.error(f"Error fetching annual stats: {e}")
            return {"success": False, "error": str(e)}

    async def get_ceo_summary(self, week_tab: Optional[str] = None) -> Dict:
        """Get summary metrics for CEO dashboard."""
        if week_tab is None:
            week_tab = self.get_week_tab_name(date.today())

        try:
            stats = await self.get_weekly_stats(week_tab)

            if not stats.get("success"):
                return stats

            totals = stats.get("totals", {})

            return {
                "success": True,
                "week": week_tab,
                "appointments_set": totals.get("appointments_set", 0),
                "appointments_confirmed": totals.get("appointments_confirmed", 0),
                "in_homes_attended": totals.get("in_homes_attended", 0),
                "jobs_sold": totals.get("jobs_sold", 0),
                "conversion_rate": totals.get("conversion_rate", 0.0),
                "weekly_sales_total": totals.get("weekly_sales_total", 0.0)
            }

        except Exception as e:
            logger.error(f"Error fetching CEO summary: {e}")
            return {
                "success": False,
                "error": str(e)
            }


# Global singleton
sales_service = SalesService()
