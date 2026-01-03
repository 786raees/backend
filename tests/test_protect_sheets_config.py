"""
Comprehensive tests for protect_sheets.py configuration.

Tests cover:
1. Column visibility configuration
2. Column index mapping
3. Editable columns configuration
4. Business logic requirements
"""
import pytest
import sys
import os

# Add scripts directory to path
sys.path.insert(0, os.path.join(os.path.dirname(os.path.dirname(os.path.abspath(__file__))), 'scripts'))

# Import configuration directly from the module
from importlib import import_module


class TestColumnConfiguration:
    """Tests for column configuration in protect_sheets.py."""

    @pytest.fixture
    def config(self):
        """Load protect_sheets configuration."""
        # Read the file and extract configuration
        script_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'scripts',
            'protect_sheets.py'
        )

        config = {}
        with open(script_path, 'r') as f:
            content = f.read()

            # Extract COLUMNS_TO_HIDE
            if 'COLUMNS_TO_HIDE = [' in content:
                # Parse the list format
                import re
                match = re.search(r"COLUMNS_TO_HIDE = \[(.*?)\]", content, re.DOTALL)
                if match:
                    # Parse the ranges
                    ranges_str = match.group(1)
                    config['columns_to_hide'] = []
                    for range_match in re.finditer(r"\{'start': (\d+), 'end': (\d+)\}", ranges_str):
                        config['columns_to_hide'].append({
                            'start': int(range_match.group(1)),
                            'end': int(range_match.group(2))
                        })

            # Extract EDITABLE_COLUMNS
            match = re.search(r"EDITABLE_COLUMNS = \[(.*?)\]", content)
            if match:
                config['editable_columns'] = [
                    int(x.strip()) for x in match.group(1).split(',') if x.strip().isdigit()
                ]

        return config

    def test_columns_to_hide_exists(self, config):
        """Test COLUMNS_TO_HIDE is defined."""
        assert 'columns_to_hide' in config
        assert len(config['columns_to_hide']) > 0

    def test_column_f_is_hidden(self, config):
        """Test Column F (index 5) is hidden - Pallets."""
        hidden_indices = set()
        for r in config['columns_to_hide']:
            for i in range(r['start'], r['end']):
                hidden_indices.add(i)
        assert 5 in hidden_indices, "Column F (Pallets) should be hidden"

    def test_column_g_is_hidden(self, config):
        """Test Column G (index 6) is hidden - Sell $/SQM."""
        hidden_indices = set()
        for r in config['columns_to_hide']:
            for i in range(r['start'], r['end']):
                hidden_indices.add(i)
        assert 6 in hidden_indices, "Column G (Sell $/SQM) should be hidden"

    def test_column_h_is_hidden(self, config):
        """Test Column H (index 7) is hidden - Cost $/SQM."""
        hidden_indices = set()
        for r in config['columns_to_hide']:
            for i in range(r['start'], r['end']):
                hidden_indices.add(i)
        assert 7 in hidden_indices, "Column H (Cost $/SQM) should be hidden"

    def test_column_i_is_visible(self, config):
        """Test Column I (index 8) is VISIBLE - Delivery Fee."""
        hidden_indices = set()
        for r in config['columns_to_hide']:
            for i in range(r['start'], r['end']):
                hidden_indices.add(i)
        assert 8 not in hidden_indices, "Column I (Delivery Fee) should be VISIBLE"

    def test_column_j_is_visible(self, config):
        """Test Column J (index 9) is VISIBLE - Laying Fee."""
        hidden_indices = set()
        for r in config['columns_to_hide']:
            for i in range(r['start'], r['end']):
                hidden_indices.add(i)
        assert 9 not in hidden_indices, "Column J (Laying Fee) should be VISIBLE"

    def test_columns_k_through_q_hidden(self, config):
        """Test Columns K-Q (indices 10-16) are hidden - Financial calcs."""
        hidden_indices = set()
        for r in config['columns_to_hide']:
            for i in range(r['start'], r['end']):
                hidden_indices.add(i)

        for idx in range(10, 17):  # K=10 through Q=16
            assert idx in hidden_indices, f"Column at index {idx} should be hidden"

    def test_editable_columns_defined(self, config):
        """Test EDITABLE_COLUMNS is defined."""
        assert 'editable_columns' in config
        assert len(config['editable_columns']) > 0

    def test_column_a_editable(self, config):
        """Test Column A (index 0) is editable - Day/Slot."""
        assert 0 in config['editable_columns'], "Column A (Day/Slot) should be editable"

    def test_column_b_editable(self, config):
        """Test Column B (index 1) is editable - Variety."""
        assert 1 in config['editable_columns'], "Column B (Variety) should be editable"

    def test_column_c_editable(self, config):
        """Test Column C (index 2) is editable - Suburb."""
        assert 2 in config['editable_columns'], "Column C (Suburb) should be editable"

    def test_column_d_editable(self, config):
        """Test Column D (index 3) is editable - Service Type."""
        assert 3 in config['editable_columns'], "Column D (Service Type) should be editable"

    def test_column_e_editable(self, config):
        """Test Column E (index 4) is editable - SQM."""
        assert 4 in config['editable_columns'], "Column E (SQM) should be editable"

    def test_column_i_editable(self, config):
        """Test Column I (index 8) is editable - Delivery Fee."""
        assert 8 in config['editable_columns'], "Column I (Delivery Fee) should be editable"

    def test_column_j_editable(self, config):
        """Test Column J (index 9) is editable - Laying Fee."""
        assert 9 in config['editable_columns'], "Column J (Laying Fee) should be editable"

    def test_column_f_not_editable(self, config):
        """Test Column F (index 5) is NOT editable - Pallets (formula)."""
        assert 5 not in config['editable_columns'], "Column F (Pallets) should NOT be editable"

    def test_financial_columns_not_editable(self, config):
        """Test financial columns are NOT editable."""
        for idx in range(10, 17):  # K through Q
            assert idx not in config['editable_columns'], \
                f"Financial column at index {idx} should NOT be editable"


class TestColumnMappingRequirements:
    """Tests verifying column configuration matches business requirements."""

    def test_staff_visible_columns(self):
        """Test that staff can see required columns per requirements."""
        # Staff should see: A, B, C, D, E, I, J (0, 1, 2, 3, 4, 8, 9)
        visible_for_staff = {0, 1, 2, 3, 4, 8, 9}

        # Read configuration
        script_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'scripts',
            'protect_sheets.py'
        )

        with open(script_path, 'r') as f:
            content = f.read()

        # Check that visible columns are not in hidden ranges
        import re
        hidden_indices = set()
        for match in re.finditer(r"\{'start': (\d+), 'end': (\d+)\}", content):
            start = int(match.group(1))
            end = int(match.group(2))
            for i in range(start, end):
                hidden_indices.add(i)

        for col_idx in visible_for_staff:
            assert col_idx not in hidden_indices, \
                f"Column index {col_idx} should be visible for staff"

    def test_staff_hidden_columns(self):
        """Test that financial columns are hidden from staff per requirements."""
        # Staff should NOT see: F, G, H, K-Q (5, 6, 7, 10-16)
        hidden_from_staff = {5, 6, 7, 10, 11, 12, 13, 14, 15, 16}

        # Read configuration
        script_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'scripts',
            'protect_sheets.py'
        )

        with open(script_path, 'r') as f:
            content = f.read()

        # Extract hidden indices
        import re
        hidden_indices = set()
        for match in re.finditer(r"\{'start': (\d+), 'end': (\d+)\}", content):
            start = int(match.group(1))
            end = int(match.group(2))
            for i in range(start, end):
                hidden_indices.add(i)

        for col_idx in hidden_from_staff:
            assert col_idx in hidden_indices, \
                f"Column index {col_idx} should be hidden from staff"


class TestHiddenColumnRanges:
    """Tests for the two separate hidden column ranges."""

    @pytest.fixture
    def hidden_ranges(self):
        """Extract hidden column ranges from config."""
        script_path = os.path.join(
            os.path.dirname(os.path.dirname(os.path.abspath(__file__))),
            'scripts',
            'protect_sheets.py'
        )

        ranges = []
        with open(script_path, 'r') as f:
            content = f.read()

        import re
        for match in re.finditer(r"\{'start': (\d+), 'end': (\d+)\}", content):
            ranges.append({
                'start': int(match.group(1)),
                'end': int(match.group(2))
            })

        return ranges

    def test_two_hidden_ranges(self, hidden_ranges):
        """Test there are exactly two hidden ranges."""
        assert len(hidden_ranges) == 2, "Should have exactly two hidden column ranges"

    def test_first_range_covers_f_g_h(self, hidden_ranges):
        """Test first range covers columns F, G, H (5, 6, 7)."""
        first_range = hidden_ranges[0]
        assert first_range['start'] == 5, "First range should start at column F (5)"
        assert first_range['end'] == 8, "First range should end at 8 (exclusive of I)"

    def test_second_range_covers_k_through_q(self, hidden_ranges):
        """Test second range covers columns K through Q (10-16)."""
        second_range = hidden_ranges[1]
        assert second_range['start'] == 10, "Second range should start at column K (10)"
        assert second_range['end'] == 17, "Second range should end at 17 (exclusive)"

    def test_gap_between_ranges_for_i_j(self, hidden_ranges):
        """Test there's a gap between ranges for columns I, J."""
        first_end = hidden_ranges[0]['end']
        second_start = hidden_ranges[1]['start']

        gap_start = first_end  # 8 = I
        gap_end = second_start  # 10 = K

        assert gap_start == 8, "Gap should start at column I (8)"
        assert gap_end == 10, "Gap should end at column K (10)"
        assert gap_end - gap_start == 2, "Gap should cover exactly 2 columns (I, J)"
