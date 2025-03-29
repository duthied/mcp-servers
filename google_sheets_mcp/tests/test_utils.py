#!/usr/bin/env python3
"""
Tests for the Google Sheets MCP Server utility modules.

This module contains tests for the validators and converters utility modules.
"""

import os
import sys
import pytest

# Add the parent directory to the path so we can import the package
sys.path.insert(0, os.path.abspath(os.path.join(os.path.dirname(__file__), '..')))

from google_sheets_mcp.src.utils.validators import (
    validate_spreadsheet_id,
    validate_a1_notation,
    validate_chart_type,
    validate_merge_type,
    validate_value_input_option,
    validate_value_render_option,
    validate_insert_data_option,
    validate_sheet_type,
    validate_named_range_scope,
    validate_permission_role,
    validate_permission_type,
    validate_email,
    validate_domain,
    validate_permissions,
    validate_format_options,
    validate_chart_options,
)

from google_sheets_mcp.src.utils.converters import (
    a1_to_grid_range,
    a1_to_row_col,
    row_col_to_a1,
    grid_range_to_a1,
    format_cell_value,
    format_values_for_update,
    parse_sheet_values,
    rgb_to_color,
    hex_to_color,
    color_to_hex,
    create_text_format,
    create_number_format,
    determine_number_format_type,
)


class TestValidators:
    """Tests for the validators module."""

    def test_validate_spreadsheet_id(self):
        """Test validation of spreadsheet IDs."""
        # Valid spreadsheet IDs
        assert validate_spreadsheet_id("1BxiMVs0XRA5nFMdKvBdBZjgmUUqptlbs74OgvE2upms")
        assert validate_spreadsheet_id("1-CQvtC7uU-VRXbB_NsSY5YoFmNMbP9JdmDaXlRmNDkU")
        
        # Invalid spreadsheet IDs
        assert not validate_spreadsheet_id("invalid")
        assert not validate_spreadsheet_id("123")
        assert not validate_spreadsheet_id("")

    def test_validate_a1_notation(self):
        """Test validation of A1 notation ranges."""
        # Valid A1 notation
        assert validate_a1_notation("A1")
        assert validate_a1_notation("Z100")
        assert validate_a1_notation("A1:B2")
        assert validate_a1_notation("Sheet1!A1")
        assert validate_a1_notation("Sheet1!A1:B2")
        assert validate_a1_notation("'Sheet 1'!A1:B2")
        
        # Invalid A1 notation
        assert not validate_a1_notation("A")
        assert not validate_a1_notation("1")
        assert not validate_a1_notation("A:B")
        assert not validate_a1_notation("Sheet1!")

    def test_validate_chart_type(self):
        """Test validation of chart types."""
        # Valid chart types
        assert validate_chart_type("BAR")
        assert validate_chart_type("LINE")
        assert validate_chart_type("PIE")
        assert validate_chart_type("COLUMN")
        assert validate_chart_type("AREA")
        assert validate_chart_type("SCATTER")
        
        # Invalid chart types
        assert not validate_chart_type("INVALID")
        assert not validate_chart_type("bar")  # Case sensitive
        assert not validate_chart_type("")

    def test_validate_merge_type(self):
        """Test validation of merge types."""
        # Valid merge types
        assert validate_merge_type("MERGE_ALL")
        assert validate_merge_type("MERGE_COLUMNS")
        assert validate_merge_type("MERGE_ROWS")
        
        # Invalid merge types
        assert not validate_merge_type("INVALID")
        assert not validate_merge_type("merge_all")  # Case sensitive
        assert not validate_merge_type("")

    def test_validate_value_input_option(self):
        """Test validation of value input options."""
        # Valid value input options
        assert validate_value_input_option("RAW")
        assert validate_value_input_option("USER_ENTERED")
        
        # Invalid value input options
        assert not validate_value_input_option("INVALID")
        assert not validate_value_input_option("raw")  # Case sensitive
        assert not validate_value_input_option("")

    def test_validate_value_render_option(self):
        """Test validation of value render options."""
        # Valid value render options
        assert validate_value_render_option("FORMATTED_VALUE")
        assert validate_value_render_option("UNFORMATTED_VALUE")
        assert validate_value_render_option("FORMULA")
        
        # Invalid value render options
        assert not validate_value_render_option("INVALID")
        assert not validate_value_render_option("formatted_value")  # Case sensitive
        assert not validate_value_render_option("")

    def test_validate_insert_data_option(self):
        """Test validation of insert data options."""
        # Valid insert data options
        assert validate_insert_data_option("OVERWRITE")
        assert validate_insert_data_option("INSERT_ROWS")
        
        # Invalid insert data options
        assert not validate_insert_data_option("INVALID")
        assert not validate_insert_data_option("overwrite")  # Case sensitive
        assert not validate_insert_data_option("")

    def test_validate_sheet_type(self):
        """Test validation of sheet types."""
        # Valid sheet types
        assert validate_sheet_type("spreadsheet")
        assert validate_sheet_type("worksheet")
        
        # Invalid sheet types
        assert not validate_sheet_type("INVALID")
        assert not validate_sheet_type("SPREADSHEET")  # Case sensitive
        assert not validate_sheet_type("")

    def test_validate_named_range_scope(self):
        """Test validation of named range scopes."""
        # Valid named range scopes
        assert validate_named_range_scope("WORKBOOK")
        assert validate_named_range_scope("SHEET")
        
        # Invalid named range scopes
        assert not validate_named_range_scope("INVALID")
        assert not validate_named_range_scope("workbook")  # Case sensitive
        assert not validate_named_range_scope("")

    def test_validate_permission_role(self):
        """Test validation of permission roles."""
        # Valid permission roles
        assert validate_permission_role("reader")
        assert validate_permission_role("writer")
        assert validate_permission_role("commenter")
        assert validate_permission_role("owner")
        
        # Invalid permission roles
        assert not validate_permission_role("INVALID")
        assert not validate_permission_role("READER")  # Case sensitive
        assert not validate_permission_role("")

    def test_validate_permission_type(self):
        """Test validation of permission types."""
        # Valid permission types
        assert validate_permission_type("user")
        assert validate_permission_type("group")
        assert validate_permission_type("domain")
        assert validate_permission_type("anyone")
        
        # Invalid permission types
        assert not validate_permission_type("INVALID")
        assert not validate_permission_type("USER")  # Case sensitive
        assert not validate_permission_type("")

    def test_validate_email(self):
        """Test validation of email addresses."""
        # Valid email addresses
        assert validate_email("user@example.com")
        assert validate_email("user.name@example.co.uk")
        assert validate_email("user+tag@example.com")
        
        # Invalid email addresses
        assert not validate_email("invalid")
        assert not validate_email("user@")
        assert not validate_email("@example.com")
        assert not validate_email("")

    def test_validate_domain(self):
        """Test validation of domain names."""
        # Valid domain names
        assert validate_domain("example.com")
        assert validate_domain("sub.example.co.uk")
        assert validate_domain("example-domain.com")
        
        # Invalid domain names
        assert not validate_domain("invalid")
        assert not validate_domain(".com")
        assert not validate_domain("example.")
        assert not validate_domain("")

    def test_validate_permissions(self):
        """Test validation of permissions dictionaries."""
        # Valid permissions
        assert validate_permissions({
            "role": "reader",
            "type": "user",
            "emailAddress": "user@example.com"
        }) == []
        
        assert validate_permissions({
            "role": "writer",
            "type": "domain",
            "domain": "example.com"
        }) == []
        
        # Invalid permissions
        errors = validate_permissions({
            "role": "invalid",
            "type": "user",
            "emailAddress": "user@example.com"
        })
        assert len(errors) == 1
        assert "Invalid role" in errors[0]
        
        errors = validate_permissions({
            "role": "reader",
            "type": "user"
        })
        assert len(errors) == 1
        assert "emailAddress is required" in errors[0]
        
        errors = validate_permissions({
            "role": "reader",
            "type": "domain"
        })
        assert len(errors) == 1
        assert "domain is required" in errors[0]

    def test_validate_format_options(self):
        """Test validation of format options dictionaries."""
        # Valid format options
        assert validate_format_options({
            "backgroundColor": {"red": 1.0, "green": 0.0, "blue": 0.0},
            "textFormat": {"bold": True, "fontSize": 12}
        }) == []
        
        assert validate_format_options({
            "horizontalAlignment": "CENTER",
            "verticalAlignment": "MIDDLE"
        }) == []
        
        # Invalid format options
        errors = validate_format_options({
            "invalid": "option"
        })
        assert len(errors) == 1
        assert "Invalid format option" in errors[0]
        
        errors = validate_format_options({
            "horizontalAlignment": "INVALID"
        })
        assert len(errors) == 1
        assert "Invalid horizontalAlignment" in errors[0]

    def test_validate_chart_options(self):
        """Test validation of chart options dictionaries."""
        # Valid chart options
        assert validate_chart_options({
            "title": "Chart Title",
            "legend": {"position": "BOTTOM"},
            "hAxis": {"title": "X-Axis"},
            "vAxis": {"title": "Y-Axis"}
        }) == []
        
        # Invalid chart options
        errors = validate_chart_options({
            "invalid": "option"
        })
        assert len(errors) == 1
        assert "Invalid chart option" in errors[0]
        
        errors = validate_chart_options({
            "legend": {"position": "INVALID"}
        })
        assert len(errors) == 1
        assert "Invalid legend position" in errors[0]


class TestConverters:
    """Tests for the converters module."""

    def test_a1_to_grid_range(self):
        """Test conversion from A1 notation to grid range."""
        # Test single cell
        grid_range = a1_to_grid_range(0, "A1")
        assert grid_range["sheetId"] == 0
        assert grid_range["startRowIndex"] == 0
        assert grid_range["endRowIndex"] == 1
        assert grid_range["startColumnIndex"] == 0
        assert grid_range["endColumnIndex"] == 1
        
        # Test range
        grid_range = a1_to_grid_range(0, "A1:B2")
        assert grid_range["sheetId"] == 0
        assert grid_range["startRowIndex"] == 0
        assert grid_range["endRowIndex"] == 2
        assert grid_range["startColumnIndex"] == 0
        assert grid_range["endColumnIndex"] == 2
        
        # Test with sheet name
        grid_range = a1_to_grid_range(0, "Sheet1!A1:B2")
        assert grid_range["sheetId"] == 0
        assert grid_range["startRowIndex"] == 0
        assert grid_range["endRowIndex"] == 2
        assert grid_range["startColumnIndex"] == 0
        assert grid_range["endColumnIndex"] == 2

    def test_a1_to_row_col(self):
        """Test conversion from A1 notation to row and column indices."""
        # Test simple cases
        assert a1_to_row_col("A1") == (0, 0)
        assert a1_to_row_col("B2") == (1, 1)
        assert a1_to_row_col("Z26") == (25, 25)
        
        # Test multi-letter columns
        assert a1_to_row_col("AA1") == (26, 0)
        assert a1_to_row_col("AB2") == (27, 1)
        
        # Test with invalid input
        with pytest.raises(ValueError):
            a1_to_row_col("invalid")

    def test_row_col_to_a1(self):
        """Test conversion from row and column indices to A1 notation."""
        # Test simple cases
        assert row_col_to_a1(0, 0) == "A1"
        assert row_col_to_a1(1, 1) == "B2"
        assert row_col_to_a1(25, 25) == "Z26"
        
        # Test multi-letter columns
        assert row_col_to_a1(26, 0) == "AA1"
        assert row_col_to_a1(27, 1) == "AB2"

    def test_grid_range_to_a1(self):
        """Test conversion from grid range to A1 notation."""
        # Test without sheet name
        grid_range = {
            "startRowIndex": 0,
            "endRowIndex": 1,
            "startColumnIndex": 0,
            "endColumnIndex": 1
        }
        assert grid_range_to_a1("", grid_range) == "A1:A1"
        
        # Test with sheet name
        assert grid_range_to_a1("Sheet1", grid_range) == "Sheet1!A1:A1"
        
        # Test with sheet name containing special characters
        assert grid_range_to_a1("Sheet 1", grid_range) == "'Sheet 1'!A1:A1"
        
        # Test larger range
        grid_range = {
            "startRowIndex": 0,
            "endRowIndex": 2,
            "startColumnIndex": 0,
            "endColumnIndex": 2
        }
        assert grid_range_to_a1("Sheet1", grid_range) == "Sheet1!A1:B2"

    def test_format_cell_value(self):
        """Test formatting of cell values."""
        # Test various types
        assert format_cell_value(None)["userEnteredValue"]["stringValue"] == ""
        assert format_cell_value(True)["userEnteredValue"]["boolValue"] is True
        assert format_cell_value(123)["userEnteredValue"]["numberValue"] == 123
        assert format_cell_value("text")["userEnteredValue"]["stringValue"] == "text"
        assert format_cell_value("=SUM(A1:B1)")["userEnteredValue"]["formulaValue"] == "=SUM(A1:B1)"

    def test_format_values_for_update(self):
        """Test formatting of values for update requests."""
        values = [
            ["A", 1],
            [True, "=SUM(A1:B1)"]
        ]
        formatted = format_values_for_update(values)
        
        assert len(formatted) == 2
        assert len(formatted[0]["values"]) == 2
        assert formatted[0]["values"][0]["userEnteredValue"]["stringValue"] == "A"
        assert formatted[0]["values"][1]["userEnteredValue"]["numberValue"] == 1
        assert formatted[1]["values"][0]["userEnteredValue"]["boolValue"] is True
        assert formatted[1]["values"][1]["userEnteredValue"]["formulaValue"] == "=SUM(A1:B1)"

    def test_parse_sheet_values(self):
        """Test parsing of sheet values from API responses."""
        response = {
            "values": [
                ["A", "B"],
                ["C", "D"]
            ]
        }
        values = parse_sheet_values(response)
        
        assert len(values) == 2
        assert values[0][0] == "A"
        assert values[0][1] == "B"
        assert values[1][0] == "C"
        assert values[1][1] == "D"
        
        # Test with uneven rows
        response = {
            "values": [
                ["A", "B", "C"],
                ["D"]
            ]
        }
        values = parse_sheet_values(response)
        
        assert len(values) == 2
        assert len(values[0]) == 3
        assert len(values[1]) == 3
        assert values[1][1] == ""
        assert values[1][2] == ""

    def test_rgb_to_color(self):
        """Test conversion from RGB values to color dictionary."""
        color = rgb_to_color(1.0, 0.5, 0.0)
        assert color["red"] == 1.0
        assert color["green"] == 0.5
        assert color["blue"] == 0.0
        assert color["alpha"] == 1.0
        
        # Test with alpha
        color = rgb_to_color(1.0, 0.5, 0.0, 0.5)
        assert color["alpha"] == 0.5
        
        # Test with out-of-range values
        color = rgb_to_color(2.0, -0.5, 1.5)
        assert color["red"] == 1.0
        assert color["green"] == 0.0
        assert color["blue"] == 1.0

    def test_hex_to_color(self):
        """Test conversion from hex color string to color dictionary."""
        # Test RGB format
        color = hex_to_color("#FF8000")
        assert color["red"] == 1.0
        assert abs(color["green"] - 0.5) < 0.01
        assert color["blue"] == 0.0
        assert color["alpha"] == 1.0
        
        # Test RGBA format
        color = hex_to_color("#FF800080")
        assert color["red"] == 1.0
        assert abs(color["green"] - 0.5) < 0.01
        assert color["blue"] == 0.0
        assert color["alpha"] == 0.5
        
        # Test without hash
        color = hex_to_color("FF8000")
        assert color["red"] == 1.0
        assert abs(color["green"] - 0.5) < 0.01
        assert color["blue"] == 0.0
        
        # Test with invalid input
        with pytest.raises(ValueError):
            hex_to_color("invalid")

    def test_color_to_hex(self):
        """Test conversion from color dictionary to hex color string."""
        # Test RGB
        color = {"red": 1.0, "green": 0.5, "blue": 0.0}
        hex_color = color_to_hex(color)
        assert hex_color.lower() == "#ff8000"
        
        # Test RGBA
        color = {"red": 1.0, "green": 0.5, "blue": 0.0, "alpha": 0.5}
        hex_color = color_to_hex(color)
        assert hex_color.lower() == "#ff800080"

    def test_create_text_format(self):
        """Test creation of text format dictionaries."""
        # Test with all options
        text_format = create_text_format(
            bold=True,
            italic=True,
            underline=True,
            strikethrough=False,
            font_family="Arial",
            font_size=12,
            foreground_color={"red": 1.0, "green": 0.0, "blue": 0.0}
        )
        
        assert text_format["bold"] is True
        assert text_format["italic"] is True
        assert text_format["underline"] is True
        assert text_format["strikethrough"] is False
        assert text_format["fontFamily"] == "Arial"
        assert text_format["fontSize"] == 12
        assert text_format["foregroundColor"]["red"] == 1.0
        
        # Test with partial options
        text_format = create_text_format(bold=True)
        assert "bold" in text_format
        assert "italic" not in text_format

    def test_create_number_format(self):
        """Test creation of number format dictionaries."""
        # Test with type only
        number_format = create_number_format("CURRENCY")
        assert number_format["type"] == "CURRENCY"
        assert "pattern" not in number_format
        
        # Test with pattern
        number_format = create_number_format("CURRENCY", "$#,##0.00")
        assert number_format["type"] == "CURRENCY"
        assert number_format["pattern"] == "$#,##0.00"

    def test_determine_number_format_type(self):
        """Test determination of number format type from pattern."""
        assert determine_number_format_type("0%") == "PERCENT"
        assert determine_number_format_type("$0.00") == "CURRENCY"
        assert determine_number_format_type("MM/DD/YYYY") == "DATE"
        assert determine_number_format_type("HH:MM:SS") == "TIME"
        assert determine_number_format_type("0.00") == "NUMBER"
        assert determine_number_format_type("@") == "TEXT"


if __name__ == "__main__":
    pytest.main(["-xvs", __file__])
