#!/usr/bin/env python3
"""
Conversion Utilities for Google Sheets MCP Server

This module provides conversion functions for data formats.
"""

from typing import Any, Dict, List, Optional, Tuple, Union
import re
import datetime


def a1_to_grid_range(sheet_id: int, a1_range: str) -> Dict[str, Any]:
    """
    Convert an A1 notation range to a grid range.
    
    Args:
        sheet_id: The ID of the sheet.
        a1_range: The A1 notation range.
    
    Returns:
        A dictionary representing the grid range.
    """
    # Extract sheet name if present
    if '!' in a1_range:
        _, cell_range = a1_range.split('!', 1)
    else:
        cell_range = a1_range
    
    # Handle single cell case
    if ':' not in cell_range:
        col, row = a1_to_row_col(cell_range)
        return {
            'sheetId': sheet_id,
            'startRowIndex': row,
            'endRowIndex': row + 1,
            'startColumnIndex': col,
            'endColumnIndex': col + 1
        }
    
    # Handle range case
    start, end = cell_range.split(':')
    start_col, start_row = a1_to_row_col(start)
    end_col, end_row = a1_to_row_col(end)
    
    # In grid range, end indices are exclusive
    return {
        'sheetId': sheet_id,
        'startRowIndex': start_row,
        'endRowIndex': end_row + 1,
        'startColumnIndex': start_col,
        'endColumnIndex': end_col + 1
    }


def a1_to_row_col(a1_notation: str) -> Tuple[int, int]:
    """
    Convert an A1 notation cell reference to row and column indices.
    
    Args:
        a1_notation: The A1 notation cell reference (e.g., 'A1').
    
    Returns:
        A tuple of (column_index, row_index).
    """
    # Extract column letters and row number
    match = re.match(r'^([A-Za-z]+)([0-9]+)$', a1_notation)
    if not match:
        raise ValueError(f"Invalid A1 notation: {a1_notation}")
    
    col_str, row_str = match.groups()
    
    # Convert column letters to index (0-based)
    col_index = 0
    for c in col_str.upper():
        col_index = col_index * 26 + (ord(c) - ord('A') + 1)
    col_index -= 1  # Adjust to 0-based index
    
    # Convert row number to index (0-based)
    row_index = int(row_str) - 1
    
    return (col_index, row_index)


def row_col_to_a1(col_index: int, row_index: int) -> str:
    """
    Convert row and column indices to A1 notation.
    
    Args:
        col_index: The column index (0-based).
        row_index: The row index (0-based).
    
    Returns:
        The A1 notation cell reference.
    """
    # Convert column index to letters
    col_str = ''
    col_index += 1  # Adjust to 1-based index for conversion
    
    while col_index > 0:
        col_index, remainder = divmod(col_index - 1, 26)
        col_str = chr(ord('A') + remainder) + col_str
    
    # Convert row index to number (1-based)
    row_num = row_index + 1
    
    return f"{col_str}{row_num}"


def grid_range_to_a1(sheet_name: str, grid_range: Dict[str, int]) -> str:
    """
    Convert a grid range to A1 notation.
    
    Args:
        sheet_name: The name of the sheet.
        grid_range: The grid range dictionary.
    
    Returns:
        The A1 notation range.
    """
    # Extract indices
    start_row = grid_range.get('startRowIndex', 0)
    end_row = grid_range.get('endRowIndex', 0)
    start_col = grid_range.get('startColumnIndex', 0)
    end_col = grid_range.get('endColumnIndex', 0)
    
    # Convert to A1 notation
    start_a1 = row_col_to_a1(start_col, start_row)
    end_a1 = row_col_to_a1(end_col - 1, end_row - 1)  # Adjust for exclusive end indices
    
    # Format the range
    if sheet_name:
        # Quote sheet name if it contains special characters
        if any(c in sheet_name for c in [' ', '-', '(', ')', '&']):
            sheet_name = f"'{sheet_name}'"
        return f"{sheet_name}!{start_a1}:{end_a1}"
    else:
        return f"{start_a1}:{end_a1}"


def format_cell_value(value: Any, format_type: str = None) -> Dict[str, Any]:
    """
    Format a value for use in Google Sheets API.
    
    Args:
        value: The value to format.
        format_type: The type of formatting to apply.
    
    Returns:
        A dictionary representing the formatted value.
    """
    if value is None:
        return {'userEnteredValue': {'stringValue': ''}}
    
    if isinstance(value, bool):
        return {'userEnteredValue': {'boolValue': value}}
    
    if isinstance(value, (int, float)):
        return {'userEnteredValue': {'numberValue': value}}
    
    if isinstance(value, str):
        if value.startswith('='):
            return {'userEnteredValue': {'formulaValue': value}}
        else:
            return {'userEnteredValue': {'stringValue': value}}
    
    if isinstance(value, datetime.datetime):
        # Format as ISO 8601 string
        return {'userEnteredValue': {'stringValue': value.isoformat()}}
    
    # Default to string representation
    return {'userEnteredValue': {'stringValue': str(value)}}


def format_values_for_update(values: List[List[Any]]) -> List[Dict[str, Any]]:
    """
    Format a 2D array of values for use in updateCells request.
    
    Args:
        values: The 2D array of values to format.
    
    Returns:
        A list of row dictionaries for use in updateCells request.
    """
    rows = []
    
    for row in values:
        cells = []
        for value in row:
            cells.append(format_cell_value(value))
        
        rows.append({'values': cells})
    
    return rows


def parse_sheet_values(response: Dict[str, Any]) -> List[List[Any]]:
    """
    Parse values from a Google Sheets API response.
    
    Args:
        response: The API response containing values.
    
    Returns:
        A 2D array of values.
    """
    values = response.get('values', [])
    
    # Ensure all rows have the same length
    if values:
        max_cols = max(len(row) for row in values)
        for row in values:
            while len(row) < max_cols:
                row.append('')
    
    return values


def rgb_to_color(red: float, green: float, blue: float, alpha: float = 1.0) -> Dict[str, float]:
    """
    Convert RGB values to a color dictionary.
    
    Args:
        red: The red component (0-1).
        green: The green component (0-1).
        blue: The blue component (0-1).
        alpha: The alpha component (0-1).
    
    Returns:
        A color dictionary for use in Google Sheets API.
    """
    return {
        'red': max(0, min(1, red)),
        'green': max(0, min(1, green)),
        'blue': max(0, min(1, blue)),
        'alpha': max(0, min(1, alpha))
    }


def hex_to_color(hex_color: str) -> Dict[str, float]:
    """
    Convert a hex color string to a color dictionary.
    
    Args:
        hex_color: The hex color string (e.g., '#FF0000' or '#FF0000FF').
    
    Returns:
        A color dictionary for use in Google Sheets API.
    """
    # Remove '#' if present
    hex_color = hex_color.lstrip('#')
    
    # Parse RGB or RGBA
    if len(hex_color) == 6:
        r, g, b = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16)
        a = 1.0
    elif len(hex_color) == 8:
        r, g, b, a = int(hex_color[0:2], 16), int(hex_color[2:4], 16), int(hex_color[4:6], 16), int(hex_color[6:8], 16)
        a = a / 255.0
    else:
        raise ValueError(f"Invalid hex color: {hex_color}")
    
    # Convert to 0-1 range
    return rgb_to_color(r / 255.0, g / 255.0, b / 255.0, a)


def color_to_hex(color: Dict[str, float]) -> str:
    """
    Convert a color dictionary to a hex color string.
    
    Args:
        color: The color dictionary.
    
    Returns:
        A hex color string.
    """
    r = int(color.get('red', 0) * 255)
    g = int(color.get('green', 0) * 255)
    b = int(color.get('blue', 0) * 255)
    a = int(color.get('alpha', 1) * 255)
    
    if a < 255:
        return f"#{r:02x}{g:02x}{b:02x}{a:02x}"
    else:
        return f"#{r:02x}{g:02x}{b:02x}"


def create_text_format(
    bold: bool = None,
    italic: bool = None,
    underline: bool = None,
    strikethrough: bool = None,
    font_family: str = None,
    font_size: int = None,
    foreground_color: Dict[str, float] = None
) -> Dict[str, Any]:
    """
    Create a text format dictionary.
    
    Args:
        bold: Whether the text is bold.
        italic: Whether the text is italic.
        underline: Whether the text is underlined.
        strikethrough: Whether the text is struck through.
        font_family: The font family.
        font_size: The font size in points.
        foreground_color: The foreground color.
    
    Returns:
        A text format dictionary for use in Google Sheets API.
    """
    text_format = {}
    
    if bold is not None:
        text_format['bold'] = bold
    
    if italic is not None:
        text_format['italic'] = italic
    
    if underline is not None:
        text_format['underline'] = underline
    
    if strikethrough is not None:
        text_format['strikethrough'] = strikethrough
    
    if font_family is not None:
        text_format['fontFamily'] = font_family
    
    if font_size is not None:
        text_format['fontSize'] = font_size
    
    if foreground_color is not None:
        text_format['foregroundColor'] = foreground_color
    
    return text_format


def create_number_format(type_: str, pattern: str = None) -> Dict[str, str]:
    """
    Create a number format dictionary.
    
    Args:
        type_: The number format type.
        pattern: The number format pattern.
    
    Returns:
        A number format dictionary for use in Google Sheets API.
    """
    number_format = {'type': type_}
    
    if pattern is not None:
        number_format['pattern'] = pattern
    
    return number_format


def determine_number_format_type(pattern: str) -> str:
    """
    Determine the number format type based on the pattern.
    
    Args:
        pattern: The number format pattern.
    
    Returns:
        The number format type.
    """
    if '%' in pattern:
        return 'PERCENT'
    elif '$' in pattern or '€' in pattern or '£' in pattern:
        return 'CURRENCY'
    elif any(date_part in pattern.upper() for date_part in ['Y', 'M', 'D']):
        return 'DATE'
    elif any(time_part in pattern.upper() for time_part in ['H', 'MM', 'SS']):
        return 'TIME'
    elif '.' in pattern or '#' in pattern or '0' in pattern:
        return 'NUMBER'
    else:
        return 'TEXT'
