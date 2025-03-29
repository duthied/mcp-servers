#!/usr/bin/env python3
"""
Validation Utilities for Google Sheets MCP Server

This module provides validation functions for input parameters.
"""

import re
from typing import Any, Dict, List, Optional, Union


def validate_spreadsheet_id(spreadsheet_id: str) -> bool:
    """
    Validate a Google Sheets spreadsheet ID.
    
    Args:
        spreadsheet_id: The ID to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    # Spreadsheet IDs are typically 44 characters long and contain letters, numbers, hyphens, and underscores
    pattern = r'^[a-zA-Z0-9_-]{40,50}$'
    return bool(re.match(pattern, spreadsheet_id))


def validate_a1_notation(a1_notation: str) -> bool:
    """
    Validate an A1 notation range.
    
    Args:
        a1_notation: The A1 notation to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    # Basic pattern for A1 notation: optional sheet name followed by cell references
    # Examples: "Sheet1!A1:B2", "A1:B2", "Sheet1!A1", "A1"
    
    # If there's a sheet name, it might be quoted if it contains special characters
    sheet_pattern = r'(?:\'[^\']*\'|"[^"]*"|[^\'\"!]+)?'
    
    # Cell reference pattern: column letters followed by row numbers
    cell_pattern = r'[A-Za-z]+[0-9]+'
    
    # Range pattern: cell reference, optionally followed by colon and another cell reference
    range_pattern = f'{cell_pattern}(?::{cell_pattern})?'
    
    # Full pattern: optional sheet name, followed by exclamation mark if sheet name exists, followed by range
    full_pattern = f'^{sheet_pattern}(?:!{range_pattern}|{range_pattern})$'
    
    return bool(re.match(full_pattern, a1_notation))


def validate_chart_type(chart_type: str) -> bool:
    """
    Validate a chart type.
    
    Args:
        chart_type: The chart type to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    valid_types = ['BAR', 'LINE', 'PIE', 'COLUMN', 'AREA', 'SCATTER']
    return chart_type in valid_types


def validate_merge_type(merge_type: str) -> bool:
    """
    Validate a merge type.
    
    Args:
        merge_type: The merge type to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    valid_types = ['MERGE_ALL', 'MERGE_COLUMNS', 'MERGE_ROWS']
    return merge_type in valid_types


def validate_value_input_option(option: str) -> bool:
    """
    Validate a value input option.
    
    Args:
        option: The option to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    valid_options = ['RAW', 'USER_ENTERED']
    return option in valid_options


def validate_value_render_option(option: str) -> bool:
    """
    Validate a value render option.
    
    Args:
        option: The option to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    valid_options = ['FORMATTED_VALUE', 'UNFORMATTED_VALUE', 'FORMULA']
    return option in valid_options


def validate_insert_data_option(option: str) -> bool:
    """
    Validate an insert data option.
    
    Args:
        option: The option to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    valid_options = ['OVERWRITE', 'INSERT_ROWS']
    return option in valid_options


def validate_sheet_type(sheet_type: str) -> bool:
    """
    Validate a sheet type.
    
    Args:
        sheet_type: The sheet type to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    valid_types = ['spreadsheet', 'worksheet']
    return sheet_type in valid_types


def validate_named_range_scope(scope: str) -> bool:
    """
    Validate a named range scope.
    
    Args:
        scope: The scope to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    valid_scopes = ['WORKBOOK', 'SHEET']
    return scope in valid_scopes


def validate_permission_role(role: str) -> bool:
    """
    Validate a permission role.
    
    Args:
        role: The role to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    valid_roles = ['reader', 'writer', 'commenter', 'owner']
    return role in valid_roles


def validate_permission_type(type_: str) -> bool:
    """
    Validate a permission type.
    
    Args:
        type_: The type to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    valid_types = ['user', 'group', 'domain', 'anyone']
    return type_ in valid_types


def validate_email(email: str) -> bool:
    """
    Validate an email address.
    
    Args:
        email: The email to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    pattern = r'^[a-zA-Z0-9._%+-]+@[a-zA-Z0-9.-]+\.[a-zA-Z]{2,}$'
    return bool(re.match(pattern, email))


def validate_domain(domain: str) -> bool:
    """
    Validate a domain name.
    
    Args:
        domain: The domain to validate.
    
    Returns:
        True if valid, False otherwise.
    """
    pattern = r'^[a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?(\.[a-zA-Z0-9]([a-zA-Z0-9-]{0,61}[a-zA-Z0-9])?)*$'
    return bool(re.match(pattern, domain))


def validate_permissions(permissions: Dict[str, Any]) -> List[str]:
    """
    Validate permissions dictionary.
    
    Args:
        permissions: The permissions dictionary to validate.
    
    Returns:
        A list of validation errors, empty if valid.
    """
    errors = []
    
    # Validate role
    if 'role' in permissions:
        if not validate_permission_role(permissions['role']):
            errors.append(f"Invalid role: {permissions['role']}")
    
    # Validate type
    if 'type' in permissions:
        if not validate_permission_type(permissions['type']):
            errors.append(f"Invalid type: {permissions['type']}")
        
        # Validate type-specific fields
        if permissions['type'] in ['user', 'group']:
            if 'emailAddress' not in permissions:
                errors.append("emailAddress is required for user or group permission type")
            elif not validate_email(permissions['emailAddress']):
                errors.append(f"Invalid email address: {permissions['emailAddress']}")
        
        elif permissions['type'] == 'domain':
            if 'domain' not in permissions:
                errors.append("domain is required for domain permission type")
            elif not validate_domain(permissions['domain']):
                errors.append(f"Invalid domain: {permissions['domain']}")
    
    return errors


def validate_format_options(format_options: Dict[str, Any]) -> List[str]:
    """
    Validate format options dictionary.
    
    Args:
        format_options: The format options dictionary to validate.
    
    Returns:
        A list of validation errors, empty if valid.
    """
    errors = []
    
    # Check for valid keys
    valid_keys = [
        'backgroundColor', 'backgroundColorStyle', 'borders', 'horizontalAlignment',
        'hyperlinkDisplayType', 'numberFormat', 'padding', 'textDirection',
        'textFormat', 'textRotation', 'verticalAlignment', 'wrapStrategy'
    ]
    
    for key in format_options:
        if key not in valid_keys:
            errors.append(f"Invalid format option: {key}")
    
    # Validate specific format options
    if 'backgroundColor' in format_options:
        bg_color = format_options['backgroundColor']
        if not isinstance(bg_color, dict) or not all(k in ['red', 'green', 'blue', 'alpha'] for k in bg_color):
            errors.append("backgroundColor must be a dictionary with red, green, blue keys")
    
    if 'horizontalAlignment' in format_options:
        h_align = format_options['horizontalAlignment']
        valid_h_aligns = ['LEFT', 'CENTER', 'RIGHT']
        if h_align not in valid_h_aligns:
            errors.append(f"Invalid horizontalAlignment: {h_align}")
    
    if 'verticalAlignment' in format_options:
        v_align = format_options['verticalAlignment']
        valid_v_aligns = ['TOP', 'MIDDLE', 'BOTTOM']
        if v_align not in valid_v_aligns:
            errors.append(f"Invalid verticalAlignment: {v_align}")
    
    if 'textFormat' in format_options:
        text_format = format_options['textFormat']
        if not isinstance(text_format, dict):
            errors.append("textFormat must be a dictionary")
        else:
            valid_text_format_keys = [
                'foregroundColor', 'foregroundColorStyle', 'fontFamily',
                'fontSize', 'bold', 'italic', 'strikethrough', 'underline'
            ]
            for key in text_format:
                if key not in valid_text_format_keys:
                    errors.append(f"Invalid textFormat option: {key}")
    
    return errors


def validate_chart_options(options: Dict[str, Any]) -> List[str]:
    """
    Validate chart options dictionary.
    
    Args:
        options: The chart options dictionary to validate.
    
    Returns:
        A list of validation errors, empty if valid.
    """
    errors = []
    
    # Check for valid keys
    valid_top_level_keys = ['title', 'legend', 'hAxis', 'vAxis', 'pieHole']
    
    for key in options:
        if key not in valid_top_level_keys:
            errors.append(f"Invalid chart option: {key}")
    
    # Validate legend
    if 'legend' in options:
        legend = options['legend']
        if not isinstance(legend, dict):
            errors.append("legend must be a dictionary")
        elif 'position' in legend:
            valid_positions = ['RIGHT', 'TOP', 'BOTTOM', 'LEFT', 'NONE']
            if legend['position'] not in valid_positions:
                errors.append(f"Invalid legend position: {legend['position']}")
    
    # Validate axes
    for axis_key in ['hAxis', 'vAxis']:
        if axis_key in options:
            axis = options[axis_key]
            if not isinstance(axis, dict):
                errors.append(f"{axis_key} must be a dictionary")
            else:
                valid_axis_keys = ['title', 'format', 'viewWindow', 'gridlines']
                for key in axis:
                    if key not in valid_axis_keys:
                        errors.append(f"Invalid {axis_key} option: {key}")
    
    # Validate pieHole
    if 'pieHole' in options:
        pie_hole = options['pieHole']
        if not isinstance(pie_hole, (int, float)) or pie_hole < 0 or pie_hole > 0.9:
            errors.append("pieHole must be a number between 0 and 0.9")
    
    return errors
