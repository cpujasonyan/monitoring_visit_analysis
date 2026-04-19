# -*- coding: utf-8 -*-
"""Monitoring Visit Analysis - Configuration."""

# Excel column mappings (1-based column numbers)
COLUMNS = {
    'site_no': 1,
    'site_name': 2,
    'visit_id': 3,
    'visit_type': 4,
    'cra': 8,
    'co_cra': 9,
    'actual_dates': 13,
    'report_status': 14,
    'report_submit_date': 15,
    'report_final_date': 16,
    'followup_status': 17,
    'followup_date': 18,
    'archive_status': 19,
}

# Working day thresholds
SUBMIT_WD = 5           # Submit deadline = visit_end + 5WD
FINAL_WD = 10           # Finalize deadline = visit_end + 10WD
FOLLOWUP_VISIT_WD = 10  # Follow-up option A = visit_end + 10WD
FOLLOWUP_FINAL_WD = 1   # Follow-up option B = final_date + 1WD

# Status values
FINALIZED = '\u5df2\u5b9a\u7a3f'
FOLLOWUP_SENT = 'SEND_SUCCESS'
ARCHIVED_VAL = '1'

# Risk thresholds
RISK_HIGH = 2   # Any dimension >= 2 -> high risk
RISK_ATTN = 1   # Any dimension >= 1 -> attention

# Table pagination
PER_PAGE = 18

# Chart colors (matplotlib)
CC = {
    'C1': '#0D7377', 'C2': '#5EA8A7', 'C3': '#F39C12',
    'C4': '#E76F51', 'C5': '#2D9E6B', 'C6': '#1A2332',
}

# PPT RGB color tuples
PC = {
    'primary':    (0x0D, 0x73, 0x77),
    'dark':       (0x1A, 0x23, 0x32),
    'danger':     (0xE7, 0x6F, 0x51),
    'success':    (0x2D, 0x9E, 0x6B),
    'light':      (0xF0, 0xF2, 0xF5),
    'white':      (0xFF, 0xFF, 0xFF),
    'text':       (0x2C, 0x3E, 0x50),
    'subtext':    (0x7F, 0x8C, 0x8D),
    'warning':    (0xF3, 0x9C, 0x12),
    'teal_light': (0x5E, 0xA8, 0xA7),
    'hr_bg':      (0xFF, 0xF5, 0xF5),
    'hr_text':    (0xC0, 0x39, 0x2B),
    'at_bg':      (0xFF, 0xFB, 0xF0),
    'at_text':    (0xF3, 0x9C, 0x12),
    'alt_row':    (0xF8, 0xF9, 0xFA),
    'dark_card':  (0x22, 0x2E, 0x3C),
}
