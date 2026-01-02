# Example Configuration File
# Copy this to config_personal.py and customize with your settings

CONFIG = {
    # ============================================
    # FILE PATHS
    # ============================================
    'excel_path': r'C:\Users\YourName\Documents\Reports\weekly_report.xlsx',
    'sheet_name': 'Report',          # Sheet containing the charts
    'comments_sheet': 'comments',    # Sheet containing commentary text
    
    # ============================================
    # EMAIL SETTINGS
    # ============================================
    'recipient': 'team@company.com',
    'subject_prefix': 'Weekly Report',
    'week_offset': 2,  # Report week = current week - offset (e.g., 2 = report on 2 weeks ago)
    
    # ============================================
    # EMAIL CONTENT
    # ============================================
    'greeting': 'All',
    'intro_text': 'Hereby the results for week {week}.',
    'closing_intro': 'For more detailed insights, please visit:',
    'closing_offer': 'If you have any questions or need additional information, feel free to reach out.',
    'signature': '''Best regards,<br>
Your Name<br>
Your Department<br>
Your Company''',
    
    # ============================================
    # CHART SETTINGS
    # ============================================
    'chart_order': [6, 3, 4, 5],  # Order in which charts appear (by Excel chart index)
    'chart_width': 900,            # Width of charts in pixels
    
    # ============================================
    # COMMENTS MAPPING
    # Maps chart index to Excel cell range
    # ============================================
    'comments_mapping': {
        6: 'A2:A6',    # Chart 6 gets commentary from cells A2-A6
        3: 'A8:A12',   # Chart 3 gets commentary from cells A8-A12
        4: 'A14:A18',  # Chart 4 gets commentary from cells A14-A18
    },
    
    # ============================================
    # EXTRA COMMENTS
    # Additional commentary sections after all charts
    # Format: (cell_range, add_space_after)
    # ============================================
    'extra_comments': [
        ('A20:A24', True),   # Commentary from A20-A24, with spacing after
        ('A26:A30', True),   # Commentary from A26-A30, with spacing after
        ('A32:A36', True),   # Commentary from A32-A36, with spacing after
        ('A38:A42', False),  # Commentary from A38-A42, no spacing after
    ],
    
    # ============================================
    # RESOURCE LINKS
    # Links to dashboards, reports, etc.
    # Format: (link_text, url)
    # ============================================
    'resource_links': [
        ('Weekly Dashboard', 'https://dashboard.company.com/weekly'),
        ('Sales Tracker', 'https://analytics.company.com/sales'),
        ('Competition Analysis', 'https://insights.company.com/competition'),
        ('Data Platform', 'https://data.company.com/reports'),
    ],
}

# ============================================
# USAGE
# ============================================
# 1. Copy this file to config_personal.py
# 2. Update all settings above
# 3. Import in main script:
#    from config_personal import CONFIG
#    chart_paths = export_working_charts(CONFIG['excel_path'], ...)
