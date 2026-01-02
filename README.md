# Weekly Email Report Automation

Automate the creation of weekly report emails with Excel charts and comments using Python and Outlook.

## Features

- 📊 **Automated Chart Export**: Extracts charts from Excel workbooks and filters out empty ones
- 📧 **Email Generation**: Creates professionally formatted HTML emails with embedded charts
- 💬 **Dynamic Comments**: Pulls commentary from Excel cells and inserts between charts
- ⚙️ **Highly Configurable**: Customize chart order, email content, and styling
- 🔄 **Week Offset Support**: Automatically calculates reporting week with configurable offset

## Requirements

- Windows OS (required for `win32com`)
- Python 3.7+
- Microsoft Excel installed
- Microsoft Outlook installed
- Python packages:
  - `pywin32`

## Installation

1. Clone this repository:
```bash
git clone https://github.com/yourusername/weekly-email-automation.git
cd weekly-email-automation
```

2. Install required packages:
```bash
pip install pywin32
```

## Configuration

Edit the `config` dictionary in the `main()` function to match your setup:

### File Paths
```python
'excel_path': r'C:\path\to\your\report.xlsx',
'sheet_name': 'Report',           # Sheet containing charts
'comments_sheet': 'comments',     # Sheet containing comment text
```

### Email Settings
```python
'recipient': 'team@example.com',
'subject_prefix': 'Weekly Report',
'week_offset': 2,  # Report on week N-2 (e.g., if current week is 5, report week 3)
```

### Content Customization
```python
'greeting': 'All',
'intro_text': 'Hereby the results for week {week}.',
'signature': 'Best regards,<br>Your Name<br>Your Department',
```

### Chart Configuration
```python
'chart_order': [6, 3, 4, 5],  # Specify exact order of charts by their index
'chart_width': 900,           # Width of charts in email (pixels)
```

### Comments Mapping
Link Excel cell ranges to specific charts:
```python
'comments_mapping': {
    6: 'A2:A6',    # Chart 6 gets comments from cells A2-A6
    3: 'A8:A12',   # Chart 3 gets comments from cells A8-A12
    4: 'A14:A18',  # Chart 4 gets comments from cells A14-A18
},
```

### Resource Links
Add links to dashboards or reports:
```python
'resource_links': [
    ('Dashboard Name', 'https://your-url.com'),
    ('Another Resource', 'https://another-url.com'),
],
```

## Excel Workbook Structure

Your Excel file should have:

1. **Report Sheet** (or your configured `sheet_name`):
   - Contains the charts you want to export
   - Charts are numbered automatically (1, 2, 3, etc.)

2. **Comments Sheet** (or your configured `comments_sheet`):
   - Contains text commentary in cell ranges
   - Each range corresponds to a chart via `comments_mapping`
   - Text is automatically formatted with `<br>` tags for HTML

## Usage

Run the script:
```bash
python weekly_email_automation.py
```

The script will:
1. Export all charts from Excel (only those > 1KB to skip empty ones)
2. Read comments from the specified cell ranges
3. Generate an HTML email with charts and comments in your specified order
4. Open the email in Outlook as a draft for review
5. You can then review and send manually

## How It Works

### Chart Export
- Opens Excel workbook invisibly
- Exports each chart as PNG
- Checks file size (must be > 1KB)
- Skips charts that are empty/corrupt

### Comment Extraction
- Reads specified cell ranges from comments sheet
- Combines multi-cell ranges with line breaks
- Inserts as styled HTML sections

### Email Creation
- Creates Outlook email draft
- Embeds charts as inline images (not attachments)
- Applies professional HTML/CSS styling
- Calculates correct reporting week based on offset

## Customization Examples

### Change Chart Order
```python
'chart_order': [1, 5, 2, 3, 4],  # Charts appear in this exact order
```

### Add More Comment Sections
```python
'extra_comments': [
    ('A20:A24', True),   # Range with spacing after
    ('A26:A30', True),
    ('A32:A36', False),  # No spacing after
],
```

### Modify Email Styling
Edit the CSS in the `create_weekly_email()` function to change:
- Colors (change `#0078D4` hex codes)
- Fonts (change `font-family`)
- Spacing (adjust margins/padding)
- Comment box styling

## Troubleshooting

**Charts not appearing:**
- Ensure chart names don't have special characters
- Check that charts exist on the specified sheet
- Verify Excel file path is correct

**Comments not showing:**
- Check cell range syntax (e.g., 'A2:A6')
- Ensure comments sheet name matches config
- Verify cells contain text (not formulas returning empty)

**Email not opening:**
- Ensure Outlook is installed and configured
- Check that you're logged into an Outlook account
- Verify recipient email format

**"No working charts found":**
- All charts were < 1KB (likely empty)
- Check the Report sheet has visible charts
- Try opening Excel file manually to verify charts render

## Security Note

This script:
- Opens files locally on your machine
- Does not send emails automatically (manual review required)
- Does not store or transmit credentials
- Requires Excel and Outlook to be installed and configured

## License

MIT License - feel free to modify and use for your purposes.

## Contributing

Contributions are welcome! Please feel free to submit a Pull Request.

## Acknowledgments

Built with:
- [pywin32](https://github.com/mhammond/pywin32) - Windows COM automation
- Microsoft Excel & Outlook automation APIs
