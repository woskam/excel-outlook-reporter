import win32com.client as win32
from datetime import datetime
import os

def export_working_charts(excel_path, sheet_name, output_folder):
    """Export only charts larger than 1 KB"""
    
    if not os.path.exists(output_folder):
        os.makedirs(output_folder)
    
    excel = None
    chart_paths = {}  # Dictionary to store chart paths with their index
    
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        
        workbook = excel.Workbooks.Open(excel_path)
        sheet = workbook.Sheets(sheet_name)
        
        chart_count = sheet.ChartObjects().Count
        print(f"Total {chart_count} charts found, exporting...\n")
        
        for i in range(1, chart_count + 1):
            chart = sheet.ChartObjects(i)
            chart_name = chart.Name
            output_path = os.path.join(output_folder, f'chart_{i}_{chart_name}.png')
            
            # Export chart
            chart.Chart.Export(output_path)
            
            # Check if export was successful (larger than 1 KB)
            file_size = os.path.getsize(output_path) / 1024  # KB
            
            if file_size > 1.0:
                chart_paths[i] = output_path  # Save with original index
                print(f"  ✓ Chart {i} ({chart_name}): {file_size:.1f} KB - EXPORTED")
            else:
                os.remove(output_path)  # Remove empty chart
                print(f"  ✗ Chart {i} ({chart_name}): {file_size:.1f} KB - SKIPPED (empty)")
        
        workbook.Close(False)
        print(f"\n✓ {len(chart_paths)} working charts exported")
        return chart_paths
        
    except Exception as e:
        print(f"Error during export: {e}")
        raise
    finally:
        if excel:
            excel.Quit()

def read_comments_from_excel(excel_path, sheet_name, cell_range):
    """Read text from a specific cell range"""
    
    excel = None
    try:
        excel = win32.Dispatch('Excel.Application')
        excel.Visible = False
        
        workbook = excel.Workbooks.Open(excel_path)
        sheet = workbook.Sheets(sheet_name)
        
        # Read the range
        range_obj = sheet.Range(cell_range)
        
        # Combine all cells in the range into one text
        comment_lines = []
        for cell in range_obj:
            if cell.Value is not None and str(cell.Value).strip():
                comment_lines.append(str(cell.Value).strip())
        
        workbook.Close(False)
        
        # Combine all lines with line breaks
        return '<br>'.join(comment_lines) if comment_lines else ''
        
    except Exception as e:
        print(f"Error reading comments: {e}")
        return ''
    finally:
        if excel:
            excel.Quit()

def create_weekly_email(chart_paths, excel_path, config):
    """Create email with charts in custom order and comments in between"""
    
    outlook = win32.Dispatch('outlook.application')
    mail = outlook.CreateItem(0)
    
    # Week number - week offset
    current_week = datetime.now().isocalendar()[1]
    report_week = current_week - config['week_offset']
    year = datetime.now().year
    
    mail.To = config['recipient']
    mail.Subject = f"{config['subject_prefix']} - Week {report_week} {year}"
    
    html_body = f"""
    <html>
    <head>
        <style>
            body {{
                font-family: Calibri, Arial, sans-serif;
                color: #333;
                line-height: 1.6;
            }}
            h2 {{
                color: #0078D4;
                border-bottom: 2px solid #0078D4;
                padding-bottom: 5px;
            }}
            .chart-container {{
                margin: 30px 0;
                text-align: left;
            }}
            .chart-container img {{
                max-width: 100%;
                height: auto;
            }}
            .comment-section {{
                margin: 20px 0;
                padding: 15px;
                background-color: #f8f9fa;
                border-left: 4px solid #0078D4;
                text-align: left;
            }}
            .closing {{
                margin-top: 40px;
                padding-top: 20px;
                border-top: 1px solid #ddd;
            }}
            .closing p {{
                margin: 10px 0;
            }}
            a {{
                color: #0078D4;
                text-decoration: none;
            }}
            a:hover {{
                text-decoration: underline;
            }}
            .footer {{
                margin-top: 30px;
                font-size: 10px;
                color: #999;
            }}
        </style>
    </head>
    <body>
        <p>{config['greeting']},</p>
        <p>{config['intro_text'].format(week=report_week)}</p>
    """
    
    # Custom order for charts
    chart_order = config['chart_order']
    
    # Comment ranges linked to charts
    comments_mapping = config['comments_mapping']
    
    # Add charts with comments in between
    cid_counter = 1
    
    for chart_index in chart_order:
        if chart_index in chart_paths:
            # Read comment for this chart (if exists)
            if chart_index in comments_mapping:
                comment_range = comments_mapping[chart_index]
                comment_text = read_comments_from_excel(excel_path, config['comments_sheet'], comment_range)
                
                if comment_text:
                    html_body += f"""
                    <div class="comment-section">
                        {comment_text}
                    </div>
                    """
            
            # Add chart
            chart_path = chart_paths[chart_index]
            attachment = mail.Attachments.Add(chart_path)
            cid = f"chart{cid_counter}"
            attachment.PropertyAccessor.SetProperty(
                "http://schemas.microsoft.com/mapi/proptag/0x3712001F", 
                cid
            )
            
            html_body += f"""
            <div class="chart-container">
                <img src="cid:{cid}" width="{config['chart_width']}">
            </div>
            """
            
            cid_counter += 1
    
    # Extra comments at the end
    extra_comments = config.get('extra_comments', [])
    
    for comment_range, add_space in extra_comments:
        comment_text = read_comments_from_excel(excel_path, config['comments_sheet'], comment_range)
        
        if comment_text:
            html_body += f"""
            <div class="comment-section">
                {comment_text}
            </div>
            """
            
            if add_space:
                html_body += "<br>"
    
    # CLOSING
    html_body += f"""
        <div class="closing">
            <p>{config['closing_intro']}</p>
            <ul>
    """
    
    # Add resource links
    for link_text, link_url in config['resource_links']:
        html_body += f'                <li><a href="{link_url}">{link_text}</a></li>\n'
    
    html_body += f"""
            </ul>
            
            <p>{config['closing_offer']}</p>
            
            <p>{config['signature']}</p>
        </div>
        
        <div class="footer">
            <p>-----------------------</p>
        </div>
    </body>
    </html>
    """
    
    mail.HTMLBody = html_body
    return mail

def main():
    print(f"\n{'='*60}")
    print(f"Weekly Report Generator - {datetime.now().strftime('%d-%m-%Y %H:%M:%S')}")
    print(f"{'='*60}\n")
    
    # ============================================
    # CONFIGURATION
    # ============================================
    config = {
        # File paths
        'excel_path': r'C:\path\to\your\report.xlsx',
        'sheet_name': 'Report',
        'comments_sheet': 'comments',
        
        # Email settings
        'recipient': 'team@example.com',
        'subject_prefix': 'Weekly Report',
        'week_offset': 2,  # Report on week N-2
        
        # Email content
        'greeting': 'All',
        'intro_text': 'Hereby the results for week {week}.',
        'closing_intro': 'For more detailed insights, please visit:',
        'closing_offer': 'If you have any questions or need additional information, feel free to reach out.',
        'signature': 'Best regards,<br>Your Name<br>Your Department',
        
        # Chart settings
        'chart_order': [6, 3, 4, 5],  # Custom order for charts
        'chart_width': 900,
        
        # Comments mapping (chart index -> Excel range)
        'comments_mapping': {
            6: 'A2:A6',   # For Chart 6
            3: 'A8:A12',  # For Chart 3
            4: 'A14:A18', # For Chart 4
        },
        
        # Extra comments at the end (range, add_space_after)
        'extra_comments': [
            ('A20:A24', True),
            ('A26:A30', True),
            ('A32:A36', True),
            ('A38:A42', False),
        ],
        
        # Resource links (text, url)
        'resource_links': [
            ('WEEKLY REPORTING', 'https://your-reporting-url.com'),
            ('SALES TRACKER', 'https://your-tracker-url.com'),
            ('COMPETITION DASHBOARD', 'https://your-dashboard-url.com'),
            ('DATA SCRAPER', 'https://your-scraper-url.com'),
        ],
    }
    # ============================================
    
    temp_folder = os.path.join(os.environ['TEMP'], 'weekly_charts_temp')
    
    try:
        # Export only working charts
        chart_paths = export_working_charts(
            config['excel_path'], 
            config['sheet_name'], 
            temp_folder
        )
        
        if not chart_paths:
            print("\n❌ No working charts found!")
            return
        
        # Create email
        print("\nCreating email...")
        print(f"Reading comments from '{config['comments_sheet']}' sheet...")
        mail = create_weekly_email(chart_paths, config['excel_path'], config)
        
        # Show week info
        current_week = datetime.now().isocalendar()[1]
        report_week = current_week - config['week_offset']
        print(f"\n✓ Email for Week {report_week} (current week: {current_week})")
        print(f"✓ {len(chart_paths)} charts added")
        print(f"✓ Recipient: {config['recipient']}")
        
        # Open email as draft
        mail.Display()
        print("\n✓ Email opened for review")
        
        print(f"\n{'='*60}")
        print("SUCCESS!")
        print(f"{'='*60}\n")
        
    except Exception as e:
        print(f"\n❌ ERROR: {e}")
        raise

if __name__ == "__main__":
    main()
    input("\nPress Enter to close...")
