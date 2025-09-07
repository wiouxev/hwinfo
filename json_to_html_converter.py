#!/usr/bin/env python3
"""
System Report HTML Converter
Converts JSON system reports to beautiful HTML webpages

Usage:
    python report_converter.py [input.json] [output.html]
    
If no arguments provided, interactive mode will guide you through the process.
"""

import json
import sys
import os
import glob
from datetime import datetime
from pathlib import Path

def find_json_files():
    """Find JSON files in common locations"""
    common_paths = [
        ".",  # Current directory
        "C:\\temp",
        os.path.expanduser("~/Downloads"),
        os.path.expanduser("~/Desktop"),
        os.path.expanduser("~/Documents")
    ]
    
    json_files = []
    for path in common_paths:
        try:
            if os.path.exists(path):
                json_files.extend(glob.glob(os.path.join(path, "*.json")))
        except:
            continue
    
    # Remove duplicates and sort by modification time (newest first)
    json_files = list(set(json_files))
    json_files.sort(key=os.path.getmtime, reverse=True)
    
    return json_files

def get_input_file():
    """Interactive prompt for input file"""
    print("🔍 Looking for JSON files...")
    json_files = find_json_files()
    
    if json_files:
        print(f"\n📁 Found {len(json_files)} JSON files:")
        for i, file in enumerate(json_files[:10], 1):  # Show max 10 files
            try:
                size = os.path.getsize(file) / 1024  # Size in KB
                modified = datetime.fromtimestamp(os.path.getmtime(file)).strftime("%Y-%m-%d %H:%M")
                print(f"  {i}. {os.path.basename(file)} ({size:.1f} KB, {modified})")
                print(f"     📂 {os.path.dirname(file)}")
            except:
                print(f"  {i}. {os.path.basename(file)}")
                print(f"     📂 {os.path.dirname(file)}")
        
        if len(json_files) > 10:
            print(f"     ... and {len(json_files) - 10} more files")
        
        print(f"\n💡 You can:")
        print(f"   • Enter a number (1-{min(10, len(json_files))}) to select from the list above")
        print(f"   • Enter the full path to any JSON file")
        print(f"   • Drag and drop a file into this window")
        
        while True:
            user_input = input(f"\n📥 Select JSON file: ").strip().strip('"')
            
            # Check if it's a number selection
            if user_input.isdigit():
                selection = int(user_input)
                if 1 <= selection <= min(10, len(json_files)):
                    return json_files[selection - 1]
                else:
                    print(f"❌ Please enter a number between 1 and {min(10, len(json_files))}")
                    continue
            
            # Check if it's a file path
            if os.path.exists(user_input) and user_input.lower().endswith('.json'):
                return user_input
            
            print(f"❌ File not found or not a JSON file. Please try again.")
    
    else:
        print("📁 No JSON files found in common locations.")
        print("💡 Enter the full path to your JSON file:")
        
        while True:
            user_input = input("📥 JSON file path: ").strip().strip('"')
            if os.path.exists(user_input) and user_input.lower().endswith('.json'):
                return user_input
            print("❌ File not found or not a JSON file. Please try again.")

def get_output_file(input_file):
    """Interactive prompt for output file"""
    input_path = Path(input_file)
    suggested_name = f"{input_path.stem}_report.html"
    suggested_path = input_path.parent / suggested_name
    
    print(f"\n💾 Output HTML file:")
    print(f"📋 Suggested: {suggested_path}")
    print(f"💡 Press Enter to use suggested path, or enter custom path:")
    
    user_input = input(f"📤 Output file [Enter for default]: ").strip().strip('"')
    
    if not user_input:
        return str(suggested_path)
    
    # Ensure .html extension
    if not user_input.lower().endswith('.html'):
        user_input += '.html'
    
    return user_input

def format_property_value(value, category, key):
    """Format property values with special styling for certain types"""
    if category == 'Services':
        if value == 'Running':
            return f'<span class="status-running">✅ {value}</span>'
        else:
            return f'<span class="status-stopped">❌ {value}</span>'
    
    # Escape HTML characters
    if isinstance(value, str):
        value = value.replace('&', '&amp;').replace('<', '&lt;').replace('>', '&gt;')
    
    return str(value)

def generate_summary_section(data):
    """Generate a summary section with key metrics"""
    summary_items = []
    
    # Extract key information for summary
    try:
        if 'System' in data:
            if 'ComputerName' in data['System']:
                summary_items.append(f'<div class="summary-item"><h4>Computer</h4><div class="value">{data["System"]["ComputerName"]}</div></div>')
            if 'OSVersion' in data['System']:
                summary_items.append(f'<div class="summary-item"><h4>Operating System</h4><div class="value">{data["System"]["OSVersion"]}</div></div>')
        
        if 'Memory' in data:
            if 'TotalRAM' in data['Memory']:
                summary_items.append(f'<div class="summary-item"><h4>Total RAM</h4><div class="value">{data["Memory"]["TotalRAM"]}</div></div>')
        
        if 'CPU' in data:
            if 'Processor1' in data['CPU']:
                cpu_name = data['CPU']['Processor1'].split('(')[0].strip()  # Clean up CPU name
                summary_items.append(f'<div class="summary-item"><h4>Processor</h4><div class="value">{cpu_name}</div></div>')
        
        if 'Services' in data:
            running_services = sum(1 for service in data['Services'].values() if service == 'Running')
            total_services = len(data['Services'])
            summary_items.append(f'<div class="summary-item"><h4>Services Status</h4><div class="value">{running_services}/{total_services} Running</div></div>')
    
    except Exception as e:
        print(f"Warning: Could not generate summary - {e}")
    
    if summary_items:
        return f'''
        <div class="summary">
            <h3>📊 System Overview</h3>
            <div class="summary-grid">
                {"".join(summary_items)}
            </div>
        </div>'''
    return ""

def generate_sections(data):
    """Generate all the expandable sections"""
    sections_html = []
    
    # Define section icons
    section_icons = {
        'System': '🖥️',
        'Network': '🌐',
        'CPU': '⚡',
        'Memory': '🧠',
        'Disk': '💾',
        'Graphics': '🎮',
        'Software': '📦',
        'Services': '⚙️'
    }
    
    for category, properties in data.items():
        icon = section_icons.get(category, '📋')
        section_html = f'''
        <div class="section">
            <div class="section-header" onclick="toggleSection(this)">
                <span>{icon} {category} Information</span>
                <span class="toggle-icon">▼</span>
            </div>
            <div class="section-content">
        '''
        
        for key, value in properties.items():
            # Clean up property names
            clean_key = key.replace('1', '').replace('2', '').replace('3', '').replace('4', '').replace('5', '')
            clean_key = clean_key.replace('_', ' ').strip()
            
            # Format values based on category and content
            formatted_value = format_property_value(value, category, key)
            
            section_html += f'''
                <div class="property">
                    <div class="property-name">{clean_key}:</div>
                    <div class="property-value">{formatted_value}</div>
                </div>
            '''
        
        section_html += '''
            </div>
        </div>
        '''
        sections_html.append(section_html)
    
    return ''.join(sections_html)

def generate_html_template(data, computer_name, timestamp):
    """Generate the complete HTML report"""
    
    html_template = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>System Report - {computer_name}</title>
    <style>
        * {{
            margin: 0;
            padding: 0;
            box-sizing: border-box;
        }}
        
        body {{
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            line-height: 1.6;
            color: #333;
            background: linear-gradient(135deg, #667eea 0%, #764ba2 100%);
            min-height: 100vh;
            padding: 20px;
        }}
        
        .container {{
            max-width: 1200px;
            margin: 0 auto;
            background: #ffffff;
            border-radius: 15px;
            box-shadow: 0 20px 40px rgba(0,0,0,0.1);
            overflow: hidden;
        }}
        
        .header {{
            background: linear-gradient(135deg, #2c3e50 0%, #34495e 100%);
            color: white;
            padding: 40px 30px;
            text-align: center;
        }}
        
        .header h1 {{
            font-size: 2.5em;
            margin-bottom: 10px;
            font-weight: 300;
        }}
        
        .header .subtitle {{
            font-size: 1.1em;
            opacity: 0.9;
            margin-bottom: 20px;
        }}
        
        .header .timestamp {{
            background: rgba(255,255,255,0.1);
            padding: 8px 16px;
            border-radius: 20px;
            display: inline-block;
            font-size: 0.9em;
        }}
        
        .content {{
            padding: 30px;
        }}
        
        .section {{
            margin-bottom: 25px;
            border-radius: 10px;
            overflow: hidden;
            box-shadow: 0 4px 6px rgba(0,0,0,0.05);
            border: 1px solid #e1e8ed;
        }}
        
        .section-header {{
            background: linear-gradient(135deg, #3498db 0%, #2980b9 100%);
            color: white;
            padding: 20px 25px;
            cursor: pointer;
            font-size: 1.2em;
            font-weight: 600;
            display: flex;
            justify-content: space-between;
            align-items: center;
            transition: all 0.3s ease;
        }}
        
        .section-header:hover {{
            background: linear-gradient(135deg, #2980b9 0%, #3498db 100%);
            transform: translateY(-1px);
        }}
        
        .section-header.active {{
            background: linear-gradient(135deg, #27ae60 0%, #2ecc71 100%);
        }}
        
        .toggle-icon {{
            font-size: 1.2em;
            transition: transform 0.3s ease;
        }}
        
        .toggle-icon.rotated {{
            transform: rotate(180deg);
        }}
        
        .section-content {{
            padding: 0;
            max-height: 0;
            overflow: hidden;
            transition: all 0.4s ease;
            background: #f8f9fa;
        }}
        
        .section-content.expanded {{
            padding: 25px;
            max-height: 2000px;
        }}
        
        .property {{
            display: flex;
            margin-bottom: 15px;
            padding: 12px 16px;
            background: white;
            border-radius: 8px;
            border-left: 4px solid #3498db;
            transition: all 0.2s ease;
        }}
        
        .property:hover {{
            box-shadow: 0 2px 8px rgba(0,0,0,0.1);
            transform: translateX(2px);
        }}
        
        .property-name {{
            font-weight: 600;
            min-width: 200px;
            color: #2c3e50;
            margin-right: 20px;
        }}
        
        .property-value {{
            color: #34495e;
            flex: 1;
            word-break: break-word;
        }}
        
        .status-running {{
            color: #27ae60;
            font-weight: bold;
            padding: 4px 8px;
            background: #d5f4e6;
            border-radius: 12px;
            font-size: 0.9em;
        }}
        
        .status-stopped {{
            color: #e74c3c;
            font-weight: bold;
            padding: 4px 8px;
            background: #fdeaea;
            border-radius: 12px;
            font-size: 0.9em;
        }}
        
        .summary {{
            background: linear-gradient(135deg, #f39c12 0%, #e67e22 100%);
            color: white;
            padding: 20px;
            margin-bottom: 30px;
            border-radius: 10px;
            text-align: center;
        }}
        
        .summary-grid {{
            display: grid;
            grid-template-columns: repeat(auto-fit, minmax(200px, 1fr));
            gap: 20px;
            margin-top: 15px;
        }}
        
        .summary-item {{
            background: rgba(255,255,255,0.1);
            padding: 15px;
            border-radius: 8px;
        }}
        
        .summary-item h4 {{
            margin-bottom: 5px;
            font-size: 0.9em;
            opacity: 0.9;
        }}
        
        .summary-item .value {{
            font-size: 1.2em;
            font-weight: bold;
        }}
        
        .footer {{
            text-align: center;
            padding: 30px;
            background: #f8f9fa;
            color: #7f8c8d;
            border-top: 1px solid #e1e8ed;
        }}
        
        .expand-all {{
            background: #3498db;
            color: white;
            border: none;
            padding: 12px 24px;
            border-radius: 25px;
            cursor: pointer;
            font-size: 1em;
            margin-bottom: 20px;
            transition: all 0.3s ease;
        }}
        
        .expand-all:hover {{
            background: #2980b9;
            transform: translateY(-2px);
            box-shadow: 0 4px 8px rgba(0,0,0,0.2);
        }}
        
        @media (max-width: 768px) {{
            .property {{
                flex-direction: column;
            }}
            
            .property-name {{
                min-width: auto;
                margin-bottom: 8px;
                margin-right: 0;
            }}
            
            .header h1 {{
                font-size: 2em;
            }}
            
            .content {{
                padding: 20px;
            }}
        }}
    </style>
</head>
<body>
    <div class="container">
        <div class="header">
            <h1>🖥️ System Information Report</h1>
            <p class="subtitle">Computer: <strong>{computer_name}</strong></p>
            <div class="timestamp">Generated: {timestamp}</div>
        </div>
        
        <div class="content">
            <button class="expand-all" onclick="toggleAllSections()">📂 Expand All Sections</button>
            
            {generate_summary_section(data)}
            
            {generate_sections(data)}
        </div>
        
        <div class="footer">
            <p>Report generated by PowerShell System Information Script</p>
            <p>Converted to HTML by Python Report Converter</p>
        </div>
    </div>
    
    <script>
        let allExpanded = false;
        
        function toggleSection(header) {{
            const content = header.nextElementSibling;
            const icon = header.querySelector('.toggle-icon');
            
            header.classList.toggle('active');
            content.classList.toggle('expanded');
            icon.classList.toggle('rotated');
        }}
        
        function toggleAllSections() {{
            const headers = document.querySelectorAll('.section-header');
            const button = document.querySelector('.expand-all');
            
            headers.forEach(header => {{
                const content = header.nextElementSibling;
                const icon = header.querySelector('.toggle-icon');
                
                if (!allExpanded) {{
                    header.classList.add('active');
                    content.classList.add('expanded');
                    icon.classList.add('rotated');
                }} else {{
                    header.classList.remove('active');
                    content.classList.remove('expanded');
                    icon.classList.remove('rotated');
                }}
            }});
            
            allExpanded = !allExpanded;
            button.textContent = allExpanded ? '📁 Collapse All Sections' : '📂 Expand All Sections';
        }}
        
        // Auto-expand System section
        document.addEventListener('DOMContentLoaded', function() {{
            const firstSection = document.querySelector('.section-header');
            if (firstSection) {{
                toggleSection(firstSection);
            }}
        }});
    </script>
</body>
</html>"""
    
    return html_template

def convert_json_to_html(json_file_path, output_file_path=None):
    """Main conversion function"""
    try:
        # Read JSON file (handle UTF-8 BOM from PowerShell)
        with open(json_file_path, 'r', encoding='utf-8-sig') as f:
            data = json.load(f)
        
        # Extract computer name and timestamp
        computer_name = "Unknown Computer"
        timestamp = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        
        # Try to get computer name from the data
        if 'System' in data and 'ComputerName' in data['System']:
            computer_name = data['System']['ComputerName']
        
        # Generate output filename if not provided
        if output_file_path is None:
            json_path = Path(json_file_path)
            output_file_path = json_path.parent / f"{json_path.stem}_report.html"
        
        # Generate HTML
        html_content = generate_html_template(data, computer_name, timestamp)
        
        # Write HTML file
        with open(output_file_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        
        print(f"✅ Successfully converted JSON to HTML!")
        print(f"📁 Input:  {json_file_path}")
        print(f"🌐 Output: {output_file_path}")
        print(f"📊 Data sections: {len(data)}")
        
        return str(output_file_path)
        
    except FileNotFoundError:
        print(f"❌ Error: Could not find JSON file: {json_file_path}")
        return None
    except json.JSONDecodeError as e:
        print(f"❌ Error: Invalid JSON file - {e}")
        return None
    except Exception as e:
        print(f"❌ Error: {e}")
        return None

def main():
    """Command line interface with interactive prompts"""
    print("🎨 System Report HTML Converter")
    print("=" * 40)
    
    # Check if arguments were provided (command line mode)
    if len(sys.argv) >= 2:
        print("📋 Command line mode detected")
        json_file = sys.argv[1]
        output_file = sys.argv[2] if len(sys.argv) > 2 else None
        
        if not os.path.exists(json_file):
            print(f"❌ Error: File '{json_file}' does not exist!")
            sys.exit(1)
    
    else:
        # Interactive mode
        print("🤝 Interactive mode - I'll guide you through the process!\n")
        
        try:
            json_file = get_input_file()
            output_file = get_output_file(json_file)
        except KeyboardInterrupt:
            print(f"\n\n👋 Conversion cancelled by user. Goodbye!")
            sys.exit(0)
    
    print(f"\n🔄 Converting...")
    print(f"📥 Input:  {json_file}")
    print(f"📤 Output: {output_file}")
    
    result = convert_json_to_html(json_file, output_file)
    
    if result:
        print(f"\n🎉 Conversion completed successfully!")
        
        # Ask if user wants to open the file
        try:
            open_file = input(f"\n🌐 Would you like to open the HTML file now? (y/n): ").strip().lower()
            if open_file in ['y', 'yes', '']:
                try:
                    os.startfile(result)  # Windows
                    print(f"🚀 Opening {result} in your default browser...")
                except:
                    print(f"💡 Please open this file manually: {result}")
        except KeyboardInterrupt:
            print(f"\n👋 Done!")
        
        sys.exit(0)
    else:
        sys.exit(1)

if __name__ == "__main__":
    main()
