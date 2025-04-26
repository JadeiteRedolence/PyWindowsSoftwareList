import datetime
import pandas as pd
import tabulate  # This dependency was missing

def export_to_markdown(software_list, output_path):
    """Export software list to Markdown format"""
    df = pd.DataFrame(software_list)
    
    # Rename columns for better display
    df = df.rename(columns={
        "DisplayName": "Name",
        "DisplayVersion": "Version",
        "Publisher": "Publisher",
        "InstallDate": "Install Date",
        "InstallLocation": "Install Location"
    })
    
    # Generate Markdown
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write("# Installed Software List\n\n")
        f.write(f"Generated on: {datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}\n\n")
        f.write(df.to_markdown(index=False)) 