import pandas as pd

def export_to_excel(software_list, output_path):
    """Export software list to Excel format"""
    df = pd.DataFrame(software_list)
    
    # Rename columns for better display
    df = df.rename(columns={
        "DisplayName": "Name",
        "DisplayVersion": "Version",
        "Publisher": "Publisher",
        "InstallDate": "Install Date",
        "InstallLocation": "Install Location"
    })
    
    # Create Excel writer
    writer = pd.ExcelWriter(output_path, engine='openpyxl')
    
    # Convert dataframe to Excel
    df.to_excel(writer, sheet_name='Installed Software', index=False)
    
    # Adjust column width
    worksheet = writer.sheets['Installed Software']
    for idx, col in enumerate(df.columns):
        max_len = max(
            df[col].astype(str).map(len).max(),
            len(col)
        ) + 2
        # Excel column index starts at 1
        worksheet.column_dimensions[chr(65 + idx)].width = max_len
    
    # Save Excel file
    writer.close() 