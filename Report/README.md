# Report Directory

This directory contains all generated system information reports. Each run of the tool creates a new timestamped subdirectory with comprehensive system data.

## Directory Structure

```
Report/YYYYMMDD_HHMMSS/
├── excel/                    # Excel spreadsheets
│   └── software_list.xlsx   # Installed software in spreadsheet format
├── html/                     # Individual HTML reports  
│   └── software_list.html   # Software list in HTML table format
├── json/                     # Structured data files (raw data)
│   ├── complete_report_data.json  # All collected data combined
│   ├── system_info.json           # Detailed system specifications
│   ├── software_list.json         # Software inventory data
│   └── dev_environment.json       # Development environment info
├── logs/                     # Debug and execution logs
│   └── debug_YYYYMMDD_HHMMSS.log  # Detailed execution log
├── markdown/                 # Markdown documentation
│   └── software_list.md     # Software list in Markdown format
├── reports/                  # Comprehensive reports
│   └── system_report.html   # **Main comprehensive report**
└── index.html               # Navigation index linking to all reports
```

## Main Files

- **`reports/system_report.html`** - The main comprehensive report with tabbed interface
- **`index.html`** - Navigation page linking to all generated reports
- **`json/complete_report_data.json`** - Complete raw data for further processing
- **`logs/debug_*.log`** - Execution logs for troubleshooting

## Data Privacy

⚠️ **Important**: The generated reports may contain sensitive system information including:
- Installed software lists
- System hardware specifications
- Network configuration details
- Environment variables and paths
- User account information

**Do not share these reports publicly or store them in unsecured locations.**

## File Formats

- **HTML**: Interactive reports with search and filtering capabilities
- **JSON**: Machine-readable structured data for automation
- **Excel**: Spreadsheet format for data analysis
- **Markdown**: Human-readable documentation format
- **Logs**: Detailed execution information for debugging 