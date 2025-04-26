import datetime
from pathlib import Path

# Output directory configuration
def get_output_directory():
    """Get the output directory with timestamp"""
    timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
    output_dir = Path("Report") / timestamp
    return output_dir

# Output file names
OUTPUT_FILENAMES = {
    "json": "software_list.json",
    "html": "software_list.html",
    "markdown": "software_list.md",
    "excel": "software_list.xlsx"
}

# Column display names for better readability
COLUMN_DISPLAY_NAMES = {
    "DisplayName": "Name",
    "DisplayVersion": "Version",
    "Publisher": "Publisher",
    "InstallDate": "Install Date",
    "InstallLocation": "Install Location"
} 