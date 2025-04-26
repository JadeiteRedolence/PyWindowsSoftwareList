from .json_exporter import export_to_json
from .html_exporter import export_to_html
from .markdown_exporter import export_to_markdown
from .excel_exporter import export_to_excel

def export_all_formats(software_list, output_dir):
    """Export software list to all supported formats"""
    export_to_json(software_list, output_dir / "software_list.json")
    export_to_html(software_list, output_dir / "software_list.html")
    
    try:
        export_to_markdown(software_list, output_dir / "software_list.md")
    except ImportError:
        print("Warning: Missing 'tabulate' package. Markdown export skipped.")
    
    export_to_excel(software_list, output_dir / "software_list.xlsx") 