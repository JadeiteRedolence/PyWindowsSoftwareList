import json
 
def export_to_json(software_list, output_path):
    """Export software list to JSON format"""
    with open(output_path, 'w', encoding='utf-8') as f:
        json.dump(software_list, f, indent=2, ensure_ascii=False) 