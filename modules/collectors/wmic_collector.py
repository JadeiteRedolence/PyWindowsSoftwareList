import subprocess

def get_software_from_wmic():
    """Get installed software using WMIC command"""
    try:
        result = subprocess.run(
            ["wmic", "product", "get", "Name,Version,Vendor,InstallDate"],
            capture_output=True, text=True, check=True
        )
        
        lines = result.stdout.strip().split('\n')
        if len(lines) <= 1:  # Only header or empty
            return []
        
        # Parse header
        header = lines[0].strip().split()
        
        software_list = []
        for line in lines[1:]:
            if line.strip():
                parts = line.strip().split('  ')
                parts = [p for p in parts if p]
                
                if len(parts) >= 3:
                    software_info = {
                        "DisplayName": parts[0],
                        "DisplayVersion": parts[1] if len(parts) > 1 else "",
                        "Publisher": parts[2] if len(parts) > 2 else "",
                        "InstallDate": parts[3] if len(parts) > 3 else "",
                        "InstallLocation": ""
                    }
                    software_list.append(software_info)
        
        return software_list
    except Exception as e:
        print(f"Error getting software via WMIC: {e}")
        return [] 