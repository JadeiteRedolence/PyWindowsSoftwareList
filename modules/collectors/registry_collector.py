import winreg

def get_software_from_registry():
    """Get installed software list from Windows registry"""
    software_list = []
    
    # Registry paths for installed software
    reg_paths = [
        r"SOFTWARE\Microsoft\Windows\CurrentVersion\Uninstall",
        r"SOFTWARE\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"
    ]
    
    # Registry keys to retrieve
    keys_to_get = ["DisplayName", "DisplayVersion", "Publisher", "InstallDate", "InstallLocation"]
    
    for reg_path in reg_paths:
        try:
            registry_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, reg_path)
            
            # Enumerate all subkeys
            for i in range(0, winreg.QueryInfoKey(registry_key)[0]):
                subkey_name = winreg.EnumKey(registry_key, i)
                subkey_path = f"{reg_path}\\{subkey_name}"
                
                try:
                    software_key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, subkey_path)
                    software_info = {}
                    
                    # Get the values we're interested in
                    for key in keys_to_get:
                        try:
                            value, _ = winreg.QueryValueEx(software_key, key)
                            software_info[key] = value
                        except FileNotFoundError:
                            software_info[key] = ""
                    
                    # Only add if it has a display name (valid software entry)
                    if "DisplayName" in software_info and software_info["DisplayName"]:
                        software_list.append(software_info)
                    
                    winreg.CloseKey(software_key)
                except (WindowsError, FileNotFoundError):
                    continue
                
            winreg.CloseKey(registry_key)
        except (WindowsError, FileNotFoundError):
            continue
    
    return software_list 