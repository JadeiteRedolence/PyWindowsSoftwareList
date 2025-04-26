import os
import subprocess
import json
import winreg
from datetime import datetime

def get_installed_software_from_registry(registry_key, flag=0):
    """
    从注册表获取已安装软件列表
    
    参数:
    - registry_key: 注册表路径
    - flag: 32位或64位，0=默认，KEY_WOW64_32KEY=0x0200, KEY_WOW64_64KEY=0x0100
    
    返回:
    - 软件列表，每个软件包含名称、版本、发布者、安装日期等信息
    """
    software_list = []
    
    try:
        # 打开注册表
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, registry_key, 0, winreg.KEY_READ | flag) as key:
            # 计算子项的数量
            subkey_count = winreg.QueryInfoKey(key)[0]
            
            # 遍历子项
            for i in range(subkey_count):
                try:
                    # 获取子项名称
                    software_name = winreg.EnumKey(key, i)
                    # 打开子项
                    with winreg.OpenKey(key, software_name) as subkey:
                        try:
                            # 获取软件信息
                            display_name = winreg.QueryValueEx(subkey, "DisplayName")[0]
                            
                            # 获取其他属性（如果存在）
                            software_info = {
                                "name": display_name,
                                "registry_path": f"{registry_key}\\{software_name}"
                            }
                            
                            # 尝试获取版本
                            try:
                                software_info["version"] = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                            except:
                                pass
                                
                            # 尝试获取发布者
                            try:
                                software_info["publisher"] = winreg.QueryValueEx(subkey, "Publisher")[0]
                            except:
                                pass
                                
                            # 尝试获取安装位置
                            try:
                                software_info["install_location"] = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                            except:
                                pass
                                
                            # 尝试获取卸载字符串
                            try:
                                software_info["uninstall_string"] = winreg.QueryValueEx(subkey, "UninstallString")[0]
                            except:
                                pass
                                
                            # 尝试获取安装日期
                            try:
                                install_date = winreg.QueryValueEx(subkey, "InstallDate")[0]
                                if isinstance(install_date, str) and len(install_date) == 8:
                                    # 将YYYYMMDD格式转换为标准日期格式
                                    year = install_date[0:4]
                                    month = install_date[4:6]
                                    day = install_date[6:8]
                                    software_info["install_date"] = f"{year}-{month}-{day}"
                            except:
                                pass
                                
                            # 将信息添加到列表
                            software_list.append(software_info)
                                
                        except (FileNotFoundError, WindowsError):
                            # 如果无法获取DisplayName，则跳过
                            continue
                except (FileNotFoundError, WindowsError):
                    continue
    except (FileNotFoundError, WindowsError) as e:
        print(f"Error accessing registry key {registry_key}: {e}")
        
    return software_list

def get_installed_software_from_powershell():
    """
    使用PowerShell获取已安装软件列表
    
    返回:
    - 从PowerShell获取的软件列表
    """
    software_list = []
    
    try:
        # 使用PowerShell命令获取已安装软件
        ps_command = "Get-ItemProperty HKLM:\\Software\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\*, HKLM:\\Software\\Microsoft\\Windows\\CurrentVersion\\Uninstall\\* | Where-Object { $_.DisplayName -ne $null } | Select-Object DisplayName, DisplayVersion, Publisher, InstallDate | ConvertTo-Json"
        
        # 执行PowerShell命令
        result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            # 解析JSON结果
            try:
                ps_result = json.loads(result.stdout)
                
                # 确保结果是列表（有时只有一个项目时会返回单个对象）
                if not isinstance(ps_result, list):
                    ps_result = [ps_result]
                
                # 处理结果
                for item in ps_result:
                    software_info = {
                        "name": item.get("DisplayName", ""),
                        "version": item.get("DisplayVersion", ""),
                        "publisher": item.get("Publisher", ""),
                        "source": "powershell"
                    }
                    
                    # 处理安装日期（如果存在）
                    install_date = item.get("InstallDate")
                    if install_date and isinstance(install_date, str) and len(install_date) == 8:
                        year = install_date[0:4]
                        month = install_date[4:6]
                        day = install_date[6:8]
                        software_info["install_date"] = f"{year}-{month}-{day}"
                    
                    software_list.append(software_info)
            except json.JSONDecodeError as e:
                print(f"Error parsing PowerShell output: {e}")
        else:
            print(f"PowerShell command failed: {result.stderr}")
    except Exception as e:
        print(f"Error running PowerShell command: {e}")
    
    return software_list

def get_uwp_apps():
    """
    获取已安装的UWP应用列表
    
    返回:
    - UWP应用列表
    """
    uwp_apps = []
    
    try:
        # PowerShell命令获取UWP应用
        ps_command = "Get-AppxPackage | Select-Object Name, PackageFullName, Version, Publisher | ConvertTo-Json"
        
        # 执行PowerShell命令
        result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            # 解析JSON结果
            try:
                ps_result = json.loads(result.stdout)
                
                # 确保结果是列表
                if not isinstance(ps_result, list):
                    ps_result = [ps_result]
                
                # 处理结果
                for item in ps_result:
                    app_info = {
                        "name": item.get("Name", ""),
                        "full_name": item.get("PackageFullName", ""),
                        "version": item.get("Version", ""),
                        "publisher": item.get("Publisher", ""),
                        "type": "UWP"
                    }
                    
                    uwp_apps.append(app_info)
            except json.JSONDecodeError as e:
                print(f"Error parsing UWP apps output: {e}")
        else:
            print(f"PowerShell command for UWP apps failed: {result.stderr}")
    except Exception as e:
        print(f"Error running PowerShell command for UWP apps: {e}")
    
    return uwp_apps

def get_startup_items():
    """
    获取开机自启动项
    
    返回:
    - 自启动项列表
    """
    startup_items = []
    
    # 注册表中的自启动项路径
    startup_paths = [
        (winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\Run"),
        (winreg.HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\Run"),
        (winreg.HKEY_CURRENT_USER, "Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce"),
        (winreg.HKEY_LOCAL_MACHINE, "Software\\Microsoft\\Windows\\CurrentVersion\\RunOnce")
    ]
    
    # 获取注册表中的自启动项
    for hkey, path in startup_paths:
        try:
            with winreg.OpenKey(hkey, path) as key:
                # 获取值的数量
                value_count = winreg.QueryInfoKey(key)[1]
                
                # 遍历所有值
                for i in range(value_count):
                    try:
                        name, value, _ = winreg.EnumValue(key, i)
                        
                        registry_location = "HKCU" if hkey == winreg.HKEY_CURRENT_USER else "HKLM"
                        registry_path = f"{registry_location}\\{path}"
                        
                        startup_items.append({
                            "name": name,
                            "command": value,
                            "location": registry_path,
                            "type": "Registry"
                        })
                    except WindowsError:
                        continue
        except WindowsError:
            continue
    
    # 获取启动文件夹中的项目
    startup_folders = [
        os.path.join(os.environ["APPDATA"], "Microsoft\\Windows\\Start Menu\\Programs\\Startup"),
        os.path.join(os.environ["PROGRAMDATA"], "Microsoft\\Windows\\Start Menu\\Programs\\Startup")
    ]
    
    for folder in startup_folders:
        if os.path.exists(folder):
            for item in os.listdir(folder):
                item_path = os.path.join(folder, item)
                if os.path.isfile(item_path):
                    startup_items.append({
                        "name": item,
                        "location": folder,
                        "path": item_path,
                        "type": "Startup Folder"
                    })
    
    return startup_items

def get_all_installed_software():
    """
    获取所有已安装的软件列表，合并来自不同来源的结果
    
    返回:
    - 综合的软件列表
    """
    all_software = []
    
    # 获取64位软件
    software_64bit = get_installed_software_from_registry(
        "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Uninstall", 
        winreg.KEY_WOW64_64KEY
    )
    for software in software_64bit:
        software["architecture"] = "64-bit"
        # 确保使用通用字段名，与HTML报告中的期望匹配
        if "name" in software and "DisplayName" not in software:
            software["DisplayName"] = software["name"]
        if "version" in software and "DisplayVersion" not in software:
            software["DisplayVersion"] = software["version"]
        if "publisher" in software and "Publisher" not in software:
            software["Publisher"] = software["publisher"]
        if "install_date" in software and "InstallDate" not in software:
            software["InstallDate"] = software["install_date"]
        if "install_location" in software and "InstallLocation" not in software:
            software["InstallLocation"] = software["install_location"]
        all_software.append(software)
    
    # 获取32位软件
    software_32bit = get_installed_software_from_registry(
        "SOFTWARE\\Wow6432Node\\Microsoft\\Windows\\CurrentVersion\\Uninstall", 
        winreg.KEY_WOW64_32KEY
    )
    for software in software_32bit:
        software["architecture"] = "32-bit"
        # 确保使用通用字段名，与HTML报告中的期望匹配
        if "name" in software and "DisplayName" not in software:
            software["DisplayName"] = software["name"]
        if "version" in software and "DisplayVersion" not in software:
            software["DisplayVersion"] = software["version"]
        if "publisher" in software and "Publisher" not in software:
            software["Publisher"] = software["publisher"]
        if "install_date" in software and "InstallDate" not in software:
            software["InstallDate"] = software["install_date"]
        if "install_location" in software and "InstallLocation" not in software:
            software["InstallLocation"] = software["install_location"]
        all_software.append(software)
    
    # 获取PowerShell结果作为备份方法
    if not all_software:
        try:
            ps_software = get_installed_software_from_powershell()
            for software in ps_software:
                # 确保使用通用字段名
                if "name" in software and "DisplayName" not in software:
                    software["DisplayName"] = software["name"]
                if "version" in software and "DisplayVersion" not in software:
                    software["DisplayVersion"] = software["version"]
                if "publisher" in software and "Publisher" not in software:
                    software["Publisher"] = software["publisher"]
                if "install_date" in software and "InstallDate" not in software:
                    software["InstallDate"] = software["install_date"]
                all_software.append(software)
        except Exception as e:
            print(f"Error getting software from PowerShell: {e}")
    
    # 获取UWP应用
    try:
        uwp_apps = get_uwp_apps()
        for app in uwp_apps:
            app_info = {
                "DisplayName": app.get("name", ""),
                "DisplayVersion": app.get("version", ""),
                "Publisher": app.get("publisher", ""),
                "architecture": "UWP",
                "type": "UWP"
            }
            all_software.append(app_info)
    except Exception as e:
        print(f"Error getting UWP apps: {e}")
    
    # 去重（基于名称和版本）
    seen = set()
    unique_software = []
    
    for software in all_software:
        # 创建唯一标识符
        name = software.get("DisplayName", software.get("name", ""))
        version = software.get("DisplayVersion", software.get("version", ""))
        key = (name, version)
        
        if key not in seen and name:  # 确保名称不为空
            seen.add(key)
            unique_software.append(software)
    
    # 按名称排序
    unique_software.sort(key=lambda x: x.get("DisplayName", x.get("name", "")).lower())
    
    # 如果列表仍然为空，尝试其他方法收集软件
    if not unique_software:
        print("Warning: No software found using registry. Trying alternative methods...")
        try:
            import winapps
            for app in winapps.list_installed():
                app_info = {
                    "DisplayName": app.name,
                    "DisplayVersion": app.version,
                    "Publisher": app.publisher,
                    "InstallDate": app.install_date,
                    "InstallLocation": app.install_location,
                    "source": "winapps"
                }
                unique_software.append(app_info)
        except ImportError:
            print("winapps module not available")
        except Exception as e:
            print(f"Error using winapps module: {e}")
    
    return unique_software

def save_software_list(output_dir, filename="installed_apps.json"):
    """
    保存软件列表到JSON文件
    
    参数:
    - output_dir: 输出目录
    - filename: 输出文件名
    
    返回:
    - 保存结果信息
    """
    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 获取软件列表
        software_list = get_all_installed_software()
        
        # 保存软件列表
        output_file = os.path.join(output_dir, filename)
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(software_list, f, indent=2, ensure_ascii=False)
        
        # 保存启动项
        startup_items = get_startup_items()
        startup_file = os.path.join(output_dir, "startup_items.json")
        with open(startup_file, 'w', encoding='utf-8') as f:
            json.dump(startup_items, f, indent=2, ensure_ascii=False)
        
        return {
            "success": True,
            "message": f"找到 {len(software_list)} 个已安装的软件，已保存到 {output_file}",
            "software_count": len(software_list),
            "startup_count": len(startup_items),
            "software_file": output_file,
            "startup_file": startup_file
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"保存软件列表时出错: {str(e)}",
            "error": str(e)
        } 