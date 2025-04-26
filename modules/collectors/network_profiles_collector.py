import os
import subprocess
import json
from datetime import datetime
import xml.etree.ElementTree as ET

def get_wifi_profiles():
    """
    获取Windows上保存的WiFi配置文件信息
    
    返回:
    - 包含WiFi配置文件信息的列表
    """
    wifi_profiles = []
    
    try:
        # 获取所有WiFi配置文件名称
        result = subprocess.run(["netsh", "wlan", "show", "profiles"], capture_output=True, text=True, encoding='gb2312')
        
        if result.returncode == 0:
            lines = result.stdout.splitlines()
            profile_lines = [line for line in lines if "所有用户配置文件" in line or "All User Profile" in line]
            
            for line in profile_lines:
                # 提取配置文件名称
                try:
                    if "所有用户配置文件" in line:
                        profile_name = line.split(":", 1)[1].strip()
                    else:
                        profile_name = line.split(":", 1)[1].strip()
                    
                    # 获取特定WiFi配置文件的详细信息，包括密码
                    detail_cmd = ["netsh", "wlan", "show", "profile", f"name={profile_name}", "key=clear"]
                    detail_result = subprocess.run(detail_cmd, capture_output=True, text=True, encoding='gb2312')
                    
                    if detail_result.returncode == 0:
                        profile_info = {
                            "name": profile_name,
                            "details": {}
                        }
                        
                        detail_lines = detail_result.stdout.splitlines()
                        
                        # 提取常见的WiFi信息
                        for detail_line in detail_lines:
                            if ":" in detail_line:
                                key, value = detail_line.split(":", 1)
                                key = key.strip()
                                value = value.strip()
                                
                                # 处理关键信息
                                if "SSID 名称" in key or "SSID Name" in key:
                                    profile_info["details"]["ssid"] = value
                                elif "身份验证" in key or "Authentication" in key:
                                    profile_info["details"]["authentication"] = value
                                elif "加密" in key or "Cipher" in key:
                                    profile_info["details"]["encryption"] = value
                                elif "安全密钥" in key or "Security key" in key:
                                    profile_info["details"]["security_key_present"] = ("存在" in value or "Present" in value)
                                elif "关键内容" in key or "Key Content" in key:
                                    profile_info["details"]["password"] = value
                                elif "连接模式" in key or "Connection mode" in key:
                                    profile_info["details"]["connection_mode"] = value
                        
                        wifi_profiles.append(profile_info)
                except Exception as e:
                    print(f"处理WiFi配置文件 {profile_name} 时出错: {e}")
    except Exception as e:
        print(f"获取WiFi配置文件时出错: {e}")
    
    return wifi_profiles

def get_wifi_profiles_xml():
    """
    获取并保存所有WiFi配置文件的XML配置
    
    返回:
    - 包含WiFi配置文件XML内容的字典
    """
    wifi_xml_profiles = {}
    
    try:
        # 获取所有WiFi配置文件名称
        result = subprocess.run(["netsh", "wlan", "show", "profiles"], capture_output=True, text=True, encoding='gb2312')
        
        if result.returncode == 0:
            lines = result.stdout.splitlines()
            profile_lines = [line for line in lines if "所有用户配置文件" in line or "All User Profile" in line]
            
            for line in profile_lines:
                # 提取配置文件名称
                try:
                    if "所有用户配置文件" in line:
                        profile_name = line.split(":", 1)[1].strip()
                    else:
                        profile_name = line.split(":", 1)[1].strip()
                    
                    # 导出WiFi配置文件为XML
                    export_cmd = ["netsh", "wlan", "export", "profile", f"name={profile_name}", "folder=%TEMP%", "key=clear"]
                    export_result = subprocess.run(export_cmd, capture_output=True, text=True, encoding='gb2312')
                    
                    if export_result.returncode == 0:
                        # 找到生成的XML文件路径
                        temp_dir = os.environ.get('TEMP', '')
                        xml_file_path = os.path.join(temp_dir, f"WLAN-{profile_name}.xml")
                        
                        if os.path.exists(xml_file_path):
                            # 读取XML文件内容
                            with open(xml_file_path, 'r', encoding='utf-8') as f:
                                xml_content = f.read()
                                
                            wifi_xml_profiles[profile_name] = {
                                "xml_content": xml_content,
                                "file_path": xml_file_path
                            }
                            
                            # 解析XML内容，提取SSID和密码
                            try:
                                tree = ET.fromstring(xml_content)
                                ns = {'ns': 'http://www.microsoft.com/networking/WLAN/profile/v1'}
                                
                                # 提取SSID
                                ssid_elem = tree.find('.//ns:SSID/ns:name', ns)
                                ssid = ssid_elem.text if ssid_elem is not None else None
                                
                                # 提取密码
                                key_elem = tree.find('.//ns:keyMaterial', ns)
                                password = key_elem.text if key_elem is not None else None
                                
                                wifi_xml_profiles[profile_name]["parsed"] = {
                                    "ssid": ssid,
                                    "password": password
                                }
                            except Exception as parse_err:
                                print(f"解析WiFi配置文件XML {profile_name} 时出错: {parse_err}")
                except Exception as e:
                    print(f"导出WiFi配置文件 {profile_name} 时出错: {e}")
    except Exception as e:
        print(f"获取WiFi配置文件XML时出错: {e}")
    
    return wifi_xml_profiles

def get_vpn_connections():
    """
    获取Windows VPN连接信息
    
    返回:
    - 包含VPN连接信息的列表
    """
    vpn_connections = []
    
    try:
        # 使用PowerShell获取VPN连接信息
        ps_command = "Get-VpnConnection | Select-Object Name, ServerAddress, TunnelType, AuthenticationMethod, EncryptionLevel, ConnectionStatus | ConvertTo-Json"
        result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            vpn_data = json.loads(result.stdout)
            
            # 确保结果是列表
            if not isinstance(vpn_data, list):
                vpn_data = [vpn_data]
            
            for vpn in vpn_data:
                vpn_connections.append({
                    "name": vpn.get("Name"),
                    "server_address": vpn.get("ServerAddress"),
                    "tunnel_type": vpn.get("TunnelType"),
                    "authentication_method": vpn.get("AuthenticationMethod"),
                    "encryption_level": vpn.get("EncryptionLevel"),
                    "connection_status": vpn.get("ConnectionStatus")
                })
    except Exception as e:
        print(f"获取VPN连接信息时出错: {e}")
    
    return vpn_connections

def get_network_shares():
    """
    获取网络共享信息
    
    返回:
    - 包含网络共享信息的列表
    """
    network_shares = []
    
    try:
        # 使用PowerShell获取网络共享信息
        ps_command = "Get-SmbShare | Select-Object Name, Path, Description, ShareState | ConvertTo-Json"
        result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            share_data = json.loads(result.stdout)
            
            # 确保结果是列表
            if not isinstance(share_data, list):
                share_data = [share_data]
            
            for share in share_data:
                network_shares.append({
                    "name": share.get("Name"),
                    "path": share.get("Path"),
                    "description": share.get("Description"),
                    "share_state": share.get("ShareState")
                })
    except Exception as e:
        print(f"获取网络共享信息时出错: {e}")
    
    return network_shares

def get_network_drives():
    """
    获取映射的网络驱动器信息
    
    返回:
    - 包含映射的网络驱动器信息的列表
    """
    network_drives = []
    
    try:
        # 使用PowerShell获取映射的网络驱动器
        ps_command = "Get-PSDrive -PSProvider FileSystem | Where-Object { $_.DisplayRoot -ne $null } | Select-Object Name, Root, DisplayRoot, Description | ConvertTo-Json"
        result = subprocess.run(["powershell", "-Command", ps_command], capture_output=True, text=True)
        
        if result.returncode == 0 and result.stdout.strip():
            drive_data = json.loads(result.stdout)
            
            # 确保结果是列表
            if not isinstance(drive_data, list):
                drive_data = [drive_data]
            
            for drive in drive_data:
                network_drives.append({
                    "name": drive.get("Name"),
                    "root": drive.get("Root"),
                    "network_path": drive.get("DisplayRoot"),
                    "description": drive.get("Description")
                })
    except Exception as e:
        print(f"获取映射的网络驱动器信息时出错: {e}")
    
    return network_drives

def collect_all_network_profiles():
    """
    收集所有网络配置文件信息
    
    返回:
    - 包含所有网络配置文件信息的字典
    """
    network_profiles = {
        "collection_time": datetime.now().isoformat(),
        "wifi_profiles": get_wifi_profiles(),
        "wifi_xml_profiles": get_wifi_profiles_xml(),
        "vpn_connections": get_vpn_connections(),
        "network_shares": get_network_shares(),
        "network_drives": get_network_drives()
    }
    
    return network_profiles

def save_network_profiles(output_dir, filename="network_profiles.json"):
    """
    保存网络配置文件信息到JSON文件
    
    参数:
    - output_dir: 输出目录
    - filename: 输出文件名
    
    返回:
    - 保存结果信息
    """
    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 收集网络配置文件信息
        network_profiles = collect_all_network_profiles()
        
        # 保存网络配置文件信息
        output_file = os.path.join(output_dir, filename)
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(network_profiles, f, indent=2, ensure_ascii=False)
        
        return {
            "success": True,
            "message": f"网络配置文件信息已保存到 {output_file}",
            "file": output_file
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"保存网络配置文件信息时出错: {str(e)}",
            "error": str(e)
        } 