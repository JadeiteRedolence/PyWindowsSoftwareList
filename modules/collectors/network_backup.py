import subprocess
import os
import json
from pathlib import Path

def backup_wifi_profiles(output_dir):
    """
    备份所有Wi-Fi网络配置文件
    """
    output_dir = Path(output_dir)
    wifi_dir = output_dir / "wifi_profiles"
    
    # 确保输出目录存在
    os.makedirs(wifi_dir, exist_ok=True)
    
    try:
        # 导出所有Wi-Fi配置文件
        result = subprocess.run(
            ["netsh", "wlan", "export", "profile", "key=clear", f"folder={wifi_dir}"],
            capture_output=True, text=True, check=True
        )
        return {
            "success": True,
            "message": f"Wi-Fi配置文件已成功备份到 {wifi_dir}",
            "details": result.stdout
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "message": "Wi-Fi配置文件备份失败",
            "error": str(e),
            "details": e.stderr
        }

def backup_network_settings(output_dir):
    """
    备份TCP/IP配置
    """
    output_dir = Path(output_dir)
    network_file = output_dir / "network_settings.txt"
    
    try:
        # 导出TCP/IP配置
        result = subprocess.run(
            ["netsh", "-c", "interface", "dump"],
            capture_output=True, text=True, check=True
        )
        
        # 写入配置文件
        with open(network_file, "w", encoding="utf-8") as f:
            f.write(result.stdout)
        
        return {
            "success": True,
            "message": f"网络设置已成功备份到 {network_file}",
            "file_path": str(network_file)
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "message": "网络设置备份失败",
            "error": str(e),
            "details": e.stderr
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"网络设置备份失败: {str(e)}"
        }

def backup_wired_profiles(output_dir):
    """
    备份有线网络配置文件（802.1X认证设置）
    """
    output_dir = Path(output_dir)
    wired_dir = output_dir / "wired_profiles"
    
    # 确保输出目录存在
    os.makedirs(wired_dir, exist_ok=True)
    
    try:
        # 导出有线网络配置文件
        result = subprocess.run(
            ["netsh", "lan", "export", "profile", f"folder={wired_dir}"],
            capture_output=True, text=True, check=True
        )
        return {
            "success": True,
            "message": f"有线网络配置文件已成功备份到 {wired_dir}",
            "details": result.stdout
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "message": "有线网络配置文件备份失败",
            "error": str(e),
            "details": e.stderr
        }

def list_wifi_profiles():
    """
    列出所有已保存的Wi-Fi网络配置文件
    """
    try:
        result = subprocess.run(
            ["netsh", "wlan", "show", "profiles"],
            capture_output=True, text=True, check=True
        )
        return {
            "success": True,
            "profiles": result.stdout
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "message": "无法列出Wi-Fi配置文件",
            "error": str(e),
            "details": e.stderr
        }

def restore_wifi_profile(profile_file):
    """
    恢复单个Wi-Fi网络配置文件
    """
    profile_file = Path(profile_file)
    
    if not profile_file.exists() or not profile_file.is_file():
        return {
            "success": False,
            "message": f"配置文件 {profile_file} 不存在或不是文件"
        }
        
    try:
        # 使用netsh命令添加Wi-Fi配置文件
        result = subprocess.run(
            ["netsh", "wlan", "add", "profile", f"filename={profile_file}"],
            capture_output=True, text=True, check=True
        )
        return {
            "success": True,
            "message": f"Wi-Fi配置文件 {profile_file.name} 恢复成功",
            "details": result.stdout
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "message": f"Wi-Fi配置文件 {profile_file.name} 恢复失败",
            "error": str(e),
            "details": e.stderr
        }

def restore_all_wifi_profiles(profiles_dir):
    """
    恢复目录中的所有Wi-Fi网络配置文件
    """
    profiles_dir = Path(profiles_dir)
    
    if not profiles_dir.exists() or not profiles_dir.is_dir():
        return {
            "success": False,
            "message": f"配置文件目录 {profiles_dir} 不存在或不是目录"
        }
    
    results = []
    
    # 查找所有XML文件并恢复
    for profile_file in profiles_dir.glob("*.xml"):
        result = restore_wifi_profile(profile_file)
        results.append({
            "file": profile_file.name,
            "result": result
        })
    
    return {
        "success": True,
        "message": f"尝试恢复 {len(results)} 个Wi-Fi配置文件",
        "details": results
    }

def restore_network_settings(settings_file):
    """
    恢复TCP/IP配置
    """
    settings_file = Path(settings_file)
    
    if not settings_file.exists() or not settings_file.is_file():
        return {
            "success": False,
            "message": f"设置文件 {settings_file} 不存在或不是文件"
        }
        
    try:
        # 使用netsh命令恢复网络设置
        result = subprocess.run(
            ["netsh", "-f", f"{settings_file}"],
            capture_output=True, text=True, check=True
        )
        return {
            "success": True,
            "message": "网络设置恢复成功",
            "details": result.stdout
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "message": "网络设置恢复失败",
            "error": str(e),
            "details": e.stderr
        } 