import subprocess
import os
from pathlib import Path

def backup_drivers(output_dir):
    """
    备份所有已安装的驱动程序到指定目录
    """
    output_dir = Path(output_dir)
    drivers_dir = output_dir / "drivers"
    
    # 确保输出目录存在
    os.makedirs(drivers_dir, exist_ok=True)
    
    # 使用DISM工具导出所有驱动程序
    try:
        result = subprocess.run(
            ["dism", "/online", "/export-driver", f"/destination:{drivers_dir}"],
            capture_output=True, text=True, check=True
        )
        return {
            "success": True,
            "message": f"驱动程序已成功备份到 {drivers_dir}",
            "details": result.stdout
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "message": f"驱动程序备份失败",
            "error": str(e),
            "details": e.stderr
        }

def list_drivers():
    """
    列出系统中所有已安装的驱动程序
    """
    try:
        result = subprocess.run(
            ["pnputil", "/enum-drivers"],
            capture_output=True, text=True, check=True
        )
        return {
            "success": True,
            "drivers": result.stdout
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "error": str(e),
            "details": e.stderr
        }

def get_driver_details():
    """
    获取有关设备和驱动程序的详细信息
    """
    try:
        # 获取设备列表
        result = subprocess.run(
            ["driverquery", "/v", "/fo", "list"],
            capture_output=True, text=True, check=True
        )
        return {
            "success": True,
            "device_list": result.stdout
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "error": str(e),
            "details": e.stderr
        }
        
def restore_drivers(drivers_path):
    """
    从备份目录恢复驱动程序
    """
    drivers_path = Path(drivers_path)
    
    if not drivers_path.exists() or not drivers_path.is_dir():
        return {
            "success": False,
            "message": f"驱动程序路径 {drivers_path} 不存在或不是目录"
        }
        
    try:
        # 使用pnputil命令添加驱动程序
        result = subprocess.run(
            ["pnputil", "/add-driver", f"{drivers_path}\\*.inf", "/subdirs", "/install"],
            capture_output=True, text=True, check=True
        )
        return {
            "success": True,
            "message": "驱动程序恢复成功",
            "details": result.stdout
        }
    except subprocess.CalledProcessError as e:
        return {
            "success": False,
            "message": "驱动程序恢复失败",
            "error": str(e),
            "details": e.stderr
        } 