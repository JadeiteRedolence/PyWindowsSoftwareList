import subprocess
import os
import platform
import json
import winreg
from pathlib import Path
import socket
import psutil
from datetime import datetime
import uuid
import re
import ctypes
import logging
import shutil

def get_powershell_path():
    """
    获取PowerShell的绝对路径
    
    返回:
    - PowerShell执行文件的绝对路径
    """
    try:
        # 常见的PowerShell路径列表
        possible_paths = [
            # PowerShell 7+ (cross-platform)
            r"C:\Program Files\PowerShell\7\pwsh.exe",
            r"C:\Program Files\PowerShell\7-preview\pwsh.exe", 
            # Windows PowerShell 5.1
            r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe",
            r"C:\Windows\SysWOW64\WindowsPowerShell\v1.0\powershell.exe",
        ]
        
        # 检查每个可能的路径
        for path in possible_paths:
            if os.path.exists(path):
                logging.debug(f"Found PowerShell at: {path}")
                return path
        
        # 如果找不到，尝试使用where命令查找
        try:
            result = subprocess.run(['where', 'powershell'], capture_output=True, text=True, timeout=5)
            if result.returncode == 0 and result.stdout.strip():
                path = result.stdout.strip().split('\n')[0]
                logging.debug(f"Found PowerShell via 'where' command: {path}")
                return path
        except (subprocess.TimeoutExpired, Exception) as e:
            logging.warning(f"Error finding PowerShell with 'where' command: {e}")
        
        # 尝试查找pwsh
        try:
            result = subprocess.run(['where', 'pwsh'], capture_output=True, text=True, timeout=5)
            if result.returncode == 0 and result.stdout.strip():
                path = result.stdout.strip().split('\n')[0]
                logging.debug(f"Found pwsh via 'where' command: {path}")
                return path
        except (subprocess.TimeoutExpired, Exception) as e:
            logging.warning(f"Error finding pwsh with 'where' command: {e}")
        
        # 最后尝试默认路径
        logging.warning("Could not find PowerShell absolute path, using 'powershell' as fallback")
        return "powershell"
        
    except Exception as e:
        logging.error(f"Error getting PowerShell path: {e}")
        return "powershell"

def run_powershell_command(command, timeout=30):
    """
    执行PowerShell命令，带超时机制
    
    参数:
    - command: PowerShell命令字符串
    - timeout: 超时时间（秒），默认30秒
    
    返回:
    - subprocess.CompletedProcess对象
    """
    try:
        powershell_path = get_powershell_path()
        
        # 检测PowerShell版本并调整命令
        is_pwsh7 = powershell_path.endswith('pwsh.exe')
        
        # 对于PowerShell 7+，将Get-WmiObject替换为Get-CimInstance
        if is_pwsh7 and 'Get-WmiObject' in command:
            logging.debug("Converting Get-WmiObject to Get-CimInstance for PowerShell 7+")
            # 替换常见的WMI命令
            command = command.replace('Get-WmiObject -Class Win32_Processor', 'Get-CimInstance -ClassName Win32_Processor')
            command = command.replace('Get-WmiObject -Class Win32_PhysicalMemory', 'Get-CimInstance -ClassName Win32_PhysicalMemory')
            command = command.replace('Get-WmiObject -Class Win32_VideoController', 'Get-CimInstance -ClassName Win32_VideoController')
            command = command.replace('Get-WmiObject Win32_BaseBoard', 'Get-CimInstance -ClassName Win32_BaseBoard')
            command = command.replace('Get-WmiObject Win32_BIOS', 'Get-CimInstance -ClassName Win32_BIOS')
            command = command.replace('Get-WmiObject Win32_PnPSignedDriver', 'Get-CimInstance -ClassName Win32_PnPSignedDriver')
        
        # 构建完整命令，确保使用UTF-8编码
        if is_pwsh7:
            # PowerShell 7+ 支持更好的Unicode
            full_cmd = [powershell_path, '-Command', command]
        else:
            # Windows PowerShell 5.1 - 在命令内部设置编码
            encoded_cmd = f"[Console]::OutputEncoding = [System.Text.Encoding]::UTF8; $OutputEncoding = [System.Text.Encoding]::UTF8; {command}"
            full_cmd = [powershell_path, '-Command', encoded_cmd]
        
        logging.debug(f"Executing PowerShell command with timeout {timeout}s: {command[:100]}...")
        
        # 执行命令，带超时，使用UTF-8编码并禁用错误处理
        result = subprocess.run(
            full_cmd, 
            capture_output=True, 
            text=True, 
            shell=False,  # 使用绝对路径，不需要shell=True
            encoding='utf-8',
            errors='ignore',  # 忽略编码错误，防止崩溃
            timeout=timeout
        )
        
        if result.returncode != 0:
            logging.warning(f"PowerShell command failed with return code {result.returncode}: {result.stderr[:200]}")
        
        return result
        
    except subprocess.TimeoutExpired as e:
        logging.error(f"PowerShell command timed out after {timeout} seconds: {command[:100]}")
        # 返回一个模拟的失败结果
        return subprocess.CompletedProcess([], returncode=1, stdout="", stderr=f"Command timed out after {timeout} seconds")
    except Exception as e:
        logging.error(f"Error executing PowerShell command: {e}")
        return subprocess.CompletedProcess([], returncode=1, stdout="", stderr=str(e))

def run_wmic_command(wmi_class, properties=None, timeout=30):
    """
    使用原生WMIC工具执行WMI查询，作为PowerShell的备用方案
    
    参数:
    - wmi_class: WMI类名，如 'Win32_Processor'
    - properties: 要查询的属性列表，如果为None则查询所有
    - timeout: 超时时间（秒）
    
    返回:
    - subprocess.CompletedProcess对象
    """
    try:
        # 构建WMIC命令
        wmic_path = r"C:\Windows\System32\wbem\wmic.exe"
        
        # 检查WMIC是否存在
        if not os.path.exists(wmic_path):
            logging.warning("WMIC tool not found, trying alternative path")
            wmic_path = "wmic"  # 使用PATH中的版本
        
        if properties:
            prop_list = ",".join(properties)
            cmd = [wmic_path, wmi_class, "get", prop_list, "/format:csv"]
        else:
            cmd = [wmic_path, wmi_class, "get", "/format:csv"]
        
        logging.debug(f"Executing WMIC command: {' '.join(cmd)}")
        
        result = subprocess.run(
            cmd,
            capture_output=True,
            text=True,
            shell=False,
            encoding='gbk',  # WMIC输出通常是GBK编码
            errors='ignore',
            timeout=timeout
        )
        
        if result.returncode != 0:
            logging.warning(f"WMIC command failed with return code {result.returncode}: {result.stderr[:200]}")
        
        return result
        
    except subprocess.TimeoutExpired as e:
        logging.error(f"WMIC command timed out after {timeout} seconds")
        return subprocess.CompletedProcess([], returncode=1, stdout="", stderr=f"Command timed out after {timeout} seconds")
    except Exception as e:
        logging.error(f"Error executing WMIC command: {e}")
        return subprocess.CompletedProcess([], returncode=1, stdout="", stderr=str(e))

def get_system_info():
    """
    获取系统基本信息
    
    返回:
    - 系统信息字典
    """
    system_info = {
        "platform": platform.platform(),
        "system": platform.system(),
        "release": platform.release(),
        "version": platform.version(),
        "architecture": platform.machine(),
        "processor": platform.processor(),
        "hostname": socket.gethostname(),
        "ip_address": socket.gethostbyname(socket.gethostname()),
        "mac_address": ':'.join(re.findall('..', '%012x' % uuid.getnode()))
    }
    
    return system_info

def get_cpu_info():
    """
    获取CPU详细信息
    
    返回:
    - CPU信息字典
    """
    cpu_info = {}
    
    try:
        if platform.system() == "Windows":
            # 使用PowerShell获取CPU信息
            logging.debug("Getting CPU information")
            cmd = "$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'; [System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'; [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'; Get-WmiObject -Class Win32_Processor | Select-Object Name, NumberOfCores, NumberOfLogicalProcessors, MaxClockSpeed, L2CacheSize, L3CacheSize | ConvertTo-Json"
            result = run_powershell_command(cmd, timeout=20)
            
            if result.returncode == 0 and result.stdout.strip():
                try:
                    cpu_data = json.loads(result.stdout)
                    
                    # 处理单CPU和多CPU的情况
                    if isinstance(cpu_data, list):
                        cpu_info = cpu_data
                    else:
                        cpu_info = [cpu_data]
                    
                    # 将时钟速度转换为GHz
                    for cpu in cpu_info:
                        if "MaxClockSpeed" in cpu:
                            cpu["MaxClockSpeedGHz"] = round(cpu["MaxClockSpeed"] / 1000, 2)
                    
                    logging.debug(f"Successfully retrieved CPU information for {len(cpu_info)} processors")
                except json.JSONDecodeError as e:
                    logging.warning(f"Error parsing CPU JSON: {e}, output: {result.stdout[:100]}")
                    cpu_info = []  # 重置以便尝试备用方案
            else:
                logging.warning(f"PowerShell command failed: {result.stderr}")
                cpu_info = []  # 重置以便尝试备用方案
            
            # 如果PowerShell失败，尝试使用WMIC
            if not cpu_info:
                logging.debug("Trying WMIC as fallback for CPU info")
                wmic_result = run_wmic_command("cpu", ["Name", "NumberOfCores", "NumberOfLogicalProcessors", "MaxClockSpeed"])
                
                if wmic_result.returncode == 0 and wmic_result.stdout.strip():
                    try:
                        # 解析WMIC CSV输出
                        lines = wmic_result.stdout.strip().split('\n')
                        if len(lines) > 1:  # 跳过标题行
                            for line in lines[1:]:
                                if line.strip() and not line.startswith('Node,'):
                                    fields = line.split(',')
                                    if len(fields) >= 5:  # Node, MaxClockSpeed, Name, NumberOfCores, NumberOfLogicalProcessors
                                        try:
                                            cpu_item = {
                                                "Name": fields[2].strip() if len(fields) > 2 else "",
                                                "NumberOfCores": int(fields[3]) if fields[3].strip().isdigit() else None,
                                                "NumberOfLogicalProcessors": int(fields[4]) if fields[4].strip().isdigit() else None,
                                                "MaxClockSpeed": int(fields[1]) if fields[1].strip().isdigit() else None
                                            }
                                            
                                            # 转换时钟速度为GHz
                                            if cpu_item["MaxClockSpeed"]:
                                                cpu_item["MaxClockSpeedGHz"] = round(cpu_item["MaxClockSpeed"] / 1000, 2)
                                            
                                            if not cpu_info:
                                                cpu_info = []
                                            cpu_info.append(cpu_item)
                                            
                                        except (ValueError, IndexError) as e:
                                            logging.debug(f"Error parsing WMIC CPU line: {e}")
                                            continue
                        
                        if cpu_info:
                            logging.debug(f"Successfully retrieved CPU information via WMIC for {len(cpu_info)} processors")
                        
                    except Exception as e:
                        logging.warning(f"Error parsing WMIC CPU output: {e}")
            
            # 如果仍然没有数据，使用psutil作为最后手段
            if not cpu_info:
                logging.debug("Using psutil for CPU info as final fallback")
                cpu_info = [{
                    "Name": platform.processor(),
                    "NumberOfCores": psutil.cpu_count(logical=False),
                    "NumberOfLogicalProcessors": psutil.cpu_count(logical=True),
                    "Frequency": psutil.cpu_freq(),
                }]
        else:
            # Linux系统使用lscpu
            cpu_info = {"platform_not_supported": True}
    except Exception as e:
        logging.error(f"Error getting CPU information: {e}", exc_info=True)
        
        # 使用psutil作为最后手段
        try:
            cpu_info = [{
                "Name": platform.processor(),
                "NumberOfCores": psutil.cpu_count(logical=False),
                "NumberOfLogicalProcessors": psutil.cpu_count(logical=True),
                "Error": str(e)
            }]
        except:
            cpu_info = {"error": str(e)}
    
    return cpu_info

def get_memory_info():
    """
    获取内存信息
    
    返回:
    - 内存信息字典
    """
    memory_info = {}
    
    try:
        if platform.system() == "Windows":
            # 首先尝试从psutil获取内存信息（更可靠）
            logging.debug("Getting memory information using psutil")
            svmem = psutil.virtual_memory()
            memory_info["total_bytes"] = svmem.total
            memory_info["total_gb"] = round(svmem.total / (1024**3), 2)
            memory_info["available_bytes"] = svmem.available
            memory_info["available_gb"] = round(svmem.available / (1024**3), 2)
            memory_info["used_bytes"] = svmem.used
            memory_info["used_gb"] = round(svmem.used / (1024**3), 2)
            memory_info["percent"] = svmem.percent
            
            memory_info["current_usage"] = {
                "total_mb": round(svmem.total / (1024**2), 2),
                "used_mb": round(svmem.used / (1024**2), 2),
                "free_mb": round(svmem.available / (1024**2), 2),
                "usage_percent": svmem.percent
            }
            
            # 尝试获取内存模块详细信息
            logging.debug("Attempting to get memory module details via PowerShell")
            try:
                ps_cmd = "$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'; [System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'; [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'; Get-WmiObject -Class Win32_PhysicalMemory | Select-Object Capacity, Speed, Manufacturer, PartNumber, DeviceLocator | ConvertTo-Json"
                result = run_powershell_command(ps_cmd, timeout=15)
                
                if result.returncode == 0 and result.stdout.strip():
                    try:
                        memory_modules = json.loads(result.stdout)
                        
                        # 处理单个模块的情况
                        if isinstance(memory_modules, dict):
                            memory_modules = [memory_modules]
                        
                        # 转换容量为GB
                        for module in memory_modules:
                            if "Capacity" in module and module["Capacity"]:
                                try:
                                    capacity_bytes = int(module["Capacity"])
                                    capacity_gb = round(capacity_bytes / (1024**3), 2)
                                    module["Capacity"] = f"{capacity_gb} GB"
                                except (ValueError, TypeError):
                                    module["Capacity"] = "Unknown"
                        
                        memory_info["modules"] = memory_modules
                        logging.debug(f"Found {len(memory_modules)} memory modules")
                    except json.JSONDecodeError as e:
                        logging.warning(f"Error parsing memory modules JSON: {e}")
                else:
                    logging.warning(f"Failed to get memory module details: {result.stderr}")
            except Exception as e:
                logging.warning(f"Error getting memory module details: {e}")
            
            # 如果PowerShell失败，尝试使用WMIC获取内存模块信息
            if "modules" not in memory_info or not memory_info["modules"]:
                logging.debug("Trying WMIC as fallback for memory module info")
                wmic_result = run_wmic_command("memorychip", ["Capacity", "Speed", "Manufacturer", "PartNumber", "DeviceLocator"])
                
                if wmic_result.returncode == 0 and wmic_result.stdout.strip():
                    try:
                        # 解析WMIC CSV输出
                        lines = wmic_result.stdout.strip().split('\n')
                        memory_modules = []
                        
                        if len(lines) > 1:  # 跳过标题行
                            for line in lines[1:]:
                                if line.strip() and not line.startswith('Node,'):
                                    fields = line.split(',')
                                    if len(fields) >= 6:  # Node, Capacity, DeviceLocator, Manufacturer, PartNumber, Speed
                                        try:
                                            capacity_str = fields[1].strip() if len(fields) > 1 else ""
                                            
                                            # 转换容量为GB
                                            if capacity_str and capacity_str.isdigit():
                                                capacity_bytes = int(capacity_str)
                                                capacity_gb = round(capacity_bytes / (1024**3), 2)
                                                capacity_display = f"{capacity_gb} GB"
                                            else:
                                                capacity_display = "Unknown"
                                            
                                            module_item = {
                                                "Capacity": capacity_display,
                                                "DeviceLocator": fields[2].strip() if len(fields) > 2 else "",
                                                "Manufacturer": fields[3].strip() if len(fields) > 3 else "",
                                                "PartNumber": fields[4].strip() if len(fields) > 4 else "",
                                                "Speed": fields[5].strip() if len(fields) > 5 else ""
                                            }
                                            
                                            memory_modules.append(module_item)
                                            
                                        except (ValueError, IndexError) as e:
                                            logging.debug(f"Error parsing WMIC memory line: {e}")
                                            continue
                        
                        if memory_modules:
                            memory_info["modules"] = memory_modules
                            logging.debug(f"Successfully retrieved memory module information via WMIC for {len(memory_modules)} modules")
                        
                    except Exception as e:
                        logging.warning(f"Error parsing WMIC memory output: {e}")
        else:
            # Linux系统
            memory_info = {"platform_not_supported": True}
    except Exception as e:
        logging.error(f"Error getting memory information: {e}", exc_info=True)
        memory_info = {"error": str(e)}
    
    return memory_info

def get_disk_info():
    """
    获取磁盘信息
    
    返回:
    - 磁盘信息字典
    """
    disk_info = {}
    
    try:
        if platform.system() == "Windows":
            # 使用psutil获取磁盘信息（更可靠的方法）
            logging.debug("Collecting disk information using psutil")
            physical_disks = []
            volumes = []
            
            # 获取物理分区信息
            partitions = psutil.disk_partitions(all=False)
            
            for partition in partitions:
                try:
                    usage = psutil.disk_usage(partition.mountpoint)
                    
                    # 构建分区信息
                    volume_info = {
                        "device": partition.device,
                        "mountpoint": partition.mountpoint,
                        "fstype": partition.fstype,
                        "opts": partition.opts,
                        "total": usage.total,
                        "used": usage.used,
                        "free": usage.free,
                        "percent": usage.percent,
                        "size": usage.total,  # 兼容旧格式
                        "SizeRemaining": usage.free,  # 兼容旧格式
                        "DriveLetter": partition.mountpoint.replace(":\\", ""),
                        "FileSystem": partition.fstype,
                        # 格式化大小
                        "SizeGB": round(usage.total / (1024**3), 2),
                        "FreeGB": round(usage.free / (1024**3), 2),
                        "UsedGB": round(usage.used / (1024**3), 2),
                        "UsedPercent": usage.percent
                    }
                    
                    volumes.append(volume_info)
                    
                    # 添加到物理磁盘列表
                    physical_disk = {
                        "DeviceId": partition.device,
                        "FriendlyName": f"Drive {partition.mountpoint}",
                        "MediaType": "Unknown",
                        "Size": usage.total,
                        "HealthStatus": "Healthy",
                        "SizeFormatted": f"{round(usage.total / (1024**3), 2)} GB" if usage.total < 1024**4 else f"{round(usage.total / (1024**4), 2)} TB"
                    }
                    
                    if physical_disk not in physical_disks:
                        physical_disks.append(physical_disk)
                        
                    logging.debug(f"Added disk {partition.device} with {round(usage.total / (1024**3), 2)} GB total space")
                except Exception as e:
                    logging.warning(f"Error accessing partition {partition.mountpoint}: {e}")
            
            # 保存收集到的信息
            disk_info["physical_disks"] = physical_disks
            disk_info["volumes"] = volumes
            
            # 尝试使用PowerShell命令获取更详细的信息（备选方法）
            try:
                logging.debug("Attempting to get additional disk info via PowerShell")
                # 使用英文区域设置执行PowerShell命令
                cmd = "$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'; [System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'; [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'; Get-PhysicalDisk | Select-Object DeviceId, FriendlyName, MediaType, Size, HealthStatus | ConvertTo-Json"
                result = run_powershell_command(cmd, timeout=20)
                
                if result.returncode == 0 and result.stdout.strip():
                    additional_disks = json.loads(result.stdout)
                    
                    # 处理单磁盘和多磁盘的情况
                    if isinstance(additional_disks, dict):
                        additional_disks = [additional_disks]
                    
                    # 只有在成功获取到更详细信息时才更新physical_disks
                    if additional_disks and len(additional_disks) > 0:
                        disk_info["physical_disks_detailed"] = additional_disks
                        logging.debug(f"Successfully retrieved additional disk details via PowerShell")
                else:
                    logging.warning(f"PowerShell command failed: {result.stderr}")
            except Exception as e:
                logging.warning(f"Failed to get additional disk info via PowerShell: {e}")
        else:
            # Linux系统
            disk_info = {"platform_not_supported": True}
    except Exception as e:
        logging.error(f"Error getting disk information: {e}", exc_info=True)
        disk_info = {"error": str(e)}
    
    # 确保返回的数据是可用的
    if not disk_info.get("volumes") and not disk_info.get("physical_disks"):
        # 创建基本的磁盘信息
        try:
            disk_info["basic_disk_info"] = []
            for letter in "CDEFGHIJKLMNOPQRSTUVWXYZ":
                drive = f"{letter}:"
                if os.path.exists(drive):
                    try:
                        total, used, free = shutil.disk_usage(drive)
                        disk_info["basic_disk_info"].append({
                            "device": drive,
                            "size": total,
                            "free": free,
                            "fstype": "Unknown",
                            "SizeGB": round(total / (1024**3), 2),
                            "FreeGB": round(free / (1024**3), 2)
                        })
                        logging.debug(f"Added basic disk info for drive {drive}")
                    except Exception as inner_e:
                        logging.warning(f"Could not get disk usage for {drive}: {inner_e}")
        except Exception as fallback_e:
            logging.error(f"Failed to create basic disk info: {fallback_e}")
    
    return disk_info

def get_graphics_info():
    """
    获取显卡信息
    
    返回:
    - 显卡信息字典
    """
    graphics_info = {}
    
    try:
        if platform.system() == "Windows":
            # 使用PowerShell获取显卡信息
            logging.debug("Getting graphics card information")
            cmd = "$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'; [System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'; [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'; Get-WmiObject -Class Win32_VideoController | Select-Object Name, AdapterRAM, DriverVersion, VideoProcessor | ConvertTo-Json"
            result = run_powershell_command(cmd, timeout=15)
            
            if result.returncode == 0 and result.stdout.strip():
                try:
                    gpu_data = json.loads(result.stdout)
                    
                    # 处理单显卡和多显卡的情况
                    if isinstance(gpu_data, dict):
                        gpu_data = [gpu_data]
                    
                    # 转换显存大小为MB/GB
                    for gpu in gpu_data:
                        if "AdapterRAM" in gpu and gpu["AdapterRAM"]:
                            try:
                                ram_bytes = int(gpu["AdapterRAM"])
                                gpu["VideoRAM_GB"] = round(ram_bytes / (1024**3), 2)
                            except (ValueError, TypeError):
                                gpu["VideoRAM_GB"] = "Unknown"
                    
                    graphics_info["adapters"] = gpu_data
                    logging.debug(f"Found {len(gpu_data)} graphics adapters")
                except json.JSONDecodeError as e:
                    logging.warning(f"Error parsing graphics card JSON: {e}, output: {result.stdout[:100]}")
                    gpu_data = []
            else:
                logging.warning(f"Failed to get graphics information: {result.stderr}")
                gpu_data = []
            
            # 如果PowerShell失败，尝试使用WMIC
            if "adapters" not in graphics_info or not graphics_info["adapters"]:
                logging.debug("Trying WMIC as fallback for graphics info")
                wmic_result = run_wmic_command("path win32_videocontroller", ["Name", "AdapterRAM", "DriverVersion", "VideoProcessor"])
                
                if wmic_result.returncode == 0 and wmic_result.stdout.strip():
                    try:
                        # 解析WMIC CSV输出
                        lines = wmic_result.stdout.strip().split('\n')
                        gpu_data = []
                        
                        if len(lines) > 1:  # 跳过标题行
                            for line in lines[1:]:
                                if line.strip() and not line.startswith('Node,'):
                                    fields = line.split(',')
                                    if len(fields) >= 5:  # Node, AdapterRAM, DriverVersion, Name, VideoProcessor
                                        try:
                                            gpu_item = {
                                                "Name": fields[3].strip() if len(fields) > 3 else "",
                                                "AdapterRAM": fields[1].strip() if len(fields) > 1 else "",
                                                "DriverVersion": fields[2].strip() if len(fields) > 2 else "",
                                                "VideoProcessor": fields[4].strip() if len(fields) > 4 else ""
                                            }
                                            
                                            # 转换显存大小
                                            if gpu_item["AdapterRAM"] and gpu_item["AdapterRAM"].isdigit():
                                                try:
                                                    ram_bytes = int(gpu_item["AdapterRAM"])
                                                    gpu_item["VideoRAM_GB"] = round(ram_bytes / (1024**3), 2)
                                                except (ValueError, TypeError):
                                                    gpu_item["VideoRAM_GB"] = "Unknown"
                                            else:
                                                gpu_item["VideoRAM_GB"] = "Unknown"
                                            
                                            gpu_data.append(gpu_item)
                                            
                                        except (ValueError, IndexError) as e:
                                            logging.debug(f"Error parsing WMIC graphics line: {e}")
                                            continue
                        
                        if gpu_data:
                            graphics_info["adapters"] = gpu_data
                            logging.debug(f"Successfully retrieved graphics information via WMIC for {len(gpu_data)} adapters")
                        
                    except Exception as e:
                        logging.warning(f"Error parsing WMIC graphics output: {e}")
            
            # 如果仍然没有数据，提供基本信息
            if "adapters" not in graphics_info or not graphics_info["adapters"]:
                graphics_info["adapters"] = [{"Name": "Unknown", "VideoRAM_GB": "Unknown", "DriverVersion": "Unknown"}]
                logging.debug("Using unknown graphics adapter as fallback")
        else:
            # Linux系统
            graphics_info = {"platform_not_supported": True}
    except Exception as e:
        logging.error(f"Error getting graphics information: {e}", exc_info=True)
        graphics_info = {"error": str(e)}
    
    return graphics_info

def get_network_adapters():
    """
    获取网络适配器信息
    
    返回:
    - 网络适配器信息列表
    """
    network_adapters = []
    
    try:
        if platform.system() == "Windows":
            # 使用PowerShell获取网络适配器信息
            cmd = "Get-NetAdapter | Select-Object Name, InterfaceDescription, Status, MacAddress, LinkSpeed | ConvertTo-Json"
            result = run_powershell_command(cmd, timeout=15)
            
            if result.returncode == 0 and result.stdout.strip():
                adapters = json.loads(result.stdout)
                
                # 处理单适配器和多适配器的情况
                if isinstance(adapters, dict):
                    adapters = [adapters]
                
                network_adapters = adapters
                
                # 获取每个适配器的IP配置
                for adapter in network_adapters:
                    try:
                        name = adapter.get("Name", "")
                        cmd = f"Get-NetIPAddress -InterfaceAlias '{name}' | Select-Object IPAddress, PrefixLength, AddressFamily | ConvertTo-Json"
                        result = run_powershell_command(cmd, timeout=10)
                        
                        if result.returncode == 0 and result.stdout.strip():
                            ip_data = json.loads(result.stdout)
                            
                            # 处理单IP和多IP的情况
                            if not isinstance(ip_data, list):
                                ip_data = [ip_data]
                            
                            adapter["IPConfigurations"] = ip_data
                    except Exception as e:
                        print(f"获取适配器 {name} 的IP配置时出错: {e}")
        else:
            # Linux系统
            network_adapters = [{"platform_not_supported": True}]
    except Exception as e:
        print(f"获取网络适配器信息时出错: {e}")
        network_adapters = [{"error": str(e)}]
    
    return network_adapters

def get_user_accounts():
    """
    获取用户账户信息
    
    返回:
    - 用户账户信息列表
    """
    user_accounts = []
    
    try:
        if platform.system() == "Windows":
            # 使用PowerShell获取本地用户账户信息
            cmd = "Get-LocalUser | Select-Object Name, Enabled, PasswordRequired, LastLogon, Description | ConvertTo-Json"
            result = run_powershell_command(cmd, timeout=10)
            
            if result.returncode == 0 and result.stdout.strip():
                accounts = json.loads(result.stdout)
                
                # 处理单用户和多用户的情况
                if isinstance(accounts, dict):
                    accounts = [accounts]
                
                user_accounts = accounts
        else:
            # Linux系统
            user_accounts = [{"platform_not_supported": True}]
    except Exception as e:
        print(f"获取用户账户信息时出错: {e}")
        user_accounts = [{"error": str(e)}]
    
    return user_accounts

def get_installed_drivers():
    """
    获取已安装的驱动程序
    
    返回:
    - 驱动程序信息列表
    """
    drivers = []
    
    try:
        if platform.system() == "Windows":
            # 使用PowerShell获取驱动程序信息
            cmd = "Get-WmiObject Win32_PnPSignedDriver | Where-Object {$_.DeviceName} | Select-Object DeviceName, DriverVersion, Manufacturer | ConvertTo-Json -Depth 1"
            result = run_powershell_command(cmd, timeout=25)
            
            if result.returncode == 0 and result.stdout.strip():
                driver_data = json.loads(result.stdout)
                
                # 处理单驱动和多驱动的情况
                if isinstance(driver_data, dict):
                    driver_data = [driver_data]
                
                drivers = driver_data
        else:
            # Linux系统
            drivers = [{"platform_not_supported": True}]
    except Exception as e:
        print(f"获取驱动程序信息时出错: {e}")
        drivers = [{"error": str(e)}]
    
    return drivers

def get_startup_items():
    """
    获取开机启动项
    
    返回:
    - 启动项信息列表
    """
    startup_items = []
    
    try:
        if platform.system() == "Windows":
            # 使用PowerShell获取开机启动项
            cmd = "Get-CimInstance Win32_StartupCommand | Select-Object Name, Command, Location, User | ConvertTo-Json"
            result = run_powershell_command(cmd, timeout=15)
            
            if result.returncode == 0 and result.stdout.strip():
                items = json.loads(result.stdout)
                
                # 处理单项和多项的情况
                if isinstance(items, dict):
                    items = [items]
                
                startup_items = items
        else:
            # Linux系统
            startup_items = [{"platform_not_supported": True}]
    except Exception as e:
        print(f"获取开机启动项时出错: {e}")
        startup_items = [{"error": str(e)}]
    
    return startup_items

def get_scheduled_tasks():
    """
    获取计划任务
    
    返回:
    - 计划任务信息列表
    """
    scheduled_tasks = []
    
    try:
        if platform.system() == "Windows":
            # 使用PowerShell获取计划任务
            cmd = "Get-ScheduledTask | Where-Object {$_.State -ne 'Disabled'} | Select-Object TaskName, TaskPath, State | ConvertTo-Json"
            result = run_powershell_command(cmd, timeout=20)
            
            if result.returncode == 0 and result.stdout.strip():
                tasks = json.loads(result.stdout)
                
                # 处理单任务和多任务的情况
                if isinstance(tasks, dict):
                    tasks = [tasks]
                
                # 只获取前100个任务，避免数据过大
                scheduled_tasks = tasks[:100]
        else:
            # Linux系统
            scheduled_tasks = [{"platform_not_supported": True}]
    except Exception as e:
        print(f"获取计划任务时出错: {e}")
        scheduled_tasks = [{"error": str(e)}]
    
    return scheduled_tasks

def get_environment_variables():
    """
    获取环境变量
    
    返回:
    - 环境变量字典
    """
    # 直接使用os.environ获取环境变量
    # 过滤掉可能包含敏感信息的变量
    sensitive_vars = ['password', 'key', 'secret', 'token', 'credential']
    
    env_vars = {}
    for key, value in os.environ.items():
        # 跳过可能包含敏感信息的变量
        if any(s in key.lower() for s in sensitive_vars):
            continue
        env_vars[key] = value
    
    return env_vars

def collect_all_system_info():
    """
    收集所有系统信息
    
    返回:
    - 包含所有系统信息的字典
    """
    # 获取基本系统信息
    basic_info = {
        "platform": platform.platform(),
        "system": platform.system(),
        "release": platform.release(),
        "version": platform.version(),
        "architecture": platform.machine(),
        "processor": platform.processor(),
        "hostname": socket.gethostname(),
        "ip_address": socket.gethostbyname(socket.gethostname()),
        "mac_address": ':'.join(re.findall('..', '%012x' % uuid.getnode())),
        "os": get_os_info(),
        "hardware": {
            "cpu": get_cpu_info(),
            "memory": get_memory_info(),
            "disks": get_disk_info(),
            "graphics": get_graphics_info(),
            "motherboard": get_motherboard_info(),
            "bios": get_bios_info()
        }
    }
    
    system_info = {
        "collection_time": datetime.now().isoformat(),
        "basic_info": basic_info,
        "cpu": get_cpu_info(),
        "memory": get_memory_info(),
        "disk": get_disk_info(),
        "graphics": get_graphics_info(),
        "network_adapters": get_network_adapters(),
        "user_accounts": get_user_accounts(),
        "drivers": get_installed_drivers(),
        "startup_items": get_startup_items(),
        "scheduled_tasks": get_scheduled_tasks(),
        "environment_variables": get_environment_variables()
    }
    
    return system_info

def save_system_info(output_dir, filename="system_info.json"):
    """
    保存系统信息到JSON文件
    
    参数:
    - output_dir: 输出目录
    - filename: 输出文件名
    
    返回:
    - 保存结果信息
    """
    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 收集系统信息
        system_info = collect_all_system_info()
        
        # 保存系统信息
        output_file = os.path.join(output_dir, filename)
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(system_info, f, indent=2, ensure_ascii=False)
        
        return {
            "success": True,
            "message": f"系统信息已保存到 {output_file}",
            "file": output_file
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"保存系统信息时出错: {str(e)}",
            "error": str(e)
        }

def get_os_info():
    """
    获取操作系统详细信息
    
    返回:
    - 操作系统信息字典
    """
    os_info = {
        "name": platform.system(),
        "version": platform.version(),
        "release": platform.release(),
        "build": platform.win32_ver()[1],
        "architecture": platform.architecture()[0],
        "machine": platform.machine(),
        "processor": platform.processor(),
        "node": platform.node(),
        "is_64bit": platform.machine().endswith('64'),
        "system_drive": os.environ.get("SystemDrive", "C:"),
        "windows_dir": os.environ.get("windir", ""),
        "installation_date": get_windows_install_date()
    }
    
    return os_info

def get_windows_install_date():
    """
    获取Windows安装日期
    
    返回:
    - 安装日期字符串
    """
    try:
        with winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, 
                           r"SOFTWARE\Microsoft\Windows NT\CurrentVersion") as key:
            install_date = winreg.QueryValueEx(key, "InstallDate")[0]
            return datetime.fromtimestamp(install_date).isoformat()
    except Exception as e:
        return f"Error getting install date: {str(e)}"

def get_hardware_info():
    """
    获取硬件配置信息
    
    返回:
    - 硬件信息字典
    """
    hardware_info = {
        "cpu": get_cpu_info(),
        "memory": get_memory_info(),
        "disks": get_disk_info(),
        "graphics": get_graphics_info(),
        "motherboard": get_motherboard_info(),
        "bios": get_bios_info()
    }
    
    return hardware_info

def get_motherboard_info():
    """
    获取主板信息
    
    返回:
    - 主板信息字典
    """
    motherboard_info = {}
    
    try:
        # 使用PowerShell代替WMI直接调用
        logging.debug("Getting motherboard information")
        ps_cmd = "$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'; [System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'; [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'; Get-WmiObject Win32_BaseBoard | Select-Object Manufacturer, Product, SerialNumber, Version | ConvertTo-Json"
        result = run_powershell_command(ps_cmd, timeout=15)
        
        if result.returncode == 0 and result.stdout.strip():
            try:
                mb_data = json.loads(result.stdout)
                motherboard_info = mb_data
                logging.debug("Successfully retrieved motherboard information")
            except json.JSONDecodeError as e:
                logging.warning(f"Error parsing motherboard JSON: {e}, output: {result.stdout[:100]}")
        else:
            # 尝试使用备用方法
            logging.warning(f"Failed to get motherboard info: {result.stderr}")
        
        # 如果PowerShell失败，尝试使用WMIC
        if not motherboard_info:
            logging.debug("Trying WMIC as fallback for motherboard info")
            wmic_result = run_wmic_command("baseboard", ["Manufacturer", "Product", "SerialNumber", "Version"])
            
            if wmic_result.returncode == 0 and wmic_result.stdout.strip():
                try:
                    # 解析WMIC CSV输出
                    lines = wmic_result.stdout.strip().split('\n')
                    
                    if len(lines) > 1:  # 跳过标题行
                        for line in lines[1:]:
                            if line.strip() and not line.startswith('Node,'):
                                fields = line.split(',')
                                if len(fields) >= 5:  # Node, Manufacturer, Product, SerialNumber, Version
                                    motherboard_info = {
                                        "Manufacturer": fields[1].strip() if len(fields) > 1 else "",
                                        "Product": fields[2].strip() if len(fields) > 2 else "",
                                        "SerialNumber": fields[3].strip() if len(fields) > 3 else "",
                                        "Version": fields[4].strip() if len(fields) > 4 else ""
                                    }
                                    logging.debug("Successfully retrieved motherboard information via WMIC")
                                    break
                except Exception as e:
                    logging.warning(f"Error parsing WMIC motherboard output: {e}")
        
        # 如果仍然失败，尝试简单的解析方法作为后备
        if not motherboard_info:
            logging.debug("Trying simple PowerShell format as final fallback")
            cmd = "Get-WmiObject Win32_BaseBoard | Format-List Manufacturer, Product, SerialNumber, Version"
            result = run_powershell_command(cmd, timeout=10)
            
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    if ':' in line:
                        key, value = line.split(':', 1)
                        motherboard_info[key.strip()] = value.strip()
                        
        # 如果还是没有数据，提供基本信息
        if not motherboard_info:
            motherboard_info = {
                "Manufacturer": "Unknown",
                "Product": "Unknown", 
                "SerialNumber": "Unknown",
                "Version": "Unknown"
            }
            logging.debug("Using unknown motherboard info as final fallback")
            
    except Exception as e:
        logging.error(f"Error getting motherboard information: {e}", exc_info=True)
        motherboard_info["error"] = str(e)
    
    return motherboard_info

def get_bios_info():
    """
    获取BIOS信息
    
    返回:
    - BIOS信息字典
    """
    bios_info = {}
    
    try:
        # 使用PowerShell代替WMI直接调用
        logging.debug("Getting BIOS information")
        ps_cmd = "$PSDefaultParameterValues['Out-File:Encoding'] = 'utf8'; [System.Threading.Thread]::CurrentThread.CurrentCulture = 'en-US'; [System.Threading.Thread]::CurrentThread.CurrentUICulture = 'en-US'; Get-WmiObject Win32_BIOS | Select-Object Manufacturer, Name, SMBIOSBIOSVersion, ReleaseDate | ConvertTo-Json"
        result = run_powershell_command(ps_cmd, timeout=15)
        
        if result.returncode == 0 and result.stdout.strip():
            try:
                bios_data = json.loads(result.stdout)
                bios_info = bios_data
                
                # 处理BIOS日期格式
                if "ReleaseDate" in bios_info and bios_info["ReleaseDate"]:
                    # WMI返回的日期格式通常是 YYYYMMDDHHMMSS.MMMMMM+MMM
                    release_date = bios_info["ReleaseDate"]
                    try:
                        if len(release_date) >= 8:
                            # 确保前8个字符都是数字
                            date_part = release_date[:8]
                            if date_part.isdigit():
                                formatted_date = f"{date_part[0:4]}-{date_part[4:6]}-{date_part[6:8]}"
                                bios_info["ReleaseDate"] = formatted_date
                    except:
                        pass  # 保持原格式
                
                logging.debug("Successfully retrieved BIOS information")
            except json.JSONDecodeError as e:
                logging.warning(f"Error parsing BIOS JSON: {e}, output: {result.stdout[:100]}")
        else:
            # 尝试使用备用方法
            logging.warning(f"Failed to get BIOS info: {result.stderr}")
        
        # 如果PowerShell失败，尝试使用WMIC
        if not bios_info:
            logging.debug("Trying WMIC as fallback for BIOS info")
            wmic_result = run_wmic_command("bios", ["Manufacturer", "Name", "SMBIOSBIOSVersion", "ReleaseDate"])
            
            if wmic_result.returncode == 0 and wmic_result.stdout.strip():
                try:
                    # 解析WMIC CSV输出
                    lines = wmic_result.stdout.strip().split('\n')
                    
                    if len(lines) > 1:  # 跳过标题行
                        for line in lines[1:]:
                            if line.strip() and not line.startswith('Node,'):
                                fields = line.split(',')
                                if len(fields) >= 5:  # Node, Manufacturer, Name, ReleaseDate, SMBIOSBIOSVersion
                                    bios_info = {
                                        "Manufacturer": fields[1].strip() if len(fields) > 1 else "",
                                        "Name": fields[2].strip() if len(fields) > 2 else "",
                                        "ReleaseDate": fields[3].strip() if len(fields) > 3 else "",
                                        "SMBIOSBIOSVersion": fields[4].strip() if len(fields) > 4 else ""
                                    }
                                    
                                    # 处理BIOS日期格式
                                    if bios_info["ReleaseDate"] and len(bios_info["ReleaseDate"]) >= 8:
                                        try:
                                            release_date = bios_info["ReleaseDate"]
                                            # 确保前8个字符都是数字
                                            date_part = release_date[:8]
                                            if date_part.isdigit():
                                                formatted_date = f"{date_part[0:4]}-{date_part[4:6]}-{date_part[6:8]}"
                                                bios_info["ReleaseDate"] = formatted_date
                                        except:
                                            pass  # 保持原格式
                                    
                                    logging.debug("Successfully retrieved BIOS information via WMIC")
                                    break
                except Exception as e:
                    logging.warning(f"Error parsing WMIC BIOS output: {e}")
        
        # 如果仍然失败，尝试简单的解析方法作为后备
        if not bios_info:
            logging.debug("Trying simple PowerShell format as final fallback")
            cmd = "Get-WmiObject Win32_BIOS | Format-List Manufacturer, Name, SMBIOSBIOSVersion, ReleaseDate"
            result = run_powershell_command(cmd, timeout=10)
            
            if result.returncode == 0:
                for line in result.stdout.split('\n'):
                    if ':' in line:
                        key, value = line.split(':', 1)
                        bios_info[key.strip()] = value.strip()
        
        # 如果还是没有数据，提供基本信息
        if not bios_info:
            bios_info = {
                "Manufacturer": "Unknown",
                "Name": "Unknown", 
                "SMBIOSBIOSVersion": "Unknown",
                "ReleaseDate": "Unknown"
            }
            logging.debug("Using unknown BIOS info as final fallback")
            
    except Exception as e:
        logging.error(f"Error getting BIOS information: {e}", exc_info=True)
        bios_info["error"] = str(e)
    
    return bios_info

def get_network_info():
    """
    获取网络配置信息
    
    返回:
    - 网络配置信息字典
    """
    network_info = {
        "hostname": socket.gethostname(),
        "interfaces": get_network_interfaces(),
        "dns_servers": get_dns_servers(),
        "ip_config": get_ip_config()
    }
    
    return network_info

def get_network_interfaces():
    """
    获取网络接口信息
    
    返回:
    - 网络接口信息列表
    """
    interfaces = []
    
    try:
        # 使用ipconfig命令获取网络接口信息
        result = subprocess.check_output('ipconfig /all', shell=True, text=True, encoding='gbk')
        
        # 解析ipconfig输出
        sections = re.split(r'\r?\n\r?\n', result)
        for section in sections:
            if not section.strip():
                continue
                
            lines = section.split('\n')
            if len(lines) > 0:
                # 第一行包含接口名称
                interface_name = lines[0].strip(':')
                interface = {"name": interface_name.strip()}
                
                for line in lines[1:]:
                    line = line.strip()
                    if ':' in line:
                        key, value = line.split(':', 1)
                        interface[key.strip()] = value.strip()
                
                interfaces.append(interface)
    except Exception as e:
        interfaces = [{"error": str(e)}]
    
    return interfaces

def get_dns_servers():
    """
    获取DNS服务器信息
    
    返回:
    - DNS服务器列表
    """
    dns_servers = []
    
    try:
        # 使用ipconfig命令获取DNS服务器信息
        result = subprocess.check_output('ipconfig /all', shell=True, text=True, encoding='gbk')
        
        # 查找DNS服务器条目
        dns_lines = re.findall(r'DNS Servers[\.\s]+: (.*)', result)
        for line in dns_lines:
            dns_servers.append(line.strip())
            
        # 查找后续的DNS服务器条目（同一网卡的多个DNS）
        dns_continuation_lines = re.findall(r'(?<=DNS Servers)[\.\s]+: (.*)', result)
        for line in dns_continuation_lines:
            if line.strip() not in dns_servers:
                dns_servers.append(line.strip())
    except Exception as e:
        dns_servers = [f"Error: {str(e)}"]
    
    return dns_servers

def get_ip_config():
    """
    获取IP配置信息
    
    返回:
    - IP配置信息字典
    """
    ip_config = {}
    
    try:
        # 执行ipconfig命令
        result = subprocess.check_output('ipconfig /all', shell=True, text=True, encoding='gbk')
        ip_config["raw_output"] = result
        
        # 提取主要信息
        ip_addresses = re.findall(r'IPv4 Address[\.\s]+: (.*)', result)
        ip_config["ipv4_addresses"] = [ip.strip() for ip in ip_addresses]
        
        subnet_masks = re.findall(r'Subnet Mask[\.\s]+: (.*)', result)
        ip_config["subnet_masks"] = [mask.strip() for mask in subnet_masks]
        
        default_gateways = re.findall(r'Default Gateway[\.\s]+: (.*)', result)
        ip_config["default_gateways"] = [gw.strip() for gw in default_gateways if gw.strip()]
        
        mac_addresses = re.findall(r'Physical Address[\.\s]+: (.*)', result)
        ip_config["mac_addresses"] = [mac.strip() for mac in mac_addresses]
    except Exception as e:
        ip_config["error"] = str(e)
    
    return ip_config

def get_shared_folders():
    """
    获取共享文件夹信息
    
    返回:
    - 共享文件夹信息列表
    """
    shared_folders = []
    
    try:
        # 使用net share命令获取共享文件夹
        result = subprocess.check_output('net share', shell=True, text=True, encoding='gbk')
        
        # 解析共享文件夹列表
        lines = result.split('\n')
        capture = False
        
        for i, line in enumerate(lines):
            if '----------' in line:
                capture = True
                continue
            
            if capture and line.strip():
                parts = re.split(r'\s{2,}', line.strip())
                if len(parts) >= 2:
                    share_name = parts[0].strip()
                    share_path = parts[1].strip()
                    
                    # 跳过系统默认共享
                    if share_name in ['C$', 'D$', 'E$', 'F$', 'ADMIN$', 'IPC$'] or '$' in share_name:
                        continue
                    
                    shared_folders.append({
                        "name": share_name,
                        "path": share_path
                    })
    except Exception as e:
        shared_folders = [{"error": str(e)}]
    
    return shared_folders

def get_services():
    """
    获取系统服务信息
    
    返回:
    - 系统服务信息列表
    """
    services = []
    
    try:
        # 使用net start命令获取运行中的服务
        result = subprocess.check_output('net start', shell=True, text=True, encoding='gbk')
        
        # 解析服务列表
        lines = result.split('\n')
        started_services = []
        
        capture = False
        for line in lines:
            if line.strip() == "以下服务已经启动:" or line.strip() == "The following services are started:":
                capture = True
                continue
            
            if "命令成功完成" in line or "Command completed successfully" in line:
                capture = False
                continue
                
            if capture and line.strip():
                started_services.append(line.strip())
        
        # 使用wmic获取所有服务的详细信息
        wmic_cmd = 'wmic service get Caption, DisplayName, Name, PathName, StartMode, State /format:list'
        result = subprocess.check_output(wmic_cmd, shell=True, text=True)
        
        # 解析服务详细信息
        service_sections = result.split('\n\n')
        for section in service_sections:
            if not section.strip():
                continue
                
            service = {}
            for line in section.split('\n'):
                if '=' in line:
                    key, value = line.split('=', 1)
                    service[key.strip()] = value.strip()
            
            if service:
                service["running"] = service.get("State") == "Running"
                service["auto_start"] = service.get("StartMode") == "Auto"
                services.append(service)
    except Exception as e:
        services = [{"error": str(e)}]
    
    return services

def get_installed_software():
    """
    获取已安装软件信息
    
    返回:
    - 已安装软件列表
    """
    software_list = []
    
    try:
        # 从注册表读取已安装软件信息
        registry_locations = [
            (winreg.HKEY_LOCAL_MACHINE, r"Software\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_LOCAL_MACHINE, r"Software\WOW6432Node\Microsoft\Windows\CurrentVersion\Uninstall"),
            (winreg.HKEY_CURRENT_USER, r"Software\Microsoft\Windows\CurrentVersion\Uninstall")
        ]
        
        for hkey, key_path in registry_locations:
            try:
                with winreg.OpenKey(hkey, key_path) as key:
                    # 获取子键数量
                    subkey_count = winreg.QueryInfoKey(key)[0]
                    
                    for i in range(subkey_count):
                        try:
                            subkey_name = winreg.EnumKey(key, i)
                            with winreg.OpenKey(key, subkey_name) as subkey:
                                software = {}
                                
                                # 读取软件信息
                                try:
                                    software["name"] = winreg.QueryValueEx(subkey, "DisplayName")[0]
                                except:
                                    # 没有显示名称的条目跳过
                                    continue
                                
                                try:
                                    software["version"] = winreg.QueryValueEx(subkey, "DisplayVersion")[0]
                                except:
                                    software["version"] = ""
                                
                                try:
                                    software["publisher"] = winreg.QueryValueEx(subkey, "Publisher")[0]
                                except:
                                    software["publisher"] = ""
                                
                                try:
                                    software["install_date"] = winreg.QueryValueEx(subkey, "InstallDate")[0]
                                except:
                                    software["install_date"] = ""
                                
                                try:
                                    software["install_location"] = winreg.QueryValueEx(subkey, "InstallLocation")[0]
                                except:
                                    software["install_location"] = ""
                                
                                try:
                                    software["uninstall_string"] = winreg.QueryValueEx(subkey, "UninstallString")[0]
                                except:
                                    software["uninstall_string"] = ""
                                
                                software_list.append(software)
                        except Exception as e:
                            continue
            except Exception as e:
                continue
    except Exception as e:
        software_list = [{"error": str(e)}]
    
    return software_list

def get_system_info():
    """
    获取系统全面信息
    
    返回:
    - 系统信息字典
    """
    system_info = {
        "collection_time": datetime.now().isoformat(),
        "os": get_os_info(),
        "hardware": get_hardware_info(),
        "network": get_network_info(),
        "user_accounts": get_user_accounts(),
        "startup_items": get_startup_items(),
        "installed_drivers": get_installed_drivers(),
        "environment_variables": get_environment_variables(),
        "shared_folders": get_shared_folders(),
        "scheduled_tasks": get_scheduled_tasks(),
        "services": get_services(),
        "installed_software": get_installed_software()
    }
    
    return system_info

def save_system_info(output_dir, filename="system_info.json"):
    """
    保存系统信息到JSON文件
    
    参数:
    - output_dir: 输出目录
    - filename: 输出文件名
    
    返回:
    - 保存文件的路径
    """
    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 收集系统信息
        system_info = get_system_info()
        
        # 保存到JSON文件
        output_file = os.path.join(output_dir, filename)
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(system_info, f, ensure_ascii=False, indent=2)
        
        return output_file
    except Exception as e:
        print(f"保存系统信息时出错: {e}")
        return None 