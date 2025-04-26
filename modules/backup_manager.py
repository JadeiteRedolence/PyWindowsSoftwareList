import os
import datetime
import json
from pathlib import Path

# 导入收集器模块
from .collectors.software_collector import get_all_installed_software
from .collectors.system_info_collector import collect_all_system_info
from .collectors.driver_collector import backup_drivers, list_drivers
from .collectors.network_backup import backup_wifi_profiles, backup_network_settings, backup_wired_profiles
from .collectors.dev_env_collector import save_dev_environment_info

# 导入导出器模块
from .exporters.html_report_exporter import generate_report_from_directory

class BackupManager:
    """系统信息和设置备份管理器"""
    
    def __init__(self, base_dir="Report"):
        """
        初始化备份管理器
        
        参数:
        - base_dir: 备份基础目录，默认为"Report"
        """
        self.base_dir = Path(base_dir)
        self.timestamp = datetime.datetime.now().strftime("%Y%m%d_%H%M%S")
        self.output_dir = self.base_dir / self.timestamp
        self.summary = {
            "timestamp": self.timestamp,
            "output_dir": str(self.output_dir),
            "steps": []
        }
    
    def create_output_dir(self):
        """创建输出目录"""
        os.makedirs(self.output_dir, exist_ok=True)
        return self.output_dir
    
    def backup_software_list(self):
        """备份已安装的软件清单"""
        print("收集软件清单...")
        
        try:
            software_list = get_all_installed_software()
            
            # 将软件列表保存为JSON文件
            software_file = self.output_dir / "software_list.json"
            with open(software_file, 'w', encoding='utf-8') as f:
                json.dump(software_list, f, indent=2, ensure_ascii=False)
            
            result = {
                "step": "软件清单备份",
                "status": "成功",
                "message": f"找到 {len(software_list)} 个软件项",
                "file": str(software_file)
            }
        except Exception as e:
            result = {
                "step": "软件清单备份",
                "status": "失败",
                "error": str(e)
            }
        
        self.summary["steps"].append(result)
        print(f"软件清单备份: {result['status']}")
        return result
    
    def backup_system_info(self):
        """备份系统信息"""
        print("收集系统信息...")
        
        try:
            info = collect_all_system_info(self.output_dir)
            
            result = {
                "step": "系统信息备份",
                "status": "成功",
                "message": "系统信息已保存",
                "details": info
            }
        except Exception as e:
            result = {
                "step": "系统信息备份",
                "status": "失败",
                "error": str(e)
            }
        
        self.summary["steps"].append(result)
        print(f"系统信息备份: {result['status']}")
        return result
    
    def backup_all_drivers(self):
        """备份所有驱动程序"""
        print("备份驱动程序...")
        
        try:
            result = backup_drivers(self.output_dir)
            
            step_result = {
                "step": "驱动程序备份",
                "status": "成功" if result["success"] else "失败",
                "message": result["message"]
            }
            
            if not result["success"] and "error" in result:
                step_result["error"] = result["error"]
                
        except Exception as e:
            step_result = {
                "step": "驱动程序备份",
                "status": "失败",
                "error": str(e)
            }
        
        self.summary["steps"].append(step_result)
        print(f"驱动程序备份: {step_result['status']}")
        return step_result
    
    def backup_network_configs(self):
        """备份网络配置"""
        print("备份网络配置...")
        
        results = []
        
        # 备份Wi-Fi配置文件
        try:
            wifi_result = backup_wifi_profiles(self.output_dir)
            results.append({
                "type": "Wi-Fi配置",
                "status": "成功" if wifi_result["success"] else "失败",
                "message": wifi_result["message"]
            })
        except Exception as e:
            results.append({
                "type": "Wi-Fi配置",
                "status": "失败",
                "error": str(e)
            })
        
        # 备份TCP/IP配置
        try:
            network_result = backup_network_settings(self.output_dir)
            results.append({
                "type": "TCP/IP配置",
                "status": "成功" if network_result["success"] else "失败",
                "message": network_result["message"]
            })
        except Exception as e:
            results.append({
                "type": "TCP/IP配置",
                "status": "失败",
                "error": str(e)
            })
        
        # 备份有线网络配置
        try:
            wired_result = backup_wired_profiles(self.output_dir)
            results.append({
                "type": "有线网络配置",
                "status": "成功" if wired_result["success"] else "失败",
                "message": wired_result["message"]
            })
        except Exception as e:
            results.append({
                "type": "有线网络配置",
                "status": "失败",
                "error": str(e)
            })
        
        # 汇总结果
        overall_status = all(r["status"] == "成功" for r in results)
        
        step_result = {
            "step": "网络配置备份",
            "status": "成功" if overall_status else "部分成功",
            "details": results
        }
        
        self.summary["steps"].append(step_result)
        print(f"网络配置备份: {step_result['status']}")
        return step_result
    
    def backup_dev_environment(self):
        """备份开发环境信息"""
        print("收集开发环境信息...")
        
        try:
            result = save_dev_environment_info(self.output_dir)
            
            step_result = {
                "step": "开发环境备份",
                "status": "成功" if result["success"] else "失败",
                "message": result["message"]
            }
            
            if not result["success"] and "error" in result:
                step_result["error"] = result["error"]
                
        except Exception as e:
            step_result = {
                "step": "开发环境备份",
                "status": "失败",
                "error": str(e)
            }
        
        self.summary["steps"].append(step_result)
        print(f"开发环境备份: {step_result['status']}")
        return step_result
    
    def generate_html_report(self):
        """生成HTML报告"""
        print("生成HTML报告...")
        
        try:
            html_path = self.output_dir / "system_report.html"
            result = generate_report_from_directory(self.output_dir, html_path)
            
            step_result = {
                "step": "HTML报告生成",
                "status": "成功" if result["success"] else "失败",
                "message": result["message"],
                "file": str(html_path)
            }
        except Exception as e:
            step_result = {
                "step": "HTML报告生成",
                "status": "失败",
                "error": str(e)
            }
        
        self.summary["steps"].append(step_result)
        print(f"HTML报告生成: {step_result['status']}")
        return step_result
    
    def save_summary(self):
        """保存备份摘要信息"""
        summary_path = self.output_dir / "backup_summary.json"
        
        with open(summary_path, 'w', encoding='utf-8') as f:
            json.dump(self.summary, f, indent=2, ensure_ascii=False)
        
        self.summary["summary_file"] = str(summary_path)
        return summary_path
    
    def backup_all(self):
        """执行所有备份操作"""
        print(f"开始全面系统备份，时间戳: {self.timestamp}")
        print(f"备份目录: {self.output_dir}")
        
        try:
            # 创建输出目录
            self.create_output_dir()
            
            # 执行各种备份操作
            self.backup_software_list()
            self.backup_system_info()
            self.backup_all_drivers()
            self.backup_network_configs()
            self.backup_dev_environment()
            
            # 生成HTML报告
            self.generate_html_report()
            
            # 保存摘要信息
            summary_path = self.save_summary()
            
            print("备份完成!")
            print(f"备份摘要保存在: {summary_path}")
            print(f"HTML报告保存在: {self.output_dir / 'system_report.html'}")
            
            return {
                "success": True,
                "message": "备份操作全部完成",
                "summary": self.summary
            }
            
        except Exception as e:
            print(f"备份过程中出现错误: {str(e)}")
            
            # 尝试保存已收集的信息
            try:
                self.save_summary()
            except:
                pass
                
            return {
                "success": False,
                "message": "备份过程中出现错误",
                "error": str(e),
                "summary": self.summary
            } 