#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
测试WMI备用方案的脚本
验证在PowerShell 7+环境下，Get-WmiObject命令是否能正确转换为Get-CimInstance
"""

import sys
import logging
from modules.collectors.system_info_collector import (
    run_powershell_command, 
    run_wmic_command,
    get_cpu_info,
    get_graphics_info,
    get_motherboard_info,
    get_bios_info
)

def setup_test_logging():
    """设置测试日志"""
    logging.basicConfig(
        level=logging.DEBUG,
        format='%(asctime)s - %(levelname)s - %(message)s',
        handlers=[
            logging.StreamHandler(),
            logging.FileHandler('test_wmi_fallback.log', encoding='utf-8')
        ]
    )

def test_powershell_wmi_conversion():
    """测试PowerShell WMI命令转换"""
    print("=" * 60)
    print("测试 PowerShell WMI 命令转换")
    print("=" * 60)
    
    # 测试原始的Get-WmiObject命令
    test_command = "Get-WmiObject -Class Win32_Processor | Select-Object Name, NumberOfCores | ConvertTo-Json"
    
    print(f"原始命令: {test_command}")
    result = run_powershell_command(test_command, timeout=10)
    
    print(f"返回码: {result.returncode}")
    if result.returncode == 0:
        print(f"成功输出: {result.stdout[:200]}...")
    else:
        print(f"错误输出: {result.stderr[:200]}...")
    
    return result.returncode == 0

def main():
    """主函数"""
    print("WMI备用方案测试脚本")
    print("用于验证在PowerShell 7+环境下WMI命令是否能正常工作")
    
    # 设置日志
    setup_test_logging()
    
    # 运行测试
    ps_test = test_powershell_wmi_conversion()
    print(f"PowerShell测试结果: {'通过' if ps_test else '失败'}")

if __name__ == "__main__":
    main() 