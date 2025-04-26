import datetime
import json
import os
from pathlib import Path

def generate_html_report(data, output_path):
    """
    根据收集的系统信息生成HTML报告
    
    参数:
    - data: 包含系统信息的字典
    - output_path: 输出HTML文件的路径
    """
    system_info = data.get("system_info", {})
    installed_apps = data.get("installed_apps", [])
    startup_items = data.get("startup_items", [])
    
    # 生成HTML内容
    html_content = f"""<!DOCTYPE html>
<html lang="zh-CN">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>系统信息报告</title>
    <style>
        body {{ font-family: Arial, sans-serif; margin: 20px; }}
        h1, h2, h3 {{ color: #333; }}
        table {{ border-collapse: collapse; width: 100%; margin-bottom: 20px; }}
        th, td {{ text-align: left; padding: 8px; border-bottom: 1px solid #ddd; }}
        th {{ background-color: #4CAF50; color: white; }}
        tr:nth-child(even) {{ background-color: #f2f2f2; }}
        .container {{ max-width: 1200px; margin: 0 auto; }}
        .timestamp {{ color: #666; font-style: italic; margin-bottom: 20px; }}
        .section {{ margin-bottom: 30px; }}
        .info-box {{ background-color: #f9f9f9; border: 1px solid #ddd; padding: 15px; margin-bottom: 20px; }}
        .success {{ color: green; }}
        .error {{ color: red; }}
    </style>
</head>
<body>
    <div class="container">
        <h1>Windows系统信息报告</h1>
        <p class="timestamp">生成时间: {datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")}</p>
        
        <div class="section">
            <h2>系统信息</h2>
            <div class="info-box">
                <table>
                    <tr>
                        <th>项目</th>
                        <th>值</th>
                    </tr>
                    <tr>
                        <td>操作系统</td>
                        <td>{system_info.get('os_version', '未知')}</td>
                    </tr>
                    <tr>
                        <td>计算机名</td>
                        <td>{system_info.get('hostname', '未知')}</td>
                    </tr>
                    <tr>
                        <td>处理器</td>
                        <td>{system_info.get('processor', '未知')}</td>
                    </tr>
                    <tr>
                        <td>内存</td>
                        <td>{system_info.get('memory', '未知')}</td>
                    </tr>
                </table>
            </div>
        </div>
        
        <div class="section">
            <h2>已安装的应用程序 (前20个)</h2>
            <div class="info-box">
                <table>
                    <tr>
                        <th>名称</th>
                        <th>版本</th>
                        <th>发布者</th>
                    </tr>
                    {"".join([f'''
                    <tr>
                        <td>{app.get('name', '')}</td>
                        <td>{app.get('version', '')}</td>
                        <td>{app.get('publisher', '')}</td>
                    </tr>''' for app in installed_apps[:20]])}
                </table>
                <p>总共 {len(installed_apps)} 个应用程序</p>
            </div>
        </div>
        
        <div class="section">
            <h2>启动项</h2>
            <div class="info-box">
                <table>
                    <tr>
                        <th>名称</th>
                        <th>命令</th>
                        <th>位置</th>
                    </tr>
                    {"".join([f'''
                    <tr>
                        <td>{item.get('name', '')}</td>
                        <td>{item.get('command', '')}</td>
                        <td>{item.get('location', '')}</td>
                    </tr>''' for item in startup_items])}
                </table>
            </div>
        </div>
        
        <div class="section">
            <h2>恢复指南</h2>
            <div class="info-box">
                <h3>系统重新安装后的步骤:</h3>
                <ol>
                    <li>安装操作系统后，首先安装必要的驱动程序</li>
                    <li>使用驱动程序恢复工具从备份文件夹中恢复驱动程序</li>
                    <li>恢复网络设置，包括Wi-Fi配置文件</li>
                    <li>根据应用程序列表重新安装必要的软件</li>
                    <li>检查启动项列表，配置需要自动启动的程序</li>
                </ol>
                
                <h3>驱动程序恢复命令:</h3>
                <pre>pnputil /add-driver "驱动程序备份路径\\*.inf" /subdirs /install</pre>
                
                <h3>Wi-Fi配置恢复命令:</h3>
                <pre>netsh wlan add profile filename="Wi-Fi配置文件路径"</pre>
            </div>
        </div>
    </div>
</body>
</html>
"""
    
    # 写入HTML文件
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    
    return {
        "success": True,
        "message": f"HTML报告已生成: {output_path}"
    }

def load_json_data(json_dir):
    """
    从JSON文件目录加载所有数据
    
    参数:
    - json_dir: 包含JSON数据文件的目录
    """
    json_dir = Path(json_dir)
    data = {}
    
    # 加载系统信息
    system_info_path = json_dir / "system_info.json"
    if system_info_path.exists():
        with open(system_info_path, 'r', encoding='utf-8') as f:
            data["system_info"] = json.load(f)
    
    # 加载已安装应用程序列表
    apps_path = json_dir / "installed_apps.json"
    if apps_path.exists():
        with open(apps_path, 'r', encoding='utf-8') as f:
            data["installed_apps"] = json.load(f)
    
    # 加载启动项列表
    startup_path = json_dir / "startup_items.json"
    if startup_path.exists():
        with open(startup_path, 'r', encoding='utf-8') as f:
            data["startup_items"] = json.load(f)
    
    return data

def generate_report_from_directory(json_dir, output_path=None):
    """
    从包含JSON数据文件的目录生成HTML报告
    
    参数:
    - json_dir: 包含JSON数据文件的目录
    - output_path: 输出HTML文件的路径，如果未提供则在json_dir中创建
    """
    json_dir = Path(json_dir)
    
    if not output_path:
        output_path = json_dir / "system_report.html"
    
    # 加载数据
    data = load_json_data(json_dir)
    
    # 生成报告
    return generate_html_report(data, output_path) 