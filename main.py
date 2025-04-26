import os
import sys
import json
import datetime
import shutil
import logging
from pathlib import Path

# Import modules
from modules.collectors.software_collector import get_all_installed_software
from modules.collectors.system_info_collector import collect_all_system_info, save_system_info
from modules.collectors.dev_env_collector import collect_all_dev_environment_info
from modules.exporters.exporter_manager import export_all_formats
from modules.config import get_output_directory


def setup_logging(base_output_dir):
    """Setup logging to both console and file"""
    # Create logs directory
    logs_dir = os.path.join(base_output_dir, "logs")
    os.makedirs(logs_dir, exist_ok=True)
    
    # Setup logging format
    log_format = '%(asctime)s - %(levelname)s - %(message)s'
    date_format = '%Y-%m-%d %H:%M:%S'
    
    # Create a logger
    logger = logging.getLogger()
    logger.setLevel(logging.DEBUG)
    
    # Clear any existing handlers
    for handler in logger.handlers[:]:
        logger.removeHandler(handler)
    
    # Create file handler for debug messages
    log_file = os.path.join(logs_dir, f"debug_{datetime.datetime.now().strftime('%Y%m%d_%H%M%S')}.log")
    file_handler = logging.FileHandler(log_file, encoding='utf-8')
    file_handler.setLevel(logging.DEBUG)
    file_handler.setFormatter(logging.Formatter(log_format, date_format))
    
    # Create console handler for info messages
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    console_handler.setFormatter(logging.Formatter(log_format, date_format))
    
    # Add handlers to logger
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    logging.info(f"Logging initialized. Log file: {log_file}")
    return logger


def main():
    """Main function to collect system information and export to different formats"""
    try:
        # Get output directory
        base_output_dir = get_output_directory()
        
        # Create output directory
        os.makedirs(base_output_dir, exist_ok=True)
        
        # Setup logging
        logger = setup_logging(base_output_dir)
        
        # Create subdirectories for different file types
        json_dir = base_output_dir / "json"
        html_dir = base_output_dir / "html"
        excel_dir = base_output_dir / "excel"
        markdown_dir = base_output_dir / "markdown"
        reports_dir = base_output_dir / "reports"
        
        # Create all subdirectories
        for dir_path in [json_dir, html_dir, excel_dir, markdown_dir, reports_dir]:
            os.makedirs(dir_path, exist_ok=True)
            logging.debug(f"Created directory: {dir_path}")
        
        # Collect all data
        logging.info("Collecting system information...")
        
        # 1. Get software list
        logging.info("Gathering installed software information...")
        software_list = get_all_installed_software()
        logging.info(f"Found {len(software_list)} software items")
        logging.debug(f"Software list first 5 items: {software_list[:5]}")
        
        # 2. Get system information
        logging.info("Gathering system specifications...")
        system_info = collect_all_system_info()
        
        # Save system info to file
        system_info_path = os.path.join(json_dir, "system_info.json")
        with open(system_info_path, 'w', encoding='utf-8') as f:
            json.dump(system_info, f, indent=2, ensure_ascii=False)
        logging.info(f"System information collected and saved to {system_info_path}")
        
        # 3. Get development environment information
        logging.info("Gathering development environment information...")
        dev_info = collect_all_dev_environment_info()
        logging.debug(f"Collected development environment information")
        
        # Save collected data to JSON directory
        with open(json_dir / "software_list.json", 'w', encoding='utf-8') as f:
            json.dump(software_list, f, indent=2, ensure_ascii=False)
            
        with open(json_dir / "dev_environment.json", 'w', encoding='utf-8') as f:
            json.dump(dev_info, f, indent=2, ensure_ascii=False)
        
        # Create comprehensive report data
        report_data = {
            "timestamp": datetime.datetime.now().isoformat(),
            "software_list": software_list,
            "system_info": system_info,
            "dev_environment": dev_info
        }
        
        # Save comprehensive report data
        with open(json_dir / "complete_report_data.json", 'w', encoding='utf-8') as f:
            json.dump(report_data, f, indent=2, ensure_ascii=False)
        
        # Export to different formats
        logging.info("\nExporting to different formats...")
        
        # Modify export_all_formats to use our directory structure
        export_custom_formats(software_list, base_output_dir)
        
        # Create HTML report with comprehensive data
        logging.info("Generating HTML report with all collected data...")
        create_comprehensive_html_report(report_data, reports_dir / "system_report.html")
        
        logging.info(f"\nExport complete. Files saved to {base_output_dir}")
        
        # Create an index.html file in the base directory that links to all reports
        create_index_file(base_output_dir)
        
    except Exception as e:
        logging.error(f"Error: {e}", exc_info=True)
        return 1
    
    return 0


def export_custom_formats(software_list, base_output_dir):
    """Export software list to different formats using custom directory structure"""
    from modules.exporters.json_exporter import export_to_json
    from modules.exporters.html_exporter import export_to_html
    from modules.exporters.markdown_exporter import export_to_markdown
    from modules.exporters.excel_exporter import export_to_excel
    
    # Export to JSON (already done in main function)
    # export_to_json(software_list, base_output_dir / "json" / "software_list.json")
    
    # Export to HTML
    try:
        logging.debug("Exporting to HTML format...")
        export_to_html(software_list, base_output_dir / "html" / "software_list.html")
        logging.debug("HTML export complete")
    except Exception as e:
        logging.error(f"Error exporting to HTML: {e}", exc_info=True)
    
    # Export to Markdown if tabulate is available
    try:
        logging.debug("Exporting to Markdown format...")
        export_to_markdown(software_list, base_output_dir / "markdown" / "software_list.md")
        logging.debug("Markdown export complete")
    except ImportError:
        logging.warning("Missing 'tabulate' package. Markdown export skipped.")
    except Exception as e:
        logging.error(f"Error exporting to Markdown: {e}", exc_info=True)
    
    # Export to Excel
    try:
        logging.debug("Exporting to Excel format...")
        export_to_excel(software_list, base_output_dir / "excel" / "software_list.xlsx")
        logging.debug("Excel export complete")
    except Exception as e:
        logging.error(f"Error exporting to Excel: {e}", exc_info=True)


def create_comprehensive_html_report(data, output_path):
    """Create an HTML report that includes all collected data"""
    logging.debug(f"Starting to create comprehensive HTML report at {output_path}")
    
    # Extract data
    software_list = data.get("software_list", [])
    system_info = data.get("system_info", {})
    dev_environment = data.get("dev_environment", {})
    timestamp = data.get("timestamp", datetime.datetime.now().isoformat())
    
    logging.debug(f"Report data contains {len(software_list)} software items")
    
    # Format timestamp
    try:
        dt = datetime.datetime.fromisoformat(timestamp)
        formatted_timestamp = dt.strftime("%Y-%m-%d %H:%M:%S")
    except Exception as e:
        logging.warning(f"Could not format timestamp: {e}")
        formatted_timestamp = timestamp
    
    logging.debug("Preparing HTML report helper functions")
    
    # Safe access to nested dictionaries
    def safe_get(d, keys, default=''):
        """Safely get a value from nested dictionaries"""
        result = d
        for key in keys:
            if isinstance(result, dict) and key in result:
                result = result[key]
            else:
                return default
        return result
    
    # Safe generation of HTML table rows from a list
    def generate_table_rows(items, fields, default=''):
        """Generate HTML table rows from a list of items"""
        if not items or not isinstance(items, list):
            return ""
        
        rows = []
        for item in items:
            row = "<tr>"
            for field_key in fields:
                if isinstance(item, dict) and field_key in item:
                    # Handle nested fields with dot notation
                    if "." in field_key:
                        parts = field_key.split(".")
                        value = item
                        for part in parts:
                            if isinstance(value, dict) and part in value:
                                value = value[part]
                            else:
                                value = default
                                break
                        row += f"<td>{value}</td>"
                    else:
                        row += f"<td>{item.get(field_key, default)}</td>"
                else:
                    row += f"<td>{default}</td>"
            row += "</tr>"
            rows.append(row)
        
        return "".join(rows)
    
    # Create detail card for single item dictionary
    def create_detail_card(title, data_dict, fields=None):
        if not data_dict or not isinstance(data_dict, dict):
            return f"<div class='card'><div class='card-title'>{title}</div><p>No data available</p></div>"
        
        content = f"<div class='card'><div class='card-title'>{title}</div>"
        
        if fields:
            # Use specified fields
            for field in fields:
                label = field.replace("_", " ").title()
                value = safe_get(data_dict, [field], "N/A")
                content += f"<p><strong>{label}:</strong> {value}</p>"
        else:
            # Use all fields
            for key, value in data_dict.items():
                if isinstance(value, dict):
                    continue  # Skip nested dictionaries
                label = key.replace("_", " ").title()
                content += f"<p><strong>{label}:</strong> {value}</p>"
        
        content += "</div>"
        return content
    
    logging.debug("Processing report sections")
    
    # Get hardware information
    cpu_info = safe_get(system_info, ["cpu"], [])
    if isinstance(cpu_info, list) and cpu_info:
        cpu_rows = generate_table_rows(cpu_info, ["Name", "NumberOfCores", "NumberOfLogicalProcessors", "MaxClockSpeedGHz"], "N/A")
    else:
        cpu_rows = "<tr><td colspan='4'>No CPU information available</td></tr>"
    
    # Get memory information
    memory_info = safe_get(system_info, ["memory"], {})
    memory_modules = safe_get(memory_info, ["modules"], [])
    if memory_modules:
        memory_module_rows = generate_table_rows(memory_modules, ["Capacity", "Speed", "Manufacturer", "PartNumber", "DeviceLocator"], "N/A")
    else:
        memory_module_rows = "<tr><td colspan='5'>No memory module information available</td></tr>"
    
    # Get disk information from different sources
    disk_info = {}
    
    # Try to get volumes first
    volumes = safe_get(system_info, ["disk", "volumes"], [])
    if volumes:
        disk_info["volumes"] = volumes
        volume_rows = generate_table_rows(volumes, ["mountpoint", "SizeGB", "FreeGB", "UsedPercent", "fstype"], "N/A")
    else:
        volume_rows = "<tr><td colspan='5'>No volume information available</td></tr>"
    
    # Try to get physical disks
    physical_disks = safe_get(system_info, ["disk", "physical_disks"], [])
    if physical_disks:
        disk_info["physical_disks"] = physical_disks
        physical_disk_rows = generate_table_rows(physical_disks, ["FriendlyName", "SizeFormatted", "MediaType", "HealthStatus"], "N/A")
    else:
        physical_disk_rows = "<tr><td colspan='4'>No physical disk information available</td></tr>"
    
    # Try to get basic disk info
    basic_disk_info = safe_get(system_info, ["disk", "basic_disk_info"], [])
    if basic_disk_info:
        disk_info["basic_disk_info"] = basic_disk_info
        basic_disk_rows = generate_table_rows(basic_disk_info, ["device", "SizeGB", "FreeGB"], "N/A")
    else:
        basic_disk_rows = ""
    
    # Get graphics information
    graphics_adapters = safe_get(system_info, ["graphics", "adapters"], [])
    if graphics_adapters:
        graphics_rows = generate_table_rows(graphics_adapters, ["Name", "VideoRAM_GB", "DriverVersion"], "N/A")
    else:
        graphics_rows = "<tr><td colspan='3'>No graphics information available</td></tr>"
    
    # Get network adapters
    network_adapters = safe_get(system_info, ["network_adapters"], [])
    if network_adapters:
        network_rows = generate_table_rows(network_adapters, ["Name", "InterfaceDescription", "Status", "MacAddress", "LinkSpeed"], "N/A")
    else:
        network_rows = "<tr><td colspan='5'>No network adapter information available</td></tr>"
    
    # Get motherboard information
    motherboard_info = safe_get(system_info, ["hardware", "motherboard"], {})
    
    # Get BIOS information
    bios_info = safe_get(system_info, ["hardware", "bios"], {})
    
    # Get PATH environment variables
    # From system environment variables
    system_env_vars = safe_get(system_info, ["environment_variables"], {})
    system_path = system_env_vars.get("PATH", "").split(os.pathsep) if system_env_vars else []
    
    # From development environment variables
    dev_env_vars = safe_get(dev_environment, ["environment_variables"], {})
    dev_path_entries = dev_env_vars.get("PATH_ENTRIES", []) if dev_env_vars else []
    
    logging.debug(f"Found {len(system_path)} system PATH entries and {len(dev_path_entries)} development PATH entries")
    
    # Generate system PATH rows
    system_path_rows = ""
    for i, path in enumerate(system_path):
        system_path_rows += f"<tr><td>{i+1}</td><td>{path}</td></tr>"
    
    if not system_path_rows:
        system_path_rows = "<tr><td colspan='2'>No system PATH variables found</td></tr>"
    
    # Generate dev PATH rows
    dev_path_rows = ""
    for i, path in enumerate(dev_path_entries):
        dev_path_rows += f"<tr><td>{i+1}</td><td>{path}</td></tr>"
    
    if not dev_path_rows:
        dev_path_rows = "<tr><td colspan='2'>No development PATH entries found</td></tr>"
    
    logging.debug("Generating HTML content")
    
    # Generate HTML
    html_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>System Information Report</title>
    <style>
        body {{ 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f5f5f5;
            color: #333;
        }}
        .header {{
            background-color: #2c3e50;
            color: white;
            padding: 30px 20px;
            text-align: center;
        }}
        h1 {{ 
            margin: 0;
            font-weight: 400;
        }}
        h2 {{
            color: #2c3e50;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
            margin-top: 30px;
        }}
        h3 {{
            color: #3498db;
            margin-top: 25px;
        }}
        .container {{ 
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: white;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            border-radius: 5px;
            margin-top: 20px;
            margin-bottom: 20px;
        }}
        .timestamp {{ 
            color: #777;
            font-style: italic;
            margin-bottom: 20px;
            text-align: center;
        }}
        table {{ 
            border-collapse: collapse;
            width: 100%;
            margin-bottom: 20px;
            font-size: 14px;
        }}
        th, td {{ 
            text-align: left;
            padding: 12px;
            border-bottom: 1px solid #ddd;
        }}
        tr:hover {{
            background-color: #f8f8f8;
        }}
        th {{ 
            background-color: #3498db;
            color: white;
        }}
        .section {{
            background-color: white;
            padding: 20px;
            border-radius: 5px;
            margin-bottom: 30px;
            box-shadow: 0 2px 5px rgba(0,0,0,0.1);
        }}
        .datatable-container {{
            overflow-x: auto;
            max-height: 500px;
            overflow-y: auto;
            margin-top: 20px;
        }}
        .card {{
            background-color: #f9f9f9;
            border-left: 4px solid #3498db;
            padding: 15px;
            margin-bottom: 15px;
            border-radius: 3px;
        }}
        .card-title {{
            font-weight: bold;
            margin-bottom: 5px;
        }}
        .tabs {{
            display: flex;
            flex-wrap: wrap;
            margin-bottom: 20px;
        }}
        .tab-button {{
            background-color: #f1f1f1;
            border: none;
            outline: none;
            cursor: pointer;
            padding: 12px 16px;
            transition: 0.3s;
            font-size: 16px;
            border-radius: 4px 4px 0 0;
            margin-right: 2px;
        }}
        .tab-button:hover {{
            background-color: #ddd;
        }}
        .tab-button.active {{
            background-color: #3498db;
            color: white;
        }}
        .tab-content {{
            display: none;
            padding: 20px;
            border: 1px solid #ddd;
            border-top: none;
            border-radius: 0 0 4px 4px;
        }}
        footer {{
            text-align: center;
            padding: 20px;
            color: #777;
            font-size: 14px;
            border-top: 1px solid #eee;
            margin-top: 20px;
        }}
        @media (max-width: 768px) {{
            th, td {{ 
                padding: 8px;
                font-size: 12px;
            }}
            .tabs {{
                flex-direction: column;
            }}
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>System Information Report</h1>
        <p class="timestamp">Generated on: {formatted_timestamp}</p>
    </div>
    
    <div class="container">
        <div class="tabs">
            <button class="tab-button active" onclick="openTab(event, 'system')">System Information</button>
            <button class="tab-button" onclick="openTab(event, 'hardware')">Hardware</button>
            <button class="tab-button" onclick="openTab(event, 'software')">Installed Software</button>
            <button class="tab-button" onclick="openTab(event, 'dev')">Development Environment</button>
            <button class="tab-button" onclick="openTab(event, 'network')">Network</button>
        </div>
        
        <!-- System Information Tab -->
        <div id="system" class="tab-content" style="display: block;">
            <div class="section">
                <h2>System Specifications</h2>
                <table>
                    <tr><th>Property</th><th>Value</th></tr>
                    <tr><td>Operating System</td><td>{safe_get(system_info, ['basic_info', 'system'], '')} {safe_get(system_info, ['basic_info', 'version'], '')}</td></tr>
                    <tr><td>OS Version</td><td>{safe_get(system_info, ['basic_info', 'release'], '')}</td></tr>
                    <tr><td>OS Build</td><td>{safe_get(system_info, ['hardware', 'os', 'build'], '')}</td></tr>
                    <tr><td>Computer Name</td><td>{safe_get(system_info, ['basic_info', 'hostname'], 'Unknown')}</td></tr>
                    <tr><td>Installation Date</td><td>{safe_get(system_info, ['hardware', 'os', 'installation_date'], 'Unknown')}</td></tr>
                    <tr><td>System Architecture</td><td>{safe_get(system_info, ['basic_info', 'architecture'], 'Unknown')}</td></tr>
                    <tr><td>System Drive</td><td>{safe_get(system_info, ['hardware', 'os', 'system_drive'], 'Unknown')}</td></tr>
                </table>
                
                <h3>系统环境变量 PATH</h3>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>#</th>
                            <th>路径</th>
                        </tr>
                        {system_path_rows}
                    </table>
                </div>
            </div>
        </div>
        
        <!-- Hardware Tab -->
        <div id="hardware" class="tab-content">
            <div class="section">
                <h2>处理器 (CPU)</h2>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>名称</th>
                            <th>物理核心</th>
                            <th>逻辑处理器</th>
                            <th>频率 (GHz)</th>
                        </tr>
                        {cpu_rows}
                    </table>
                </div>
                
                <h2>内存 (RAM)</h2>
                <div class="card">
                    <div class="card-title">总内存</div>
                    <p><strong>容量:</strong> {safe_get(system_info, ['memory', 'total_gb'], 'Unknown')} GB</p>
                    <p><strong>已使用:</strong> {safe_get(system_info, ['memory', 'used_gb'], 'Unknown')} GB ({safe_get(system_info, ['memory', 'percent'], 'Unknown')}%)</p>
                    <p><strong>可用:</strong> {safe_get(system_info, ['memory', 'available_gb'], 'Unknown')} GB</p>
                </div>
                
                <h3>内存条详情</h3>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>容量</th>
                            <th>速度</th>
                            <th>制造商</th>
                            <th>型号</th>
                            <th>插槽</th>
                        </tr>
                        {memory_module_rows}
                    </table>
                </div>
                
                <h2>存储设备</h2>
                <h3>物理磁盘</h3>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>名称</th>
                            <th>容量</th>
                            <th>类型</th>
                            <th>状态</th>
                        </tr>
                        {physical_disk_rows}
                    </table>
                </div>
                
                <h3>分区信息</h3>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>挂载点</th>
                            <th>容量 (GB)</th>
                            <th>剩余空间 (GB)</th>
                            <th>使用率 (%)</th>
                            <th>文件系统</th>
                        </tr>
                        {volume_rows}
                    </table>
                </div>
                
                {f'''
                <h3>基本磁盘信息</h3>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>设备</th>
                            <th>容量 (GB)</th>
                            <th>剩余空间 (GB)</th>
                        </tr>
                        {basic_disk_rows}
                    </table>
                </div>
                ''' if basic_disk_rows else ''}
                
                <h2>显卡 (GPU)</h2>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>名称</th>
                            <th>显存 (GB)</th>
                            <th>驱动版本</th>
                        </tr>
                        {graphics_rows}
                    </table>
                </div>
                
                <h2>主板和BIOS</h2>
                <div class="row">
                    <div class="col">
                        {create_detail_card("主板信息", motherboard_info, ["Manufacturer", "Product", "SerialNumber", "Version"])}
                    </div>
                    <div class="col">
                        {create_detail_card("BIOS信息", bios_info, ["Manufacturer", "Name", "SMBIOSBIOSVersion", "ReleaseDate"])}
                    </div>
                </div>
            </div>
        </div>
        
        <!-- Software Tab -->
        <div id="software" class="tab-content">
            <div class="section">
                <h2>Installed Software <small>({len(software_list)} items)</small></h2>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>Name</th>
                            <th>Version</th>
                            <th>Publisher</th>
                            <th>Install Date</th>
                            <th>Install Location</th>
                            <th>Architecture</th>
                            <th>Type</th>
                        </tr>
                        {"".join([f'''
                        <tr>
                            <td>{app.get('DisplayName', app.get('name', ''))}</td>
                            <td>{app.get('DisplayVersion', app.get('version', ''))}</td>
                            <td>{app.get('Publisher', app.get('publisher', ''))}</td>
                            <td>{app.get('InstallDate', app.get('install_date', ''))}</td>
                            <td>{app.get('InstallLocation', app.get('install_location', ''))}</td>
                            <td>{app.get('architecture', '')}</td>
                            <td>{app.get('type', 'Standard')}</td>
                        </tr>''' for app in software_list[:100]])}
                    </table>
                </div>
                <p><em>Showing first 100 applications. See full list in the HTML directory.</em></p>
            </div>
        </div>
        
        <!-- Development Environment Tab -->
        <div id="dev" class="tab-content">
            <div class="section">
                <h2>Development Environment</h2>
                
                <h3>Programming Languages</h3>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>Language</th>
                            <th>Installed</th>
                            <th>Version</th>
                            <th>Path</th>
                        </tr>
                        {"".join([f'''
                        <tr>
                            <td>{lang.capitalize()}</td>
                            <td>{"Yes" if info.get('installed', False) else "No"}</td>
                            <td>{info.get('version', 'N/A')}</td>
                            <td>{info.get('path', 'N/A')}</td>
                        </tr>''' for lang, info in dev_environment.get('languages', {}).items()])}
                    </table>
                </div>
                
                <h3>Development Tools</h3>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>Tool</th>
                            <th>Installed</th>
                            <th>Version</th>
                            <th>Path</th>
                        </tr>
                        {"".join([f'''
                        <tr>
                            <td>{info.get('name', tool)}</td>
                            <td>{"Yes" if info.get('installed', False) else "No"}</td>
                            <td>{info.get('version', 'N/A')}</td>
                            <td>{info.get('path', 'N/A')}</td>
                        </tr>''' for tool, info in dev_environment.get('development_tools', {}).items()])}
                    </table>
                </div>
                
                <h3>Source Control</h3>
                <div class="card">
                    <div class="card-title">Git</div>
                    <p>Installed: {"Yes" if dev_environment.get('git', {}).get('installed', False) else "No"}</p>
                    <p>Version: {dev_environment.get('git', {}).get('version', 'N/A')}</p>
                </div>
                
                <h3>Container Technology</h3>
                <div class="card">
                    <div class="card-title">Docker</div>
                    <p>Installed: {"Yes" if dev_environment.get('docker', {}).get('installed', False) else "No"}</p>
                    <p>Version: {dev_environment.get('docker', {}).get('version', 'N/A')}</p>
                </div>
                
                <h3>开发环境变量 PATH</h3>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>#</th>
                            <th>路径</th>
                        </tr>
                        {dev_path_rows}
                    </table>
                </div>
            </div>
        </div>
        
        <!-- Network Tab -->
        <div id="network" class="tab-content">
            <div class="section">
                <h2>Network Configuration</h2>
                
                <h3>Network Adapters</h3>
                <div class="datatable-container">
                    <table>
                        <tr>
                            <th>Name</th>
                            <th>Description</th>
                            <th>Status</th>
                            <th>MAC Address</th>
                            <th>Link Speed</th>
                        </tr>
                        {network_rows}
                    </table>
                </div>
                
                <h3>Network Information</h3>
                <table>
                    <tr><th>Property</th><th>Value</th></tr>
                    <tr><td>Hostname</td><td>{safe_get(system_info, ['basic_info', 'hostname'], 'Unknown')}</td></tr>
                    <tr><td>IP Address</td><td>{safe_get(system_info, ['basic_info', 'ip_address'], 'Unknown')}</td></tr>
                    <tr><td>MAC Address</td><td>{safe_get(system_info, ['basic_info', 'mac_address'], 'Unknown')}</td></tr>
                </table>
            </div>
        </div>
        
        <footer>
            Generated by Windows System Information Collector
        </footer>
    </div>
    
    <script>
    function openTab(evt, tabName) {{
        var i, tabcontent, tabbuttons;
        
        // Hide all tab content
        tabcontent = document.getElementsByClassName("tab-content");
        for (i = 0; i < tabcontent.length; i++) {{
            tabcontent[i].style.display = "none";
        }}
        
        // Remove "active" class from all tab buttons
        tabbuttons = document.getElementsByClassName("tab-button");
        for (i = 0; i < tabbuttons.length; i++) {{
            tabbuttons[i].className = tabbuttons[i].className.replace(" active", "");
        }}
        
        // Show the specific tab content and add "active" class to the button
        document.getElementById(tabName).style.display = "block";
        evt.currentTarget.className += " active";
    }}
    </script>
</body>
</html>
"""
    
    # Write to file
    logging.debug(f"Writing HTML report to {output_path}")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    logging.info(f"Comprehensive HTML report created at {output_path}")


def create_index_file(base_dir):
    """Create an index.html file in the base directory that links to all reports"""
    logging.debug(f"Creating index file in {base_dir}")
    
    base_dir = Path(base_dir)
    timestamp = datetime.datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    
    # List directories
    json_dir = base_dir / "json"
    html_dir = base_dir / "html"
    excel_dir = base_dir / "excel"
    markdown_dir = base_dir / "markdown"
    reports_dir = base_dir / "reports"
    
    # Create a list of files in each directory
    def get_files(directory):
        if not directory.exists():
            logging.debug(f"Directory does not exist: {directory}")
            return []
        files = [f for f in directory.iterdir() if f.is_file()]
        logging.debug(f"Found {len(files)} files in {directory}")
        return files
    
    json_files = get_files(json_dir)
    html_files = get_files(html_dir)
    excel_files = get_files(excel_dir)
    markdown_files = get_files(markdown_dir)
    report_files = get_files(reports_dir)
    
    logging.debug("Generating index HTML content")
    
    # Generate HTML
    index_content = f"""<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>System Information Reports</title>
    <style>
        body {{ 
            font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
            margin: 0;
            padding: 0;
            background-color: #f5f5f5;
            color: #333;
        }}
        .header {{
            background-color: #2c3e50;
            color: white;
            padding: 30px 20px;
            text-align: center;
        }}
        h1 {{ 
            margin: 0;
            font-weight: 400;
        }}
        h2 {{
            color: #2c3e50;
            border-bottom: 2px solid #3498db;
            padding-bottom: 10px;
            margin-top: 30px;
        }}
        .container {{ 
            max-width: 1200px;
            margin: 0 auto;
            padding: 20px;
            background-color: white;
            box-shadow: 0 0 10px rgba(0,0,0,0.1);
            border-radius: 5px;
            margin-top: 20px;
            margin-bottom: 20px;
        }}
        .timestamp {{ 
            color: #777;
            font-style: italic;
            margin-bottom: 20px;
            text-align: center;
        }}
        .file-section {{
            margin-bottom: 30px;
            padding: 15px;
            background-color: #f9f9f9;
            border-radius: 5px;
        }}
        .file-list {{
            list-style-type: none;
            padding: 0;
        }}
        .file-list li {{
            margin-bottom: 10px;
            padding: 10px;
            background-color: white;
            border-left: 4px solid #3498db;
            box-shadow: 0 1px 2px rgba(0,0,0,0.1);
        }}
        .file-list li a {{
            color: #3498db;
            text-decoration: none;
            display: block;
        }}
        .file-list li a:hover {{
            text-decoration: underline;
        }}
        .no-files {{
            color: #777;
            font-style: italic;
        }}
        footer {{
            text-align: center;
            padding: 20px;
            color: #777;
            font-size: 14px;
            border-top: 1px solid #eee;
            margin-top: 20px;
        }}
    </style>
</head>
<body>
    <div class="header">
        <h1>System Information Reports</h1>
        <p class="timestamp">Generated on: {timestamp}</p>
    </div>
    
    <div class="container">
        <div class="file-section">
            <h2>Comprehensive Reports</h2>
            <ul class="file-list">
                {"".join([f'<li><a href="reports/{f.name}">{f.name}</a></li>' for f in report_files]) if report_files else '<p class="no-files">No reports found</p>'}
            </ul>
        </div>
        
        <div class="file-section">
            <h2>HTML Files</h2>
            <ul class="file-list">
                {"".join([f'<li><a href="html/{f.name}">{f.name}</a></li>' for f in html_files]) if html_files else '<p class="no-files">No HTML files found</p>'}
            </ul>
        </div>
        
        <div class="file-section">
            <h2>Excel Files</h2>
            <ul class="file-list">
                {"".join([f'<li><a href="excel/{f.name}">{f.name}</a></li>' for f in excel_files]) if excel_files else '<p class="no-files">No Excel files found</p>'}
            </ul>
        </div>
        
        <div class="file-section">
            <h2>JSON Files</h2>
            <ul class="file-list">
                {"".join([f'<li><a href="json/{f.name}">{f.name}</a></li>' for f in json_files]) if json_files else '<p class="no-files">No JSON files found</p>'}
            </ul>
        </div>
        
        <div class="file-section">
            <h2>Markdown Files</h2>
            <ul class="file-list">
                {"".join([f'<li><a href="markdown/{f.name}">{f.name}</a></li>' for f in markdown_files]) if markdown_files else '<p class="no-files">No Markdown files found</p>'}
            </ul>
        </div>
        
        <footer>
            Generated by Windows System Information Collector
        </footer>
    </div>
</body>
</html>
"""
    
    # Write to file
    index_file_path = base_dir / "index.html"
    logging.debug(f"Writing index file to {index_file_path}")
    with open(index_file_path, 'w', encoding='utf-8') as f:
        f.write(index_content)
    logging.info(f"Index file created at {index_file_path}")


if __name__ == "__main__":
    sys.exit(main()) 