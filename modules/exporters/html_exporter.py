import datetime
import pandas as pd
import logging

def export_software_list_to_html(software_list, output_path, current_time=None):
    """
    将软件列表导出为HTML格式
    
    参数:
    - software_list: 软件列表
    - output_path: 输出路径
    - current_time: 当前时间字符串
    """
    if not software_list:
        logging.warning("Software list is empty")
        # 创建一个简单的HTML文件，显示没有找到软件
        html_content = f"""
        <!DOCTYPE html>
        <html lang="zh-CN">
        <head>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1.0">
            <title>软件列表</title>
            <style>
                body {{
                    font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                    margin: 0;
                    padding: 20px;
                    background-color: #f8f9fa;
                    color: #333;
                }}
                .container {{
                    width: 100%;
                    max-width: 1200px;
                    margin: 0 auto;
                }}
                h1 {{
                    color: #2c3e50;
                    border-bottom: 2px solid #3498db;
                    padding-bottom: 10px;
                }}
                .alert {{
                    background-color: #f8d7da;
                    color: #721c24;
                    padding: 15px;
                    border-radius: 4px;
                    margin-top: 20px;
                }}
                .timestamp {{
                    font-size: 0.8em;
                    color: #6c757d;
                    margin-top: 5px;
                }}
            </style>
        </head>
        <body>
            <div class="container">
                <h1>已安装软件列表</h1>
                <div class="timestamp">生成时间: {current_time if current_time else datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
                <div class="alert">
                    <p>未发现已安装的软件。这可能是因为:</p>
                    <ul>
                        <li>访问Windows注册表时出现权限问题</li>
                        <li>软件信息收集方法失败</li>
                        <li>系统上未安装软件(极少见)</li>
                    </ul>
                    <p>建议尝试使用管理员权限运行此工具。</p>
                </div>
            </div>
        </body>
        </html>
        """
        with open(output_path, 'w', encoding='utf-8') as f:
            f.write(html_content)
        return

    # 提取所需的字段
    processed_data = []
    logging.debug(f"Processing {len(software_list)} software items for HTML export")
    
    # 定义常见的字段映射，支持不同收集方法的字段命名
    field_mappings = {
        "name": ["DisplayName", "name", "Name"],
        "version": ["DisplayVersion", "version", "Version"],
        "publisher": ["Publisher", "publisher", "Vendor", "Company"],
        "install_date": ["InstallDate", "install_date"],
        "install_location": ["InstallLocation", "install_location", "Location", "Path"],
        "architecture": ["architecture", "Architecture", "arch"],
        "type": ["type", "Type", "ApplicationType", "ProgramType"],
        "uninstall_string": ["UninstallString", "uninstall_string", "UninstallCommand"],
        "estimated_size": ["EstimatedSize", "size", "Size"],
    }
    
    # 从不同源字段中提取统一的值
    def extract_field(item, field_names):
        for field in field_names:
            value = item.get(field)
            if value and value != "":
                return value
        return ""
    
    for sw in software_list:
        # 尝试从不同的字段名称提取一致的信息
        software_item = {
            "名称": extract_field(sw, field_mappings["name"]),
            "版本": extract_field(sw, field_mappings["version"]),
            "发布商": extract_field(sw, field_mappings["publisher"]),
            "安装日期": extract_field(sw, field_mappings["install_date"]),
            "安装位置": extract_field(sw, field_mappings["install_location"]),
            "架构": extract_field(sw, field_mappings["architecture"]),
            "类型": extract_field(sw, field_mappings["type"]) or "标准应用程序",
            "卸载字符串": extract_field(sw, field_mappings["uninstall_string"]),
            "估计大小(KB)": extract_field(sw, field_mappings["estimated_size"]),
        }
        
        # 清理空值
        if not software_item["名称"]:
            continue  # 跳过没有名称的项
            
        processed_data.append(software_item)
    
    logging.debug(f"Processed {len(processed_data)} valid software items")

    # 创建DataFrame
    df = pd.DataFrame(processed_data)
    
    # 生成HTML表格
    html_table = df.to_html(index=False, classes='table table-striped table-hover', border=0)
    
    # 创建完整HTML文档
    html_content = f"""
    <!DOCTYPE html>
    <html lang="zh-CN">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>软件列表</title>
        <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/css/bootstrap.min.css">
        <style>
            body {{
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                margin: 0;
                padding: 20px;
                background-color: #f8f9fa;
                color: #333;
                height: 100vh;
                overflow: auto;
            }}
            .container-fluid {{
                width: 100%;
                padding-right: 15px;
                padding-left: 15px;
                margin-right: auto;
                margin-left: auto;
            }}
            h1 {{
                color: #2c3e50;
                border-bottom: 2px solid #3498db;
                padding-bottom: 10px;
            }}
            .table-responsive {{
                width: 100%;
                overflow-x: auto;
            }}
            .table {{
                width: 100% !important;
                margin-bottom: 1rem;
                color: #212529;
                border-collapse: collapse;
            }}
            .table thead th {{
                position: sticky;
                top: 0;
                background-color: #f8f9fa;
                z-index: 1;
                border-bottom: 2px solid #dee2e6;
            }}
            .table-hover tbody tr:hover {{
                background-color: rgba(0, 123, 255, 0.1);
            }}
            .timestamp {{
                font-size: 0.8em;
                color: #6c757d;
                margin-top: 5px;
                margin-bottom: 15px;
            }}
            .info-box {{
                background-color: #e2f0fd;
                border-left: 5px solid #3498db;
                padding: 15px;
                margin-bottom: 20px;
                border-radius: 3px;
            }}
            @media (max-width: 768px) {{
                .container-fluid {{
                    padding: 10px;
                }}
            }}
            .search-container {{
                margin-bottom: 20px;
            }}
            #searchInput {{
                width: 100%;
                padding: 10px;
                border: 1px solid #ddd;
                border-radius: 4px;
                font-size: 16px;
            }}
            .stats {{
                margin-bottom: 15px;
                font-size: 0.9em;
                color: #6c757d;
            }}
        </style>
    </head>
    <body>
        <div class="container-fluid">
            <h1>已安装软件列表</h1>
            <div class="timestamp">生成时间: {current_time if current_time else datetime.datetime.now().strftime('%Y-%m-%d %H:%M:%S')}</div>
            
            <div class="info-box">
                <p><strong>总计:</strong> {len(processed_data)} 个软件</p>
            </div>
            
            <div class="search-container">
                <input type="text" id="searchInput" placeholder="搜索软件...">
            </div>
            
            <div class="table-responsive">
                {html_table}
            </div>
        </div>
        <script src="https://cdn.jsdelivr.net/npm/bootstrap@5.2.3/dist/js/bootstrap.bundle.min.js"></script>
        <script>
            // 自动调整表格高度以适应屏幕
            function adjustTableHeight() {{
                const container = document.querySelector('.table-responsive');
                const header = document.querySelector('h1').offsetHeight;
                const timestamp = document.querySelector('.timestamp').offsetHeight;
                const infoBox = document.querySelector('.info-box').offsetHeight;
                const searchBox = document.querySelector('.search-container').offsetHeight;
                const windowHeight = window.innerHeight;
                const padding = 60; // 考虑页面内边距
                
                container.style.maxHeight = (windowHeight - header - timestamp - infoBox - searchBox - padding) + 'px';
            }}
            
            // 搜索功能
            function setupSearch() {{
                const searchInput = document.getElementById('searchInput');
                const table = document.querySelector('.table');
                const rows = table.querySelectorAll('tbody tr');
                
                searchInput.addEventListener('keyup', function() {{
                    const searchTerm = searchInput.value.toLowerCase();
                    
                    rows.forEach(row => {{
                        const text = row.textContent.toLowerCase();
                        if(text.indexOf(searchTerm) > -1) {{
                            row.style.display = "";
                        }} else {{
                            row.style.display = "none";
                        }}
                    }});
                }});
            }}
            
            // 页面加载时调整
            window.addEventListener('load', function() {{
                adjustTableHeight();
                setupSearch();
            }});
            
            // 窗口大小改变时调整
            window.addEventListener('resize', adjustTableHeight);
        </script>
    </body>
    </html>
    """
    
    # 写入HTML文件
    logging.debug(f"Writing HTML content to {output_path}")
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(html_content)
    logging.debug(f"HTML export completed successfully")

def export_to_html(software_list, output_path, current_time=None):
    """
    Export software list to HTML format
    
    Args:
        software_list: List of software
        output_path: Path to save HTML file
        current_time: Current time string (optional)
    """
    return export_software_list_to_html(software_list, output_path, current_time) 