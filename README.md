# PyWindowsSoftwareList

## English

A comprehensive Windows system information collection tool that gathers details about installed software, hardware specifications, development environments, and more. This tool provides robust and reliable system information collection with advanced error handling and timeout mechanisms.

### ✨ Features

- **🔧 Software Inventory**: Collects detailed information about all installed software
  - Traditional Win32 applications from registry
  - UWP/Modern apps via PowerShell
  - Startup programs and services
- **💻 Hardware Details**: Comprehensive hardware information collection
  - CPU specifications (cores, threads, cache, frequency)
  - Memory details (total, usage, individual modules)
  - Storage devices (disks, partitions, usage statistics)
  - Graphics cards (name, VRAM, driver versions)
  - Motherboard and BIOS information
- **⚙️ Development Environment**: Detects programming languages, tools, and SDKs
  - Programming languages (Python, Java, Node.js, Go, Ruby, PHP, .NET)
  - IDEs and editors (VS Code, Visual Studio, IntelliJ, PyCharm, etc.)
  - SDKs (Windows, Android, Java)
  - Version control systems (Git)
  - Containerization tools (Docker)
- **🖥️ System Information**: Detailed OS and system specifications
  - Operating system details and installation date
  - Environment variables and PATH entries
  - Network adapters and configurations
  - User accounts and security settings
- **📊 Comprehensive Reporting**: Exports collected information to various formats
  - **HTML** (interactive reports with tabbed interface and search capabilities)
  - **JSON** (complete structured data)
  - **Excel** (spreadsheet format with multiple sheets)
  - **Markdown** (documentation-friendly format)

### 🚀 Recent Improvements

- **Enhanced Reliability**: Added timeout mechanisms to prevent system hangs
- **Better Error Handling**: Improved PowerShell command execution with encoding fixes
- **Performance Optimization**: Reduced collection time from potentially infinite to ~30 seconds
- **Absolute Path Resolution**: Uses absolute paths for system commands to avoid PATH issues
- **Unicode Support**: Proper handling of international characters and special symbols
- **Comprehensive Data Structure**: Fixed data mapping issues in HTML reports

### 📋 Requirements

- **Operating System**: Windows 10/11 (tested on Windows 11 build 27842)
- **Python Version**: Python 3.8+ (recommended Python 3.10+)
- **Privileges**: Administrator privileges recommended for complete system access
- **PowerShell**: Windows PowerShell 5.1 or PowerShell 7+ for advanced features

### 🔧 Installation

1. **Clone the repository**:
   ```bash
   git clone https://github.com/JadeiteRedolence/PyWindowsSoftwareList.git
   cd PyWindowsSoftwareList
   ```

2. **Install dependencies**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Verify installation**:
   ```bash
   python main.py --help
   ```

### 🎯 Usage

**Basic Usage** (recommended):
```bash
python main.py
```

**Run with administrator privileges** for complete system access:
```bash
# Right-click Command Prompt -> "Run as administrator"
python main.py
```

### 📁 Output Structure

```
Report/YYYYMMDD_HHMMSS/
  ├── excel/                    # Excel spreadsheets
  │   └── software_list.xlsx
  ├── html/                     # Individual HTML reports
  │   └── software_list.html
  ├── json/                     # Structured data files
  │   ├── complete_report_data.json
  │   ├── system_info.json
  │   ├── software_list.json
  │   └── dev_environment.json
  ├── logs/                     # Debug and execution logs
  │   └── debug_YYYYMMDD_HHMMSS.log
  ├── markdown/                 # Markdown documentation
  │   └── software_list.md
  ├── reports/                  # Comprehensive reports
  │   └── system_report.html   # Main comprehensive report
  └── index.html               # Navigation index
```

### 🎨 Report Features

- **Interactive HTML Reports**: Tabbed interface with system, hardware, software, development, and network sections
- **Search Functionality**: Built-in search for finding specific software or information
- **Responsive Design**: Works on desktop and mobile devices
- **Data Export**: Easy copy-paste of data from HTML tables
- **Professional Styling**: Clean, modern interface with proper typography

### 🛠️ Troubleshooting

**Common Issues:**

1. **Program hangs during collection**:
   - Ensure you have administrator privileges
   - Check Windows PowerShell is available and functional
   - Verify antivirus isn't blocking PowerShell execution

2. **Missing or "Unknown" information**:
   - Run as administrator for complete hardware access
   - Ensure WMI service is running
   - Check Windows PowerShell execution policy

3. **Encoding errors**:
   - The tool now handles Unicode automatically
   - Ensure system locale supports UTF-8

4. **PowerShell execution errors**:
   - Run `Set-ExecutionPolicy RemoteSigned` in PowerShell as admin
   - Check if PowerShell 7+ is installed for better Unicode support

**Performance Tips:**
- Close unnecessary applications before running
- Run on SSD for faster file operations
- Ensure adequate free disk space (>100MB)

### 🤝 Contributing

Contributions are welcome! Please feel free to submit a Pull Request. For major changes, please open an issue first to discuss what you would like to change.

### 📄 License

This project is licensed under the GNU General Public License v3.0 (GPL-3.0) - see the [LICENSE](LICENSE) file for details.

---

## 中文

一个全面的Windows系统信息收集工具，可收集已安装软件、硬件规格、开发环境等详细信息。该工具提供强大可靠的系统信息收集功能，具有先进的错误处理和超时机制。

### ✨ 功能特点

- **🔧 软件清单**：收集所有已安装软件的详细信息
  - 来自注册表的传统Win32应用程序
  - 通过PowerShell获取UWP/现代应用程序
  - 启动程序和服务
- **💻 硬件详情**：全面的硬件信息收集
  - CPU规格（核心、线程、缓存、频率）
  - 内存详情（总量、使用情况、单个模块）
  - 存储设备（磁盘、分区、使用统计）
  - 显卡（名称、显存、驱动版本）
  - 主板和BIOS信息
- **⚙️ 开发环境**：检测编程语言、工具和SDK
  - 编程语言（Python、Java、Node.js、Go、Ruby、PHP、.NET）
  - IDE和编辑器（VS Code、Visual Studio、IntelliJ、PyCharm等）
  - SDK（Windows、Android、Java）
  - 版本控制系统（Git）
  - 容器化工具（Docker）
- **🖥️ 系统信息**：详细的操作系统和系统规格
  - 操作系统详情和安装日期
  - 环境变量和PATH条目
  - 网络适配器和配置
  - 用户账户和安全设置
- **📊 综合报告**：将收集的信息导出为各种格式
  - **HTML**（带选项卡界面和搜索功能的交互式报告）
  - **JSON**（完整的结构化数据）
  - **Excel**（带多个工作表的电子表格格式）
  - **Markdown**（适合文档的格式）

### 🚀 最新改进

- **增强可靠性**：添加超时机制防止系统挂起
- **更好的错误处理**：改进PowerShell命令执行，修复编码问题
- **性能优化**：将收集时间从可能的无限期缩短到约30秒
- **绝对路径解析**：对系统命令使用绝对路径以避免PATH问题
- **Unicode支持**：正确处理国际字符和特殊符号
- **完整的数据结构**：修复HTML报告中的数据映射问题

### 📋 系统要求

- **操作系统**：Windows 10/11（在Windows 11 build 27842上测试）
- **Python版本**：Python 3.8+（推荐Python 3.10+）
- **权限**：建议使用管理员权限以获得完整的系统访问
- **PowerShell**：Windows PowerShell 5.1或PowerShell 7+用于高级功能

### 🔧 安装

1. **克隆仓库**：
   ```bash
   git clone https://github.com/JadeiteRedolence/PyWindowsSoftwareList.git
   cd PyWindowsSoftwareList
   ```

2. **安装依赖**：
   ```bash
   pip install -r requirements.txt
   ```

3. **验证安装**：
   ```bash
   python main.py
   ```

### 🎯 使用方法

**基本使用**（推荐）：
```bash
python main.py
```

**以管理员权限运行**以获得完整的系统访问：
```bash
# 右键命令提示符 -> "以管理员身份运行"
python main.py
```

### 📁 输出结构

```
Report/YYYYMMDD_HHMMSS/
  ├── excel/                    # Excel电子表格
  │   └── software_list.xlsx
  ├── html/                     # 单个HTML报告
  │   └── software_list.html
  ├── json/                     # 结构化数据文件
  │   ├── complete_report_data.json
  │   ├── system_info.json
  │   ├── software_list.json
  │   └── dev_environment.json
  ├── logs/                     # 调试和执行日志
  │   └── debug_YYYYMMDD_HHMMSS.log
  ├── markdown/                 # Markdown文档
  │   └── software_list.md
  ├── reports/                  # 综合报告
  │   └── system_report.html   # 主要综合报告
  └── index.html               # 导航索引
```

### 🎨 报告功能

- **交互式HTML报告**：具有系统、硬件、软件、开发和网络部分的选项卡界面
- **搜索功能**：用于查找特定软件或信息的内置搜索
- **响应式设计**：在桌面和移动设备上都能正常工作
- **数据导出**：从HTML表格轻松复制粘贴数据
- **专业样式**：具有适当排版的简洁现代界面

### 🛠️ 故障排除

**常见问题：**

1. **程序在收集过程中挂起**：
   - 确保拥有管理员权限
   - 检查Windows PowerShell是否可用且功能正常
   - 验证防病毒软件没有阻止PowerShell执行

2. **信息缺失或显示"Unknown"**：
   - 以管理员身份运行以获得完整的硬件访问权限
   - 确保WMI服务正在运行
   - 检查Windows PowerShell执行策略

3. **编码错误**：
   - 工具现在自动处理Unicode
   - 确保系统区域设置支持UTF-8

4. **PowerShell执行错误**：
   - 在PowerShell中以管理员身份运行`Set-ExecutionPolicy RemoteSigned`
   - 检查是否安装了PowerShell 7+以获得更好的Unicode支持

**性能提示：**
- 运行前关闭不必要的应用程序
- 在SSD上运行以加快文件操作
- 确保有足够的可用磁盘空间（>100MB）

### 🤝 贡献

欢迎贡献！请随时提交Pull Request。对于重大更改，请先开启issue讨论您想要更改的内容。

### 📄 许可证

该项目采用GNU通用公共许可证v3.0（GPL-3.0）许可 - 详情请参阅[LICENSE](LICENSE)文件。 