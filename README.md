# PyWindowsSoftwareList

## English

A comprehensive Windows system information collection tool that gathers details about installed software, hardware specifications, development environments, and more.

### Features

- **Software Inventory**: Collects detailed information about all installed software
- **Hardware Details**: Gathers information about CPU, memory, disks, graphics cards, motherboard, BIOS, etc.
- **Development Environment**: Detects programming languages, development tools, and SDKs
- **System Information**: Collects OS details, system specifications, environment variables
- **Network Configuration**: Captures network adapters, connections, and settings
- **Comprehensive Reporting**: Exports collected information to various formats:
  - HTML (interactive reports with search capabilities)
  - JSON (complete raw data)
  - Excel (tabular format)
  - Markdown (documentation-friendly)

### Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/PyWindowsSoftwareList.git
   cd PyWindowsSoftwareList
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

### Usage

Run the main script with administrator privileges for complete system access:

```
python main.py
```

The tool will:
1. Collect system information
2. Generate reports in various formats
3. Save all data to a timestamped directory in the `Report` folder

### Output Structure

```
Report/YYYYMMDD_HHMMSS/
  ├── excel/              # Excel spreadsheets
  ├── html/               # HTML reports
  ├── json/               # JSON data files
  ├── logs/               # Log files
  ├── markdown/           # Markdown documents
  ├── reports/            # Comprehensive reports
  └── index.html          # Main index page linking to all reports
```

### Requirements

- Windows operating system
- Python 3.6+
- Administrator privileges (for full system information access)
- Dependencies listed in `requirements.txt`

### License

This project is licensed under the GNU General Public License v3.0 (GPL-3.0) - see the [LICENSE](LICENSE) file for details.

---

## 中文

一个全面的Windows系统信息收集工具，可收集已安装软件、硬件规格、开发环境等详细信息。

### 功能特点

- **软件清单**：收集所有已安装软件的详细信息
- **硬件详情**：收集关于CPU、内存、磁盘、显卡、主板、BIOS等信息
- **开发环境**：检测编程语言、开发工具和SDK
- **系统信息**：收集操作系统细节、系统规格、环境变量
- **网络配置**：捕获网络适配器、连接和设置
- **综合报告**：将收集的信息导出为各种格式：
  - HTML（具有搜索功能的交互式报告）
  - JSON（完整原始数据）
  - Excel（表格格式）
  - Markdown（适合文档的格式）

### 安装

1. 克隆仓库：
   ```
   git clone https://github.com/JadeiteRedolence/PyWindowsSoftwareList.git
   cd PyWindowsSoftwareList
   ```

2. 安装所需依赖：
   ```
   pip install -r requirements.txt
   ```

### 使用方法

以管理员权限运行主脚本，以获得完整的系统访问权限：

```
python main.py
```

该工具将：
1. 收集系统信息
2. 生成各种格式的报告
3. 将所有数据保存到`Report`文件夹中的时间戳目录

### 输出结构

```
Report/YYYYMMDD_HHMMSS/
  ├── excel/              # Excel电子表格
  ├── html/               # HTML报告
  ├── json/               # JSON数据文件
  ├── logs/               # 日志文件
  ├── markdown/           # Markdown文档
  ├── reports/            # 综合报告
  └── index.html          # 链接到所有报告的主索引页
```

### 系统要求

- Windows操作系统
- Python 3.6+
- 管理员权限（用于完整的系统信息访问）
- `requirements.txt`中列出的依赖项

### 许可证

该项目采用GNU通用公共许可证v3.0（GPL-3.0）许可 - 详情请参阅[LICENSE](LICENSE)文件。 