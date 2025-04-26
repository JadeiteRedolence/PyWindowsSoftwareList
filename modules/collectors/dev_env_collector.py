import os
import subprocess
import json
import sys
import re
from pathlib import Path
import winreg

def get_installed_programming_languages():
    """
    获取已安装的编程语言和环境信息
    """
    languages = {
        "python": {"installed": False},
        "java": {"installed": False},
        "nodejs": {"installed": False},
        "go": {"installed": False},
        "ruby": {"installed": False},
        "php": {"installed": False},
        "dotnet": {"installed": False}
    }
    
    # 检查Python
    try:
        result = subprocess.run(["python", "--version"], capture_output=True, text=True)
        if result.returncode == 0:
            version = result.stdout.strip()
            languages["python"] = {
                "installed": True,
                "version": version,
                "path": sys.executable,
                "packages": []
            }
            
            # 获取已安装的Python包
            try:
                pip_result = subprocess.run([sys.executable, "-m", "pip", "list", "--format=json"], 
                                           capture_output=True, text=True)
                if pip_result.returncode == 0:
                    packages = json.loads(pip_result.stdout)
                    languages["python"]["packages"] = packages
            except:
                pass
    except:
        pass
    
    # 检查Java
    try:
        result = subprocess.run(["java", "-version"], capture_output=True, text=True, stderr=subprocess.STDOUT)
        if result.returncode == 0:
            version_text = result.stdout
            # 解析Java版本输出
            match = re.search(r'version "([^"]+)"', version_text)
            if match:
                version = match.group(1)
                languages["java"] = {
                    "installed": True,
                    "version": version
                }
                
                # 尝试获取JAVA_HOME
                java_home = os.environ.get("JAVA_HOME", "")
                if java_home:
                    languages["java"]["home"] = java_home
    except:
        pass
    
    # 检查Node.js
    try:
        result = subprocess.run(["node", "--version"], capture_output=True, text=True)
        if result.returncode == 0:
            version = result.stdout.strip()
            npm_version = ""
            
            # 获取npm版本
            try:
                npm_result = subprocess.run(["npm", "--version"], capture_output=True, text=True)
                if npm_result.returncode == 0:
                    npm_version = npm_result.stdout.strip()
            except:
                pass
                
            languages["nodejs"] = {
                "installed": True,
                "version": version,
                "npm_version": npm_version
            }
            
            # 尝试获取全局安装的包
            try:
                npm_list = subprocess.run(["npm", "list", "-g", "--json", "--depth=0"], 
                                         capture_output=True, text=True)
                if npm_list.returncode == 0:
                    packages = json.loads(npm_list.stdout)
                    if "dependencies" in packages:
                        languages["nodejs"]["global_packages"] = packages["dependencies"]
            except:
                pass
    except:
        pass
    
    # 检查Go
    try:
        result = subprocess.run(["go", "version"], capture_output=True, text=True)
        if result.returncode == 0:
            version = result.stdout.strip()
            languages["go"] = {
                "installed": True,
                "version": version
            }
            
            # 尝试获取GOPATH和GOROOT
            go_path = os.environ.get("GOPATH", "")
            go_root = os.environ.get("GOROOT", "")
            if go_path:
                languages["go"]["gopath"] = go_path
            if go_root:
                languages["go"]["goroot"] = go_root
    except:
        pass
    
    # 检查Ruby
    try:
        result = subprocess.run(["ruby", "--version"], capture_output=True, text=True)
        if result.returncode == 0:
            version = result.stdout.strip()
            languages["ruby"] = {
                "installed": True,
                "version": version
            }
            
            # 获取已安装的gem
            try:
                gem_result = subprocess.run(["gem", "list", "--local"], capture_output=True, text=True)
                if gem_result.returncode == 0:
                    gems = []
                    for line in gem_result.stdout.strip().split("\n"):
                        if line and not line.startswith("***"):
                            gems.append(line.strip())
                    languages["ruby"]["gems"] = gems
            except:
                pass
    except:
        pass
    
    # 检查PHP
    try:
        result = subprocess.run(["php", "--version"], capture_output=True, text=True)
        if result.returncode == 0:
            version_text = result.stdout.strip()
            match = re.search(r'PHP (\d+\.\d+\.\d+)', version_text)
            if match:
                version = match.group(1)
                languages["php"] = {
                    "installed": True,
                    "version": version
                }
    except:
        pass
    
    # 检查.NET SDK
    try:
        result = subprocess.run(["dotnet", "--list-sdks"], capture_output=True, text=True)
        if result.returncode == 0:
            versions = []
            for line in result.stdout.strip().split("\n"):
                if line:
                    parts = line.split("[")
                    if len(parts) > 1:
                        version = parts[0].strip()
                        path = parts[1].replace("]", "").strip()
                        versions.append({"version": version, "path": path})
            
            if versions:
                languages["dotnet"] = {
                    "installed": True,
                    "sdk_versions": versions
                }
                
                # 获取已安装的.NET运行时
                try:
                    runtime_result = subprocess.run(["dotnet", "--list-runtimes"], 
                                                   capture_output=True, text=True)
                    if runtime_result.returncode == 0:
                        runtimes = []
                        for line in runtime_result.stdout.strip().split("\n"):
                            if line:
                                parts = line.split("[")
                                if len(parts) > 1:
                                    runtime_info = parts[0].strip().split(" ")
                                    runtime_name = runtime_info[0]
                                    runtime_version = runtime_info[1]
                                    path = parts[1].replace("]", "").strip()
                                    runtimes.append({
                                        "name": runtime_name,
                                        "version": runtime_version,
                                        "path": path
                                    })
                        
                        if runtimes:
                            languages["dotnet"]["runtimes"] = runtimes
                except:
                    pass
    except:
        pass
    
    return languages

def get_development_tools():
    """
    获取已安装的开发工具和IDE信息
    """
    dev_tools = {}
    
    # 常见IDE和开发工具路径
    tool_paths = {
        "vscode": [
            r"C:\Program Files\Microsoft VS Code",
            r"C:\Users\%USERNAME%\AppData\Local\Programs\Microsoft VS Code"
        ],
        "visual_studio": [
            r"C:\Program Files\Microsoft Visual Studio",
            r"C:\Program Files (x86)\Microsoft Visual Studio"
        ],
        "intellij": [
            r"C:\Program Files\JetBrains\IntelliJ IDEA*"
        ],
        "pycharm": [
            r"C:\Program Files\JetBrains\PyCharm*"
        ],
        "eclipse": [
            r"C:\eclipse",
            r"C:\Program Files\Eclipse Foundation"
        ],
        "android_studio": [
            r"C:\Program Files\Android\Android Studio"
        ],
        "docker": [
            r"C:\Program Files\Docker\Docker"
        ]
    }
    
    # 替换用户名
    username = os.environ.get("USERNAME", "")
    for tool, paths in tool_paths.items():
        for i, path in enumerate(paths):
            if "%USERNAME%" in path:
                tool_paths[tool][i] = path.replace("%USERNAME%", username)
    
    # 检查VS Code
    vscode_found = False
    for path in tool_paths["vscode"]:
        if os.path.exists(path):
            vscode_found = True
            # 获取版本信息
            try:
                version = "Unknown"
                code_path = os.path.join(path, "bin", "code.cmd")
                if os.path.exists(code_path):
                    result = subprocess.run([code_path, "--version"], 
                                          capture_output=True, text=True)
                    if result.returncode == 0:
                        version = result.stdout.strip().split("\n")[0]
                
                dev_tools["vscode"] = {
                    "name": "Visual Studio Code",
                    "installed": True,
                    "path": path,
                    "version": version
                }
                
                # 尝试获取已安装的扩展
                try:
                    extensions_result = subprocess.run([code_path, "--list-extensions"], 
                                                     capture_output=True, text=True)
                    if extensions_result.returncode == 0:
                        extensions = extensions_result.stdout.strip().split("\n")
                        if extensions and extensions[0]:  # 确保不是空列表
                            dev_tools["vscode"]["extensions"] = extensions
                except:
                    pass
                    
                break
            except:
                dev_tools["vscode"] = {
                    "name": "Visual Studio Code",
                    "installed": True,
                    "path": path
                }
                break
    
    if not vscode_found:
        dev_tools["vscode"] = {
            "name": "Visual Studio Code",
            "installed": False
        }
    
    # 检查Visual Studio
    vs_found = False
    for base_path in tool_paths["visual_studio"]:
        if os.path.exists(base_path):
            # Visual Studio可能有多个版本
            vs_versions = []
            try:
                for vs_year in os.listdir(base_path):
                    year_path = os.path.join(base_path, vs_year)
                    if os.path.isdir(year_path):
                        for edition in os.listdir(year_path):
                            edition_path = os.path.join(year_path, edition)
                            if os.path.isdir(edition_path):
                                vs_versions.append({
                                    "year": vs_year,
                                    "edition": edition,
                                    "path": edition_path
                                })
            except:
                pass
                
            if vs_versions:
                vs_found = True
                dev_tools["visual_studio"] = {
                    "name": "Visual Studio",
                    "installed": True,
                    "versions": vs_versions
                }
                break
    
    if not vs_found:
        # 通过注册表检查Visual Studio
        try:
            vs_keys = [
                r"SOFTWARE\Microsoft\VisualStudio\15.0",  # VS 2017
                r"SOFTWARE\Microsoft\VisualStudio\14.0",  # VS 2015
                r"SOFTWARE\Microsoft\VisualStudio\12.0",  # VS 2013
                r"SOFTWARE\Microsoft\VisualStudio\11.0",  # VS 2012
                r"SOFTWARE\Microsoft\VisualStudio\10.0",  # VS 2010
            ]
            
            vs_versions = []
            for key_path in vs_keys:
                try:
                    key = winreg.OpenKey(winreg.HKEY_LOCAL_MACHINE, key_path)
                    install_dir, _ = winreg.QueryValueEx(key, "InstallDir")
                    version = key_path.split("\\")[-1]
                    vs_versions.append({
                        "version": version,
                        "path": install_dir
                    })
                    winreg.CloseKey(key)
                except:
                    pass
            
            if vs_versions:
                vs_found = True
                dev_tools["visual_studio"] = {
                    "name": "Visual Studio",
                    "installed": True,
                    "versions": vs_versions
                }
        except:
            pass
    
    if not vs_found:
        dev_tools["visual_studio"] = {
            "name": "Visual Studio",
            "installed": False
        }
    
    # 检查JetBrains IDEs (IntelliJ IDEA, PyCharm)
    for tool, paths in [("intellij", tool_paths["intellij"]), 
                        ("pycharm", tool_paths["pycharm"])]:
        tool_found = False
        for path_pattern in paths:
            if "*" in path_pattern:
                # 使用通配符匹配
                base_dir = os.path.dirname(path_pattern)
                pattern = os.path.basename(path_pattern)
                
                if os.path.exists(base_dir):
                    try:
                        matching_dirs = [d for d in os.listdir(base_dir) 
                                        if os.path.isdir(os.path.join(base_dir, d)) 
                                        and pattern.replace("*", "") in d]
                        
                        if matching_dirs:
                            tool_found = True
                            if tool == "intellij":
                                name = "IntelliJ IDEA"
                            else:
                                name = "PyCharm"
                                
                            versions = []
                            for d in matching_dirs:
                                full_path = os.path.join(base_dir, d)
                                # 尝试找到版本
                                version = "Unknown"
                                if "20" in d:  # 通常包含年份, 如"2022.1"
                                    version_match = re.search(r'(20\d\d\.\d+)', d)
                                    if version_match:
                                        version = version_match.group(1)
                                        
                                versions.append({
                                    "path": full_path,
                                    "version": version
                                })
                                
                            dev_tools[tool] = {
                                "name": name,
                                "installed": True,
                                "versions": versions
                            }
                            break
                    except:
                        pass
            elif os.path.exists(path_pattern):
                tool_found = True
                if tool == "intellij":
                    name = "IntelliJ IDEA"
                else:
                    name = "PyCharm"
                    
                dev_tools[tool] = {
                    "name": name,
                    "installed": True,
                    "path": path_pattern
                }
                break
                
        if not tool_found:
            if tool == "intellij":
                name = "IntelliJ IDEA"
            else:
                name = "PyCharm"
                
            dev_tools[tool] = {
                "name": name,
                "installed": False
            }
    
    # 检查Eclipse
    eclipse_found = False
    for path in tool_paths["eclipse"]:
        if os.path.exists(path):
            eclipse_found = True
            dev_tools["eclipse"] = {
                "name": "Eclipse",
                "installed": True,
                "path": path
            }
            break
    
    if not eclipse_found:
        dev_tools["eclipse"] = {
            "name": "Eclipse",
            "installed": False
        }
    
    # 检查Android Studio
    android_studio_found = False
    for path in tool_paths["android_studio"]:
        if os.path.exists(path):
            android_studio_found = True
            dev_tools["android_studio"] = {
                "name": "Android Studio",
                "installed": True,
                "path": path
            }
            break
    
    if not android_studio_found:
        dev_tools["android_studio"] = {
            "name": "Android Studio",
            "installed": False
        }
    
    # 检查Docker Desktop
    docker_found = False
    for path in tool_paths["docker"]:
        if os.path.exists(path):
            docker_found = True
            # 尝试获取Docker版本
            try:
                result = subprocess.run(["docker", "--version"], capture_output=True, text=True)
                if result.returncode == 0:
                    version = result.stdout.strip()
                    dev_tools["docker"] = {
                        "name": "Docker Desktop",
                        "installed": True,
                        "path": path,
                        "version": version
                    }
                else:
                    dev_tools["docker"] = {
                        "name": "Docker Desktop",
                        "installed": True,
                        "path": path
                    }
            except:
                dev_tools["docker"] = {
                    "name": "Docker Desktop",
                    "installed": True,
                    "path": path
                }
            break
    
    if not docker_found:
        # 尝试通过命令检查Docker
        try:
            result = subprocess.run(["docker", "--version"], capture_output=True, text=True)
            if result.returncode == 0:
                version = result.stdout.strip()
                docker_found = True
                dev_tools["docker"] = {
                    "name": "Docker",
                    "installed": True,
                    "version": version
                }
        except:
            pass
            
    if not docker_found:
        dev_tools["docker"] = {
            "name": "Docker Desktop",
            "installed": False
        }
    
    return dev_tools

def get_development_sdks():
    """
    获取已安装的SDK信息
    """
    sdks = {}
    
    # Windows SDK
    windows_sdk_versions = []
    sdk_paths = [
        r"C:\Program Files (x86)\Windows Kits\10",
        r"C:\Program Files (x86)\Windows Kits\8.1",
        r"C:\Program Files (x86)\Windows Kits\8.0"
    ]
    
    for path in sdk_paths:
        if os.path.exists(path):
            # 检查IncludeVersion目录来确定已安装的版本
            include_path = os.path.join(path, "Include")
            if os.path.exists(include_path):
                try:
                    versions = [v for v in os.listdir(include_path) 
                               if os.path.isdir(os.path.join(include_path, v))]
                    if versions:
                        base_version = os.path.basename(path).replace("Windows Kits\\", "")
                        windows_sdk_versions.append({
                            "base_version": base_version,
                            "versions": versions,
                            "path": path
                        })
                except:
                    pass
    
    if windows_sdk_versions:
        sdks["windows_sdk"] = {
            "installed": True,
            "sdk_info": windows_sdk_versions
        }
    else:
        sdks["windows_sdk"] = {
            "installed": False
        }
    
    # Android SDK
    android_sdk_path = os.environ.get("ANDROID_HOME", "")
    if not android_sdk_path:
        android_sdk_path = os.environ.get("ANDROID_SDK_ROOT", "")
    
    if android_sdk_path and os.path.exists(android_sdk_path):
        platforms = []
        platform_tools_version = "Unknown"
        
        # 检查已安装的Android平台版本
        platforms_path = os.path.join(android_sdk_path, "platforms")
        if os.path.exists(platforms_path):
            try:
                dirs = [d for d in os.listdir(platforms_path) 
                       if os.path.isdir(os.path.join(platforms_path, d))]
                platforms = [d.replace("android-", "") for d in dirs if d.startswith("android-")]
            except:
                pass
        
        # 检查platform-tools版本
        platform_tools_path = os.path.join(android_sdk_path, "platform-tools")
        if os.path.exists(platform_tools_path):
            try:
                result = subprocess.run([os.path.join(platform_tools_path, "adb"), "--version"], 
                                       capture_output=True, text=True)
                if result.returncode == 0:
                    match = re.search(r'Android Debug Bridge version (\d+\.\d+\.\d+)', result.stdout)
                    if match:
                        platform_tools_version = match.group(1)
            except:
                pass
        
        sdks["android_sdk"] = {
            "installed": True,
            "path": android_sdk_path,
            "platform_versions": platforms,
            "platform_tools_version": platform_tools_version
        }
    else:
        # 尝试在常见位置查找
        common_paths = [
            r"C:\Android\sdk",
            r"C:\Users\%USERNAME%\AppData\Local\Android\Sdk"
        ]
        
        username = os.environ.get("USERNAME", "")
        for i, path in enumerate(common_paths):
            if "%USERNAME%" in path:
                common_paths[i] = path.replace("%USERNAME%", username)
        
        android_sdk_found = False
        for path in common_paths:
            if os.path.exists(path):
                platforms = []
                platform_tools_version = "Unknown"
                
                # 检查已安装的Android平台版本
                platforms_path = os.path.join(path, "platforms")
                if os.path.exists(platforms_path):
                    try:
                        dirs = [d for d in os.listdir(platforms_path) 
                              if os.path.isdir(os.path.join(platforms_path, d))]
                        platforms = [d.replace("android-", "") for d in dirs if d.startswith("android-")]
                    except:
                        pass
                
                # 检查platform-tools版本
                platform_tools_path = os.path.join(path, "platform-tools")
                if os.path.exists(platform_tools_path):
                    try:
                        result = subprocess.run([os.path.join(platform_tools_path, "adb"), "--version"], 
                                              capture_output=True, text=True)
                        if result.returncode == 0:
                            match = re.search(r'Android Debug Bridge version (\d+\.\d+\.\d+)', result.stdout)
                            if match:
                                platform_tools_version = match.group(1)
                    except:
                        pass
                
                android_sdk_found = True
                sdks["android_sdk"] = {
                    "installed": True,
                    "path": path,
                    "platform_versions": platforms,
                    "platform_tools_version": platform_tools_version
                }
                break
        
        if not android_sdk_found:
            sdks["android_sdk"] = {
                "installed": False
            }
    
    # Java SDK (通过JAVA_HOME)
    java_home = os.environ.get("JAVA_HOME", "")
    if java_home and os.path.exists(java_home):
        # 获取版本
        java_version = "Unknown"
        try:
            result = subprocess.run([os.path.join(java_home, "bin", "java"), "-version"], 
                                   capture_output=True, text=True, stderr=subprocess.STDOUT)
            if result.returncode == 0:
                match = re.search(r'version "([^"]+)"', result.stdout)
                if match:
                    java_version = match.group(1)
        except:
            pass
        
        sdks["java_sdk"] = {
            "installed": True,
            "path": java_home,
            "version": java_version
        }
    else:
        sdks["java_sdk"] = {
            "installed": False
        }
    
    # .NET SDK (已在get_installed_programming_languages中收集)
    dotnet_info = {}
    try:
        result = subprocess.run(["dotnet", "--info"], capture_output=True, text=True)
        if result.returncode == 0:
            dotnet_info = {
                "installed": True,
                "info": result.stdout.strip()
            }
        else:
            dotnet_info = {
                "installed": False
            }
    except:
        dotnet_info = {
            "installed": False
        }
    
    sdks["dotnet_sdk"] = dotnet_info
    
    return sdks

def get_dev_environment_variables():
    """
    获取与开发相关的环境变量
    """
    # 感兴趣的环境变量
    env_var_prefixes = [
        "JAVA_", "ANDROID_", "PYTHON", "NODE_", "NPM_", "GO", "RUBY", "PHP_", "DOTNET_",
        "VS", "VISUAL_STUDIO", "VSCODE", "PATH", "HOME", "USER", "APPDATA", "MAVEN_",
        "GIT_", "DOCKER_", "KUBERNETES_", "AWS_", "AZURE_", "PROGRAMFILES"
    ]
    
    dev_env_vars = {}
    
    for key, value in os.environ.items():
        for prefix in env_var_prefixes:
            if key.upper().startswith(prefix) or key.upper() == prefix:
                dev_env_vars[key] = value
                break
    
    # 特别处理PATH变量，拆分为数组
    if "PATH" in dev_env_vars:
        dev_env_vars["PATH_ENTRIES"] = dev_env_vars["PATH"].split(os.pathsep)
    
    return dev_env_vars

def get_git_config():
    """
    获取Git配置信息
    """
    git_info = {
        "installed": False
    }
    
    try:
        # 检查Git是否安装
        version_result = subprocess.run(["git", "--version"], capture_output=True, text=True)
        if version_result.returncode == 0:
            git_info["installed"] = True
            git_info["version"] = version_result.stdout.strip()
            
            # 获取全局配置
            config_result = subprocess.run(["git", "config", "--global", "--list"], 
                                          capture_output=True, text=True)
            if config_result.returncode == 0:
                config = {}
                for line in config_result.stdout.strip().split("\n"):
                    if line and "=" in line:
                        key, value = line.split("=", 1)
                        config[key.strip()] = value.strip()
                
                if config:
                    git_info["global_config"] = config
    except:
        pass
    
    return git_info

def get_docker_info():
    """
    获取Docker信息
    """
    docker_info = {
        "installed": False
    }
    
    try:
        # 检查Docker是否安装
        version_result = subprocess.run(["docker", "--version"], capture_output=True, text=True)
        if version_result.returncode == 0:
            docker_info["installed"] = True
            docker_info["version"] = version_result.stdout.strip()
            
            # 获取系统信息
            try:
                info_result = subprocess.run(["docker", "info", "--format", "{{json .}}"], 
                                            capture_output=True, text=True)
                if info_result.returncode == 0:
                    info = json.loads(info_result.stdout)
                    # 提取关键信息
                    if "ServerVersion" in info:
                        docker_info["server_version"] = info["ServerVersion"]
                    if "OperatingSystem" in info:
                        docker_info["os"] = info["OperatingSystem"]
                    if "OSType" in info:
                        docker_info["os_type"] = info["OSType"]
                    if "Architecture" in info:
                        docker_info["architecture"] = info["Architecture"]
            except:
                pass
    except:
        pass
    
    return docker_info

def collect_all_dev_environment_info():
    """
    收集所有开发环境信息
    """
    dev_info = {
        "languages": get_installed_programming_languages(),
        "development_tools": get_development_tools(),
        "sdks": get_development_sdks(),
        "environment_variables": get_dev_environment_variables(),
        "git": get_git_config(),
        "docker": get_docker_info()
    }
    
    return dev_info

def save_dev_environment_info(output_dir, filename="dev_environment.json"):
    """
    保存开发环境信息到文件
    
    参数:
    - output_dir: 输出目录
    - filename: 输出文件名，默认为"dev_environment.json"
    
    返回:
    - 操作结果信息
    """
    try:
        # 确保输出目录存在
        if isinstance(output_dir, str):
            output_dir = Path(output_dir)
        
        os.makedirs(output_dir, exist_ok=True)
        output_file = output_dir / filename
        
        # 收集信息
        dev_info = collect_all_dev_environment_info()
        
        # 保存到JSON文件
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(dev_info, f, indent=2, ensure_ascii=False)
        
        return {
            "success": True,
            "message": f"开发环境信息已保存到 {output_file}",
            "file": str(output_file)
        }
        
    except Exception as e:
        return {
            "success": False,
            "message": "保存开发环境信息时出错",
            "error": str(e)
        } 