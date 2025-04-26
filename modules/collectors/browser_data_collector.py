import os
import json
import sqlite3
import shutil
import subprocess
from pathlib import Path
from datetime import datetime, timedelta

def get_chrome_profile_paths():
    """
    获取Chrome浏览器配置文件路径
    
    返回:
    - Chrome配置文件路径列表
    """
    chrome_profiles = []
    
    # Chrome默认配置文件路径
    local_app_data = os.environ.get('LOCALAPPDATA', '')
    chrome_path = os.path.join(local_app_data, 'Google', 'Chrome', 'User Data')
    
    if os.path.exists(chrome_path):
        # 查找所有配置文件目录
        for item in os.listdir(chrome_path):
            if item.startswith('Profile ') or item == 'Default':
                profile_path = os.path.join(chrome_path, item)
                if os.path.isdir(profile_path):
                    chrome_profiles.append({
                        'name': item,
                        'path': profile_path
                    })
    
    return chrome_profiles

def get_edge_profile_paths():
    """
    获取Edge浏览器配置文件路径
    
    返回:
    - Edge配置文件路径列表
    """
    edge_profiles = []
    
    # Edge默认配置文件路径
    local_app_data = os.environ.get('LOCALAPPDATA', '')
    edge_path = os.path.join(local_app_data, 'Microsoft', 'Edge', 'User Data')
    
    if os.path.exists(edge_path):
        # 查找所有配置文件目录
        for item in os.listdir(edge_path):
            if item.startswith('Profile ') or item == 'Default':
                profile_path = os.path.join(edge_path, item)
                if os.path.isdir(profile_path):
                    edge_profiles.append({
                        'name': item,
                        'path': profile_path
                    })
    
    return edge_profiles

def get_firefox_profile_paths():
    """
    获取Firefox浏览器配置文件路径
    
    返回:
    - Firefox配置文件路径列表
    """
    firefox_profiles = []
    
    # Firefox配置文件可能位置
    app_data = os.environ.get('APPDATA', '')
    firefox_path = os.path.join(app_data, 'Mozilla', 'Firefox', 'Profiles')
    
    if os.path.exists(firefox_path):
        # 查找所有配置文件目录
        for item in os.listdir(firefox_path):
            profile_path = os.path.join(firefox_path, item)
            if os.path.isdir(profile_path):
                firefox_profiles.append({
                    'name': item,
                    'path': profile_path
                })
    
    return firefox_profiles

def get_chrome_bookmarks(profile_path):
    """
    获取Chrome浏览器书签
    
    参数:
    - profile_path: Chrome配置文件路径
    
    返回:
    - 书签信息字典
    """
    bookmarks = []
    
    try:
        # Chrome书签文件路径
        bookmarks_file = os.path.join(profile_path, 'Bookmarks')
        
        if os.path.exists(bookmarks_file):
            with open(bookmarks_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # 解析书签
            def parse_bookmark_node(node, folder=""):
                if node['type'] == 'url':
                    return [{
                        'name': node.get('name', ''),
                        'url': node.get('url', ''),
                        'date_added': node.get('date_added', ''),
                        'folder': folder
                    }]
                elif node['type'] == 'folder':
                    result = []
                    current_folder = folder + '/' + node['name'] if folder else node['name']
                    for child in node.get('children', []):
                        result.extend(parse_bookmark_node(child, current_folder))
                    return result
                return []
            
            # 处理书签栏和其他书签
            for root in ['bookmark_bar', 'other', 'synced']:
                if root in data.get('roots', {}):
                    root_node = data['roots'][root]
                    folder_name = root_node.get('name', root)
                    bookmarks.extend(parse_bookmark_node(root_node, folder_name))
    except Exception as e:
        print(f"获取Chrome书签时出错: {e}")
    
    return bookmarks

def get_edge_bookmarks(profile_path):
    """
    获取Edge浏览器书签
    
    参数:
    - profile_path: Edge配置文件路径
    
    返回:
    - 书签信息字典
    """
    # Edge使用与Chrome相同的书签格式
    return get_chrome_bookmarks(profile_path)

def get_firefox_bookmarks(profile_path):
    """
    获取Firefox浏览器书签
    
    参数:
    - profile_path: Firefox配置文件路径
    
    返回:
    - 书签信息字典
    """
    bookmarks = []
    
    try:
        # Firefox书签数据库文件
        places_db = os.path.join(profile_path, 'places.sqlite')
        
        if os.path.exists(places_db):
            # 创建临时副本以避免数据库锁定问题
            temp_db = os.path.join(os.environ.get('TEMP', ''), 'temp_places.sqlite')
            shutil.copy2(places_db, temp_db)
            
            # 连接数据库
            conn = sqlite3.connect(temp_db)
            cursor = conn.cursor()
            
            # 查询书签
            query = """
            SELECT b.title, p.url, b.dateAdded, c.title
            FROM moz_bookmarks b
            JOIN moz_places p ON b.fk = p.id
            LEFT JOIN moz_bookmarks c ON b.parent = c.id
            WHERE b.type = 1
            """
            
            cursor.execute(query)
            for row in cursor.fetchall():
                title, url, date_added, folder = row
                bookmarks.append({
                    'name': title,
                    'url': url,
                    'date_added': date_added,
                    'folder': folder or 'Root'
                })
            
            # 关闭连接
            cursor.close()
            conn.close()
            
            # 删除临时文件
            try:
                os.remove(temp_db)
            except:
                pass
    except Exception as e:
        print(f"获取Firefox书签时出错: {e}")
    
    return bookmarks

def get_chrome_extensions(profile_path):
    """
    获取Chrome浏览器扩展插件
    
    参数:
    - profile_path: Chrome配置文件路径
    
    返回:
    - 扩展插件信息列表
    """
    extensions = []
    
    try:
        # Chrome扩展目录
        ext_dir = os.path.join(profile_path, 'Extensions')
        
        if os.path.exists(ext_dir):
            for ext_id in os.listdir(ext_dir):
                ext_path = os.path.join(ext_dir, ext_id)
                if os.path.isdir(ext_path):
                    # 查找最新版本
                    versions = [v for v in os.listdir(ext_path) if os.path.isdir(os.path.join(ext_path, v))]
                    if versions:
                        latest_version = sorted(versions)[-1]
                        manifest_path = os.path.join(ext_path, latest_version, 'manifest.json')
                        
                        if os.path.exists(manifest_path):
                            try:
                                with open(manifest_path, 'r', encoding='utf-8') as f:
                                    manifest = json.load(f)
                                
                                extensions.append({
                                    'id': ext_id,
                                    'name': manifest.get('name', ''),
                                    'version': manifest.get('version', ''),
                                    'description': manifest.get('description', ''),
                                    'permissions': manifest.get('permissions', [])
                                })
                            except Exception as e:
                                print(f"解析扩展 {ext_id} 的manifest.json时出错: {e}")
    except Exception as e:
        print(f"获取Chrome扩展时出错: {e}")
    
    return extensions

def get_edge_extensions(profile_path):
    """
    获取Edge浏览器扩展插件
    
    参数:
    - profile_path: Edge配置文件路径
    
    返回:
    - 扩展插件信息列表
    """
    # Edge使用与Chrome相同的扩展格式
    return get_chrome_extensions(profile_path)

def get_firefox_extensions(profile_path):
    """
    获取Firefox浏览器扩展插件
    
    参数:
    - profile_path: Firefox配置文件路径
    
    返回:
    - 扩展插件信息列表
    """
    extensions = []
    
    try:
        # Firefox扩展配置文件
        ext_file = os.path.join(profile_path, 'extensions.json')
        
        if os.path.exists(ext_file):
            with open(ext_file, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            for addon in data.get('addons', []):
                extensions.append({
                    'id': addon.get('id', ''),
                    'name': addon.get('defaultLocale', {}).get('name', '') or addon.get('name', ''),
                    'version': addon.get('version', ''),
                    'description': addon.get('defaultLocale', {}).get('description', '') or addon.get('description', ''),
                    'enabled': addon.get('active', False),
                    'type': addon.get('type', '')
                })
    except Exception as e:
        print(f"获取Firefox扩展时出错: {e}")
    
    return extensions

def get_chrome_history(profile_path):
    """
    获取Chrome浏览器历史记录
    
    参数:
    - profile_path: Chrome配置文件路径
    
    返回:
    - 历史记录信息列表
    """
    history = []
    
    try:
        # Chrome历史记录文件
        history_file = os.path.join(profile_path, 'History')
        
        if os.path.exists(history_file):
            # 创建临时副本以避免数据库锁定问题
            temp_history = os.path.join(os.path.dirname(history_file), 'history_temp')
            shutil.copy2(history_file, temp_history)
            
            # 连接到SQLite数据库
            conn = sqlite3.connect(temp_history)
            cursor = conn.cursor()
            
            # 查询最近100条历史记录
            cursor.execute("""
                SELECT url, title, last_visit_time, visit_count
                FROM urls
                ORDER BY last_visit_time DESC
                LIMIT 100
            """)
            
            for row in cursor.fetchall():
                url, title, last_visit_time, visit_count = row
                
                # Chrome存储时间为自1601年1月1日以来的微秒数
                # 转换为ISO格式的日期时间
                chrome_epoch = datetime(1601, 1, 1)
                last_visit = chrome_epoch + timedelta(microseconds=last_visit_time)
                
                history.append({
                    'url': url,
                    'title': title,
                    'last_visit': last_visit.isoformat(),
                    'visit_count': visit_count
                })
            
            # 关闭连接
            cursor.close()
            conn.close()
            
            # 删除临时文件
            try:
                os.remove(temp_history)
            except:
                pass
    except Exception as e:
        print(f"获取Chrome历史记录时出错: {e}")
    
    return history

def get_edge_history(profile_path):
    """
    获取Edge浏览器历史记录
    
    参数:
    - profile_path: Edge配置文件路径
    
    返回:
    - 历史记录信息列表
    """
    # Edge使用与Chrome相同的历史记录格式
    return get_chrome_history(profile_path)

def get_firefox_history(profile_path):
    """
    获取Firefox浏览器历史记录
    
    参数:
    - profile_path: Firefox配置文件路径
    
    返回:
    - 历史记录信息列表
    """
    history = []
    
    try:
        # Firefox历史记录文件
        history_file = os.path.join(profile_path, 'places.sqlite')
        
        if os.path.exists(history_file):
            # 创建临时副本以避免数据库锁定问题
            temp_history = os.path.join(os.path.dirname(history_file), 'places_temp.sqlite')
            shutil.copy2(history_file, temp_history)
            
            # 连接到SQLite数据库
            conn = sqlite3.connect(temp_history)
            cursor = conn.cursor()
            
            # 查询最近100条历史记录
            cursor.execute("""
                SELECT p.url, p.title, h.visit_date, p.visit_count
                FROM moz_places p
                JOIN moz_historyvisits h ON p.id = h.place_id
                ORDER BY h.visit_date DESC
                LIMIT 100
            """)
            
            for row in cursor.fetchall():
                url, title, visit_date, visit_count = row
                
                # Firefox存储时间为自1970年1月1日以来的微秒数
                # 转换为ISO格式的日期时间
                firefox_epoch = datetime(1970, 1, 1)
                last_visit = firefox_epoch + timedelta(microseconds=visit_date)
                
                history.append({
                    'url': url,
                    'title': title,
                    'last_visit': last_visit.isoformat(),
                    'visit_count': visit_count
                })
            
            # 关闭连接
            cursor.close()
            conn.close()
            
            # 删除临时文件
            try:
                os.remove(temp_history)
            except:
                pass
    except Exception as e:
        print(f"获取Firefox历史记录时出错: {e}")
    
    return history

def get_chrome_cookies(profile_path):
    """
    获取Chrome浏览器Cookies
    
    参数:
    - profile_path: Chrome配置文件路径
    
    返回:
    - Cookies摘要信息列表
    """
    cookies_summary = {}
    
    try:
        # Chrome Cookies文件
        cookies_file = os.path.join(profile_path, 'Cookies')
        
        if os.path.exists(cookies_file):
            # 创建临时副本以避免数据库锁定问题
            temp_cookies = os.path.join(os.path.dirname(cookies_file), 'cookies_temp')
            shutil.copy2(cookies_file, temp_cookies)
            
            # 连接到SQLite数据库
            conn = sqlite3.connect(temp_cookies)
            cursor = conn.cursor()
            
            # 查询域名和Cookie数量
            cursor.execute("""
                SELECT host_key, COUNT(*) as cookie_count
                FROM cookies
                GROUP BY host_key
                ORDER BY cookie_count DESC
                LIMIT 50
            """)
            
            for row in cursor.fetchall():
                host_key, cookie_count = row
                cookies_summary[host_key] = cookie_count
            
            # 关闭连接
            cursor.close()
            conn.close()
            
            # 删除临时文件
            try:
                os.remove(temp_cookies)
            except:
                pass
    except Exception as e:
        print(f"获取Chrome Cookies时出错: {e}")
    
    return cookies_summary

def get_edge_cookies(profile_path):
    """
    获取Edge浏览器Cookies
    
    参数:
    - profile_path: Edge配置文件路径
    
    返回:
    - Cookies摘要信息列表
    """
    # Edge使用与Chrome相同的Cookies格式
    return get_chrome_cookies(profile_path)

def get_firefox_cookies(profile_path):
    """
    获取Firefox浏览器Cookies
    
    参数:
    - profile_path: Firefox配置文件路径
    
    返回:
    - Cookies摘要信息列表
    """
    cookies_summary = {}
    
    try:
        # Firefox Cookies文件
        cookies_file = os.path.join(profile_path, 'cookies.sqlite')
        
        if os.path.exists(cookies_file):
            # 创建临时副本以避免数据库锁定问题
            temp_cookies = os.path.join(os.path.dirname(cookies_file), 'cookies_temp.sqlite')
            shutil.copy2(cookies_file, temp_cookies)
            
            # 连接到SQLite数据库
            conn = sqlite3.connect(temp_cookies)
            cursor = conn.cursor()
            
            # 查询域名和Cookie数量
            cursor.execute("""
                SELECT host, COUNT(*) as cookie_count
                FROM moz_cookies
                GROUP BY host
                ORDER BY cookie_count DESC
                LIMIT 50
            """)
            
            for row in cursor.fetchall():
                host, cookie_count = row
                cookies_summary[host] = cookie_count
            
            # 关闭连接
            cursor.close()
            conn.close()
            
            # 删除临时文件
            try:
                os.remove(temp_cookies)
            except:
                pass
    except Exception as e:
        print(f"获取Firefox Cookies时出错: {e}")
    
    return cookies_summary

def collect_all_browser_data():
    """
    收集所有浏览器数据
    
    返回:
    - 包含所有浏览器数据的字典
    """
    browser_data = {
        "collection_time": datetime.now().isoformat(),
        "chrome": {
            "profiles": [],
        },
        "edge": {
            "profiles": [],
        },
        "firefox": {
            "profiles": [],
        }
    }
    
    # 收集Chrome数据
    chrome_profiles = get_chrome_profile_paths()
    for profile in chrome_profiles:
        profile_data = {
            "name": profile['name'],
            "path": profile['path'],
            "bookmarks": get_chrome_bookmarks(profile['path']),
            "extensions": get_chrome_extensions(profile['path']),
            "history": get_chrome_history(profile['path']),
            "cookies": get_chrome_cookies(profile['path'])
        }
        browser_data["chrome"]["profiles"].append(profile_data)
    
    # 收集Edge数据
    edge_profiles = get_edge_profile_paths()
    for profile in edge_profiles:
        profile_data = {
            "name": profile['name'],
            "path": profile['path'],
            "bookmarks": get_edge_bookmarks(profile['path']),
            "extensions": get_edge_extensions(profile['path']),
            "history": get_edge_history(profile['path']),
            "cookies": get_edge_cookies(profile['path'])
        }
        browser_data["edge"]["profiles"].append(profile_data)
    
    # 收集Firefox数据
    firefox_profiles = get_firefox_profile_paths()
    for profile in firefox_profiles:
        profile_data = {
            "name": profile['name'],
            "path": profile['path'],
            "bookmarks": get_firefox_bookmarks(profile['path']),
            "extensions": get_firefox_extensions(profile['path']),
            "history": get_firefox_history(profile['path']),
            "cookies": get_firefox_cookies(profile['path'])
        }
        browser_data["firefox"]["profiles"].append(profile_data)
    
    return browser_data

def save_browser_data(output_dir, filename="browser_data.json"):
    """
    保存浏览器数据到JSON文件
    
    参数:
    - output_dir: 输出目录
    - filename: 输出文件名
    
    返回:
    - 保存结果信息
    """
    try:
        # 确保输出目录存在
        os.makedirs(output_dir, exist_ok=True)
        
        # 收集浏览器数据
        browser_data = collect_all_browser_data()
        
        # 保存浏览器数据
        output_file = os.path.join(output_dir, filename)
        with open(output_file, 'w', encoding='utf-8') as f:
            json.dump(browser_data, f, indent=2, ensure_ascii=False)
        
        return {
            "success": True,
            "message": f"浏览器数据已保存到 {output_file}",
            "file": output_file
        }
    except Exception as e:
        return {
            "success": False,
            "message": f"保存浏览器数据时出错: {str(e)}",
            "error": str(e)
        } 