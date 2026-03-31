"""
配置管理
- 保存和加载用户配置
- 记住上次使用的路径
"""

import json
import os
from pathlib import Path


def get_config_dir():
    """获取配置文件目录"""
    if os.name == 'nt':  # Windows
        config_dir = Path(os.environ.get('APPDATA', '')) / 'WordTableFiller'
    else:  # macOS / Linux
        config_dir = Path.home() / '.config' / 'WordTableFiller'

    config_dir.mkdir(parents=True, exist_ok=True)
    return config_dir


def get_config_path():
    """获取配置文件路径"""
    return get_config_dir() / 'config.json'


def load_config():
    """加载配置"""
    config_path = get_config_path()
    if config_path.exists():
        try:
            with open(config_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception:
            pass
    return {}


def save_config(config):
    """保存配置"""
    config_path = get_config_path()
    try:
        with open(config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    except Exception:
        pass


def get_recent_paths():
    """获取最近使用的路径"""
    config = load_config()
    return {
        'docx_path': config.get('recent_docx_path', ''),
        'excel_folder': config.get('recent_excel_folder', ''),
        'output_path': config.get('recent_output_path', ''),
    }


def save_recent_paths(docx_path=None, excel_folder=None, output_path=None):
    """保存最近使用的路径"""
    config = load_config()

    if docx_path:
        config['recent_docx_path'] = docx_path
    if excel_folder:
        config['recent_excel_folder'] = excel_folder
    if output_path:
        config['recent_output_path'] = output_path

    save_config(config)


def get_mapping_file_path(docx_path, excel_folder):
    """
    根据 Word 文件和 Excel 文件夹生成映射文件路径
    """
    import hashlib

    # 用路径生成唯一标识
    key = f"{docx_path}|{excel_folder}"
    hash_str = hashlib.md5(key.encode()).hexdigest()[:12]

    # 用 Word 文件名作为前缀，更易识别
    docx_name = os.path.splitext(os.path.basename(docx_path))[0]
    filename = f"mapping_{docx_name}_{hash_str}.json"

    return get_config_dir() / filename


def save_mapping(docx_path, excel_folder, mapping_data):
    """保存映射数据到文件"""
    mapping_path = get_mapping_file_path(docx_path, excel_folder)
    try:
        with open(mapping_path, 'w', encoding='utf-8') as f:
            json.dump(mapping_data, f, ensure_ascii=False, indent=2)
        return True
    except Exception as e:
        print(f"保存映射失败: {e}")
        return False


def load_mapping(docx_path, excel_folder):
    """加载映射数据"""
    mapping_path = get_mapping_file_path(docx_path, excel_folder)
    if mapping_path.exists():
        try:
            with open(mapping_path, 'r', encoding='utf-8') as f:
                return json.load(f)
        except Exception as e:
            print(f"加载映射失败: {e}")
    return None


def delete_mapping(docx_path, excel_folder):
    """删除保存的映射文件"""
    mapping_path = get_mapping_file_path(docx_path, excel_folder)
    if mapping_path.exists():
        try:
            mapping_path.unlink()
            return True
        except Exception as e:
            print(f"删除映射失败: {e}")
            return False
    return True
