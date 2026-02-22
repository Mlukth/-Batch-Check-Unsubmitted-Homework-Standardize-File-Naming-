
import json
import os
import datetime
import hashlib
from typing import Dict, List, Optional

class ConfigManager:
    def __init__(self, config_dir: str = "config"):
        self.config_dir = config_dir
        self.app_config_file = os.path.join(config_dir, "app_config.json")
        self.format_config_file = os.path.join(config_dir, "format_config.json")
        self.current_vars_file = os.path.join(config_dir, "current_variables.json")  # 新增：存储当前变量

        self._ensure_config_dir()
        self._ensure_default_formats()

    def _ensure_config_dir(self):
        """确保配置目录存在"""
        if not os.path.exists(self.config_dir):
            os.makedirs(self.config_dir)

    def _ensure_default_formats(self):
        """确保有默认的格式配置"""
        if not os.path.exists(self.format_config_file):
            default_formats = {
                "标准格式(文件)": {"template": "{学号} {姓名}{扩展名}", "is_folder": False},
                "标准格式(文件夹)": {"template": "{学号} {姓名}", "is_folder": True},
            }
            self._save_json(self.format_config_file, default_formats)

    # --- 核心：管理当前花名册变量 ---
    def set_current_roster_columns(self, columns: List[str]):
        """保存当前花名册的列名（作为可用变量）"""
        # 确保包含必要的系统变量
        all_vars = list(set(columns + ['扩展名']))  # '扩展名'是系统变量
        self._save_json(self.current_vars_file, all_vars)

    def get_current_roster_columns(self) -> List[str]:
        """获取当前可用的变量列表"""
        return self._load_json(self.current_vars_file) or []

    # --- 格式管理（增删改查）---
    def get_format_names(self) -> List[str]:
        """获取所有格式名称"""
        formats = self._load_json(self.format_config_file)
        return list(formats.keys()) if formats else []

    def get_format_config(self, format_name: str) -> Optional[Dict]:
        """获取指定格式配置"""
        formats = self._load_json(self.format_config_file)
        return formats.get(format_name) if formats else None

    def save_format(self, format_name: str, format_config: Dict):
        """保存格式配置"""
        formats = self._load_json(self.format_config_file) or {}
        formats[format_name] = format_config
        self._save_json(self.format_config_file, formats)

    def delete_format(self, format_name: str):
        """删除格式配置"""
        formats = self._load_json(self.format_config_file)
        if formats and format_name in formats:
            del formats[format_name]
            self._save_json(self.format_config_file, formats)

    # --- 应用配置 ---
    def load_app_config(self) -> Dict:
        """加载应用配置"""
        return self._load_json(self.app_config_file) or {}

    def save_app_config(self, config: Dict):
        """保存应用配置"""
        self._save_json(self.app_config_file, config)

    def _sort_folders_by_order(self, folders, order_mapping):
        """
        根据 order_mapping 中的序号对文件夹列表进行排序
        返回按序号升序排列的文件夹列表
        """
        # 过滤出有序号的文件夹并排序
        valid_folders = [(order_mapping.get(f, 999), f) for f in folders if f in order_mapping]
        # 按序号升序排序
        valid_folders.sort(key=lambda x: x[0])
        return [f for _, f in valid_folders]

    # --- 文件夹配置管理 ---
    def save_folder_config(self, parent_dir: str, config: dict):
        """保存文件夹选择配置"""
        # 使用父文件夹路径的哈希值作为配置键（确保唯一性）
        dir_hash = hashlib.md5(parent_dir.encode('utf-8')).hexdigest()[:8]
        config_key = f"folder_config_{dir_hash}"
        
        # 加载所有文件夹配置 - 这里必须返回字典而不是列表
        all_configs = self._load_folder_configs()
        
        # 确保 all_configs 是字典
        if not isinstance(all_configs, dict):
            all_configs = {}
        
        # 确保文件夹顺序是按序号升序排列的
        if 'selected_folders' in config and 'order_mapping' in config:
            selected = config['selected_folders']
            order_mapping = config['order_mapping']
            # 按序号升序重新排序
            sorted_folders = self._sort_folders_by_order(selected, order_mapping)
            config['selected_folders'] = sorted_folders
            # 同时更新 folder_order 以确保一致性
            config['folder_order'] = sorted_folders
        
        # 保存配置
        all_configs[config_key] = {
            'parent_dir': parent_dir,
            'config': config,
            'timestamp': datetime.datetime.now().isoformat()
        }
        
        self._save_folder_configs(all_configs)

    def load_folder_config(self, parent_dir: str) -> Optional[dict]:
        """加载文件夹选择配置 - 确保返回已排序的文件夹列表"""
        dir_hash = hashlib.md5(parent_dir.encode('utf-8')).hexdigest()[:8]
        config_key = f"folder_config_{dir_hash}"
        
        all_configs = self._load_folder_configs()
        
        if not isinstance(all_configs, dict):
            return None
        
        if config_key in all_configs:
            config = all_configs[config_key]['config']
            # 加载时再次确保顺序正确
            if 'selected_folders' in config and 'order_mapping' in config:
                selected = config['selected_folders']
                order_mapping = config['order_mapping']
                # 确保按序号升序排列
                sorted_folders = self._sort_folders_by_order(selected, order_mapping)
                config['selected_folders'] = sorted_folders
                config['folder_order'] = sorted_folders
            return config
        return None

    def _load_folder_configs(self) -> dict:
        """专用方法：加载文件夹配置数据"""
        # 为文件夹配置创建专门的文件
        folder_config_file = os.path.join(self.config_dir, "folder_configs.json")
        
        try:
            if os.path.exists(folder_config_file):
                with open(folder_config_file, 'r', encoding='utf-8') as f:
                    data = json.load(f)
                    # 确保返回的是字典
                    if isinstance(data, dict):
                        return data
        except Exception:
            pass
        
        # 如果文件不存在或读取失败，返回空字典
        return {}

    def _save_folder_configs(self, data: dict):
        """专用方法：保存文件夹配置数据"""
        folder_config_file = os.path.join(self.config_dir, "folder_configs.json")
        
        try:
            with open(folder_config_file, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            raise Exception(f"保存文件夹配置失败: {str(e)}")

    # --- 底层JSON读写 ---
    def _load_json(self, filepath: str) -> Optional[Dict]:
        """加载JSON文件"""
        try:
            if os.path.exists(filepath):
                with open(filepath, 'r', encoding='utf-8') as f:
                    return json.load(f)
        except Exception:
            pass
        return None

    def _save_json(self, filepath: str, data: Dict):
        """保存JSON文件"""
        try:
            with open(filepath, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
        except Exception as e:
            raise Exception(f"保存配置文件失败: {str(e)}")