"""
配置文件加载器
负责加载和管理应用程序配置
"""
import json
import os
from typing import Dict, Any, Optional
from pathlib import Path


class ConfigValidationError(Exception):
    """配置验证错误"""
    pass


class ConfigLoader:
    """配置文件加载器类"""
    
    # 默认配置
    DEFAULT_CONFIG = {
        "app_config": {
            "window_title": "Excel岗位分数查询工具",
            "window_size": "800x600",
            "default_output_filename": "岗位最低分数汇总.xlsx",
            "theme": "default",
            "language": "zh_CN"
        },
        "matching_config": {
            "fuzzy_threshold": 0.8,
            "ignore_case": True,
            "remove_spaces": True,
            "enable_fuzzy_matching": True,
            "max_match_attempts": 3
        },
        "logging_config": {
            "level": "INFO",
            "log_file": "app.log",
            "max_file_size": "10MB",
            "backup_count": 5,
            "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        }
    }
    
    # 配置验证规则
    VALIDATION_RULES = {
        "app_config": {
            "window_title": str,
            "window_size": str,
            "default_output_filename": str,
            "theme": str,
            "language": str
        },
        "matching_config": {
            "fuzzy_threshold": (float, int),
            "ignore_case": bool,
            "remove_spaces": bool,
            "enable_fuzzy_matching": bool,
            "max_match_attempts": int
        },
        "logging_config": {
            "level": str,
            "log_file": str,
            "max_file_size": str,
            "backup_count": int,
            "format": str
        }
    }
    
    def __init__(self, config_file: str = "config.json"):
        """
        初始化配置加载器
        
        Args:
            config_file: 配置文件路径
        """
        self.config_file = Path(config_file)
        self._config = None
    
    def load_config(self) -> Dict[str, Any]:
        """
        加载配置文件，如果文件不存在则创建默认配置
        
        Returns:
            配置字典
            
        Raises:
            json.JSONDecodeError: 配置文件格式错误
            ConfigValidationError: 配置验证失败
        """
        if not self.config_file.exists():
            # 创建默认配置文件
            import copy
            self._config = copy.deepcopy(self.DEFAULT_CONFIG)
            self.save_config()
            return self._config
        
        try:
            with open(self.config_file, 'r', encoding='utf-8') as f:
                loaded_config = json.load(f)
            
            # 合并默认配置和加载的配置
            self._config = self._merge_with_defaults(loaded_config)
            
            # 验证配置
            self._validate_config(self._config)
            
            # 如果配置有更新，保存回文件
            if self._config != loaded_config:
                self.save_config()
            
            return self._config
            
        except json.JSONDecodeError as e:
            raise json.JSONDecodeError(f"配置文件格式错误: {e}", e.doc, e.pos)
    
    def save_config(self) -> None:
        """
        保存配置到文件
        
        Raises:
            IOError: 文件写入失败
        """
        if self._config is None:
            self._config = self.DEFAULT_CONFIG.copy()
        
        try:
            # 确保目录存在
            self.config_file.parent.mkdir(parents=True, exist_ok=True)
            
            with open(self.config_file, 'w', encoding='utf-8') as f:
                json.dump(self._config, f, ensure_ascii=False, indent=4)
        except IOError as e:
            raise IOError(f"保存配置文件失败: {e}")
    
    def _merge_with_defaults(self, loaded_config: Dict[str, Any]) -> Dict[str, Any]:
        """
        将加载的配置与默认配置合并
        
        Args:
            loaded_config: 从文件加载的配置
            
        Returns:
            合并后的配置
        """
        import copy
        merged_config = copy.deepcopy(self.DEFAULT_CONFIG)
        
        for section, section_config in loaded_config.items():
            if section in merged_config:
                if isinstance(section_config, dict):
                    merged_config[section].update(section_config)
                else:
                    merged_config[section] = section_config
            else:
                merged_config[section] = section_config
        
        return merged_config
    
    def _validate_config(self, config: Dict[str, Any]) -> None:
        """
        验证配置的有效性
        
        Args:
            config: 要验证的配置
            
        Raises:
            ConfigValidationError: 配置验证失败
        """
        for section_name, section_rules in self.VALIDATION_RULES.items():
            if section_name not in config:
                continue
            
            section_config = config[section_name]
            if not isinstance(section_config, dict):
                raise ConfigValidationError(f"配置节 '{section_name}' 必须是字典类型")
            
            for key, expected_type in section_rules.items():
                if key in section_config:
                    value = section_config[key]
                    if isinstance(expected_type, tuple):
                        if not isinstance(value, expected_type):
                            raise ConfigValidationError(
                                f"配置项 '{section_name}.{key}' 类型错误，期望 {expected_type}，实际 {type(value)}"
                            )
                    else:
                        if not isinstance(value, expected_type):
                            raise ConfigValidationError(
                                f"配置项 '{section_name}.{key}' 类型错误，期望 {expected_type}，实际 {type(value)}"
                            )
        
        # 特殊验证规则
        self._validate_special_rules(config)
    
    def _validate_special_rules(self, config: Dict[str, Any]) -> None:
        """
        验证特殊规则
        
        Args:
            config: 要验证的配置
            
        Raises:
            ConfigValidationError: 配置验证失败
        """
        # 验证模糊匹配阈值
        if "matching_config" in config:
            threshold = config["matching_config"].get("fuzzy_threshold", 0.8)
            if not (0.0 <= threshold <= 1.0):
                raise ConfigValidationError("fuzzy_threshold 必须在 0.0 到 1.0 之间")
        
        # 验证日志级别
        if "logging_config" in config:
            valid_levels = ["DEBUG", "INFO", "WARNING", "ERROR", "CRITICAL"]
            level = config["logging_config"].get("level", "INFO")
            if level not in valid_levels:
                raise ConfigValidationError(f"日志级别必须是 {valid_levels} 中的一个")
        
        # 验证窗口大小格式
        if "app_config" in config:
            window_size = config["app_config"].get("window_size", "800x600")
            if not self._is_valid_window_size(window_size):
                raise ConfigValidationError("window_size 格式必须是 'widthxheight'，如 '800x600'")
    
    def _is_valid_window_size(self, size_str: str) -> bool:
        """
        验证窗口大小格式
        
        Args:
            size_str: 窗口大小字符串
            
        Returns:
            是否有效
        """
        try:
            parts = size_str.split('x')
            if len(parts) != 2:
                return False
            width, height = int(parts[0]), int(parts[1])
            return width > 0 and height > 0
        except (ValueError, IndexError):
            return False
    
    def get_app_config(self) -> Dict[str, Any]:
        """获取应用程序配置"""
        if self._config is None:
            self.load_config()
        return self._config.get("app_config", {})
    
    def get_matching_config(self) -> Dict[str, Any]:
        """获取匹配配置"""
        if self._config is None:
            self.load_config()
        return self._config.get("matching_config", {})
    
    def get_logging_config(self) -> Dict[str, Any]:
        """获取日志配置"""
        if self._config is None:
            self.load_config()
        return self._config.get("logging_config", {})
    
    def get_config_value(self, section: str, key: str, default: Any = None) -> Any:
        """
        获取指定配置值
        
        Args:
            section: 配置节名称
            key: 配置键名
            default: 默认值
            
        Returns:
            配置值
        """
        if self._config is None:
            self.load_config()
        
        return self._config.get(section, {}).get(key, default)
    
    def set_config_value(self, section: str, key: str, value: Any) -> None:
        """
        设置配置值
        
        Args:
            section: 配置节名称
            key: 配置键名
            value: 配置值
        """
        if self._config is None:
            self.load_config()
        
        if section not in self._config:
            self._config[section] = {}
        
        self._config[section][key] = value
    
    def update_config(self, updates: Dict[str, Dict[str, Any]]) -> None:
        """
        批量更新配置
        
        Args:
            updates: 要更新的配置字典
            
        Raises:
            ConfigValidationError: 配置验证失败
        """
        if self._config is None:
            self.load_config()
        
        # 创建临时配置进行验证
        import copy
        temp_config = copy.deepcopy(self._config)
        
        for section, section_updates in updates.items():
            if section not in temp_config:
                temp_config[section] = {}
            temp_config[section].update(section_updates)
        
        # 验证更新后的配置
        self._validate_config(temp_config)
        
        # 验证通过，应用更新
        self._config = temp_config
    
    def reset_to_defaults(self) -> None:
        """重置配置为默认值"""
        import copy
        self._config = copy.deepcopy(self.DEFAULT_CONFIG)
    
    def get_config_info(self) -> Dict[str, Any]:
        """
        获取配置信息
        
        Returns:
            配置信息字典
        """
        if self._config is None:
            self.load_config()
        
        return {
            "config_file": str(self.config_file),
            "file_exists": self.config_file.exists(),
            "sections": list(self._config.keys()) if self._config else [],
            "total_keys": sum(len(section) for section in self._config.values() if isinstance(section, dict)) if self._config else 0
        }