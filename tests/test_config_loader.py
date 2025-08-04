"""
配置加载器测试
"""
import json
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, mock_open

from utils.config_loader import ConfigLoader, ConfigValidationError


class TestConfigLoader(unittest.TestCase):
    """配置加载器测试类"""
    
    def setUp(self):
        """测试前准备"""
        self.temp_dir = tempfile.mkdtemp()
        self.config_file = Path(self.temp_dir) / "test_config.json"
        self.loader = ConfigLoader(str(self.config_file))
    
    def tearDown(self):
        """测试后清理"""
        if self.config_file.exists():
            self.config_file.unlink()
        os.rmdir(self.temp_dir)
    
    def test_load_default_config_when_file_not_exists(self):
        """测试文件不存在时加载默认配置"""
        config = self.loader.load_config()
        
        # 验证默认配置结构
        self.assertIn("app_config", config)
        self.assertIn("matching_config", config)
        self.assertIn("logging_config", config)
        
        # 验证默认值
        self.assertEqual(config["app_config"]["window_title"], "Excel岗位分数查询工具")
        self.assertEqual(config["matching_config"]["fuzzy_threshold"], 0.8)
        self.assertEqual(config["logging_config"]["level"], "INFO")
        
        # 验证文件已创建
        self.assertTrue(self.config_file.exists())
    
    def test_load_existing_config(self):
        """测试加载现有配置文件"""
        # 创建测试配置文件
        test_config = {
            "app_config": {
                "window_title": "测试标题",
                "window_size": "1024x768"
            },
            "matching_config": {
                "fuzzy_threshold": 0.9
            }
        }
        
        with open(self.config_file, 'w', encoding='utf-8') as f:
            json.dump(test_config, f)
        
        config = self.loader.load_config()
        
        # 验证加载的配置
        self.assertEqual(config["app_config"]["window_title"], "测试标题")
        self.assertEqual(config["app_config"]["window_size"], "1024x768")
        self.assertEqual(config["matching_config"]["fuzzy_threshold"], 0.9)
        
        # 验证默认值被合并
        self.assertIn("default_output_filename", config["app_config"])
        self.assertIn("ignore_case", config["matching_config"])
    
    def test_invalid_json_format(self):
        """测试无效JSON格式"""
        # 创建无效JSON文件
        with open(self.config_file, 'w', encoding='utf-8') as f:
            f.write("{ invalid json }")
        
        with self.assertRaises(json.JSONDecodeError):
            self.loader.load_config()
    
    def test_config_validation_success(self):
        """测试配置验证成功"""
        valid_config = {
            "app_config": {
                "window_title": "测试",
                "window_size": "800x600"
            },
            "matching_config": {
                "fuzzy_threshold": 0.8,
                "ignore_case": True
            },
            "logging_config": {
                "level": "DEBUG",
                "log_file": "test.log"
            }
        }
        
        # 应该不抛出异常
        self.loader._validate_config(valid_config)
    
    def test_config_validation_type_error(self):
        """测试配置类型验证错误"""
        invalid_config = {
            "app_config": {
                "window_title": 123  # 应该是字符串
            }
        }
        
        with self.assertRaises(ConfigValidationError) as cm:
            self.loader._validate_config(invalid_config)
        
        self.assertIn("类型错误", str(cm.exception))
    
    def test_config_validation_fuzzy_threshold_range(self):
        """测试模糊匹配阈值范围验证"""
        invalid_config = {
            "matching_config": {
                "fuzzy_threshold": 1.5  # 超出范围
            }
        }
        
        with self.assertRaises(ConfigValidationError) as cm:
            self.loader._validate_config(invalid_config)
        
        self.assertIn("fuzzy_threshold", str(cm.exception))
    
    def test_config_validation_log_level(self):
        """测试日志级别验证"""
        invalid_config = {
            "logging_config": {
                "level": "INVALID_LEVEL"
            }
        }
        
        with self.assertRaises(ConfigValidationError) as cm:
            self.loader._validate_config(invalid_config)
        
        self.assertIn("日志级别", str(cm.exception))
    
    def test_config_validation_window_size(self):
        """测试窗口大小格式验证"""
        invalid_config = {
            "app_config": {
                "window_size": "invalid_size"
            }
        }
        
        with self.assertRaises(ConfigValidationError) as cm:
            self.loader._validate_config(invalid_config)
        
        self.assertIn("window_size", str(cm.exception))
    
    def test_window_size_validation(self):
        """测试窗口大小验证方法"""
        # 有效格式
        self.assertTrue(self.loader._is_valid_window_size("800x600"))
        self.assertTrue(self.loader._is_valid_window_size("1920x1080"))
        
        # 无效格式
        self.assertFalse(self.loader._is_valid_window_size("800"))
        self.assertFalse(self.loader._is_valid_window_size("800x"))
        self.assertFalse(self.loader._is_valid_window_size("800xabc"))
        self.assertFalse(self.loader._is_valid_window_size("0x600"))
    
    def test_get_config_sections(self):
        """测试获取配置节"""
        self.loader.load_config()
        
        app_config = self.loader.get_app_config()
        matching_config = self.loader.get_matching_config()
        logging_config = self.loader.get_logging_config()
        
        self.assertIsInstance(app_config, dict)
        self.assertIsInstance(matching_config, dict)
        self.assertIsInstance(logging_config, dict)
        
        self.assertIn("window_title", app_config)
        self.assertIn("fuzzy_threshold", matching_config)
        self.assertIn("level", logging_config)
    
    def test_get_config_value(self):
        """测试获取配置值"""
        self.loader.load_config()
        
        # 获取存在的值
        title = self.loader.get_config_value("app_config", "window_title")
        self.assertEqual(title, "Excel岗位分数查询工具")
        
        # 获取不存在的值，使用默认值
        unknown = self.loader.get_config_value("app_config", "unknown_key", "default")
        self.assertEqual(unknown, "default")
        
        # 获取不存在的节
        unknown_section = self.loader.get_config_value("unknown_section", "key", "default")
        self.assertEqual(unknown_section, "default")
    
    def test_set_config_value(self):
        """测试设置配置值"""
        self.loader.load_config()
        
        # 设置现有节的新值
        self.loader.set_config_value("app_config", "new_key", "new_value")
        value = self.loader.get_config_value("app_config", "new_key")
        self.assertEqual(value, "new_value")
        
        # 设置新节的值
        self.loader.set_config_value("new_section", "key", "value")
        value = self.loader.get_config_value("new_section", "key")
        self.assertEqual(value, "value")
    
    def test_update_config(self):
        """测试批量更新配置"""
        self.loader.load_config()
        
        updates = {
            "app_config": {
                "window_title": "新标题",
                "theme": "dark"
            },
            "matching_config": {
                "fuzzy_threshold": 0.9
            }
        }
        
        self.loader.update_config(updates)
        
        # 验证更新
        self.assertEqual(self.loader.get_config_value("app_config", "window_title"), "新标题")
        self.assertEqual(self.loader.get_config_value("app_config", "theme"), "dark")
        self.assertEqual(self.loader.get_config_value("matching_config", "fuzzy_threshold"), 0.9)
    
    def test_update_config_validation_error(self):
        """测试更新配置时的验证错误"""
        self.loader.load_config()
        
        invalid_updates = {
            "matching_config": {
                "fuzzy_threshold": 2.0  # 超出范围
            }
        }
        
        with self.assertRaises(ConfigValidationError):
            self.loader.update_config(invalid_updates)
        
        # 验证原配置未被修改
        original_threshold = self.loader.get_config_value("matching_config", "fuzzy_threshold")
        self.assertEqual(original_threshold, 0.8)
    
    def test_save_config(self):
        """测试保存配置"""
        self.loader.load_config()
        self.loader.set_config_value("app_config", "test_key", "test_value")
        
        # 保存配置
        self.loader.save_config()
        
        # 创建新的加载器验证保存结果
        new_loader = ConfigLoader(str(self.config_file))
        new_config = new_loader.load_config()
        
        self.assertEqual(new_config["app_config"]["test_key"], "test_value")
    
    def test_save_config_io_error(self):
        """测试保存配置时的IO错误"""
        # 使用无效路径（包含非法字符）
        invalid_file = Path("invalid<>path/config.json")
        loader = ConfigLoader(str(invalid_file))
        loader._config = {"test": "value"}
        
        with self.assertRaises((IOError, OSError)):
            loader.save_config()
    
    def test_reset_to_defaults(self):
        """测试重置为默认配置"""
        self.loader.load_config()
        self.loader.set_config_value("app_config", "window_title", "修改的标题")
        
        # 验证修改生效
        modified_title = self.loader._config["app_config"]["window_title"]
        self.assertEqual(modified_title, "修改的标题")
        
        # 重置为默认值
        self.loader.reset_to_defaults()
        
        # 验证重置结果（直接访问内存中的配置）
        title = self.loader._config["app_config"]["window_title"]
        self.assertEqual(title, "Excel岗位分数查询工具")
    
    def test_get_config_info(self):
        """测试获取配置信息"""
        self.loader.load_config()
        info = self.loader.get_config_info()
        
        self.assertIn("config_file", info)
        self.assertIn("file_exists", info)
        self.assertIn("sections", info)
        self.assertIn("total_keys", info)
        
        self.assertTrue(info["file_exists"])
        self.assertEqual(len(info["sections"]), 3)  # app_config, matching_config, logging_config
        self.assertGreater(info["total_keys"], 0)
    
    def test_merge_with_defaults(self):
        """测试与默认配置合并"""
        loaded_config = {
            "app_config": {
                "window_title": "自定义标题"
            },
            "custom_section": {
                "custom_key": "custom_value"
            }
        }
        
        merged = self.loader._merge_with_defaults(loaded_config)
        
        # 验证自定义值被保留
        self.assertEqual(merged["app_config"]["window_title"], "自定义标题")
        self.assertEqual(merged["custom_section"]["custom_key"], "custom_value")
        
        # 验证默认值被添加
        self.assertIn("default_output_filename", merged["app_config"])
        self.assertIn("matching_config", merged)
        self.assertIn("logging_config", merged)


if __name__ == '__main__':
    unittest.main()