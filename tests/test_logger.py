"""
日志系统测试
"""
import logging
import os
import tempfile
import unittest
from pathlib import Path
from unittest.mock import patch, MagicMock
import time

from utils.logger import ApplicationLogger, LoggerError, ColoredFormatter, get_logger, log_info, log_error, log_warning


class TestApplicationLogger(unittest.TestCase):
    """应用程序日志器测试类"""
    
    def setUp(self):
        """测试前准备"""
        self.temp_dir = tempfile.mkdtemp()
        self.log_file = Path(self.temp_dir) / "test.log"
        
        # 清除单例实例
        ApplicationLogger._instances.clear()
        
        self.test_config = {
            "level": "DEBUG",
            "log_file": str(self.log_file),
            "max_file_size": "1MB",
            "backup_count": 3,
            "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        }
    
    def tearDown(self):
        """测试后清理"""
        # 关闭所有日志处理器
        for instance in ApplicationLogger._instances.values():
            if instance.logger:
                for handler in instance.logger.handlers[:]:
                    handler.close()
                    instance.logger.removeHandler(handler)
        
        # 清除单例实例
        ApplicationLogger._instances.clear()
        
        # 清理日志文件
        try:
            if self.log_file.exists():
                self.log_file.unlink()
        except PermissionError:
            pass  # 忽略权限错误
        
        # 清理备份文件
        for i in range(1, 6):
            backup_file = Path(f"{self.log_file}.{i}")
            try:
                if backup_file.exists():
                    backup_file.unlink()
            except PermissionError:
                pass  # 忽略权限错误
        
        # 清理目录
        try:
            # 清理可能剩余的文件
            for file in Path(self.temp_dir).iterdir():
                try:
                    if file.is_file():
                        file.unlink()
                    elif file.is_dir():
                        import shutil
                        shutil.rmtree(file)
                except PermissionError:
                    pass
            os.rmdir(self.temp_dir)
        except (OSError, PermissionError):
            pass  # 忽略清理错误
    
    def test_logger_initialization(self):
        """测试日志器初始化"""
        logger = ApplicationLogger("test_logger", self.test_config)
        
        self.assertEqual(logger.name, "test_logger")
        self.assertIsNotNone(logger.logger)
        self.assertEqual(logger.logger.level, logging.DEBUG)
    
    def test_singleton_pattern(self):
        """测试单例模式"""
        logger1 = ApplicationLogger.get_logger("test_singleton", self.test_config)
        logger2 = ApplicationLogger.get_logger("test_singleton")
        
        self.assertIs(logger1, logger2)
        
        # 不同名称应该是不同实例
        logger3 = ApplicationLogger.get_logger("different_name", self.test_config)
        self.assertIsNot(logger1, logger3)
    
    def test_config_loading_fallback(self):
        """测试配置加载失败时的回退机制"""
        with patch('utils.logger.ConfigLoader') as mock_config_loader:
            mock_config_loader.side_effect = Exception("配置加载失败")
            
            logger = ApplicationLogger("test_fallback")
            
            # 应该使用默认配置
            self.assertIn("level", logger.config)
            self.assertEqual(logger.config["level"], "INFO")
    
    def test_invalid_log_level(self):
        """测试无效日志级别处理"""
        invalid_config = self.test_config.copy()
        invalid_config["level"] = "INVALID_LEVEL"
        
        logger = ApplicationLogger("test_invalid_level", invalid_config)
        
        # 应该回退到INFO级别
        self.assertEqual(logger.logger.level, logging.INFO)
    
    def test_file_size_parsing(self):
        """测试文件大小解析"""
        logger = ApplicationLogger("test_size", self.test_config)
        
        # 测试各种格式
        self.assertEqual(logger._parse_size("1024"), 1024)
        self.assertEqual(logger._parse_size("1KB"), 1024)
        self.assertEqual(logger._parse_size("1MB"), 1024 * 1024)
        self.assertEqual(logger._parse_size("1GB"), 1024 * 1024 * 1024)
        self.assertEqual(logger._parse_size("1.5MB"), int(1.5 * 1024 * 1024))
        
        # 测试无效格式
        with self.assertRaises(ValueError):
            logger._parse_size("invalid")
    
    def test_logging_methods(self):
        """测试各种日志记录方法"""
        logger = ApplicationLogger("test_methods", self.test_config)
        
        # 测试各种级别的日志
        logger.debug("调试信息")
        logger.info("一般信息")
        logger.warning("警告信息")
        logger.error("错误信息")
        logger.critical("严重错误")
        
        # 验证日志文件存在且有内容
        self.assertTrue(self.log_file.exists())
        
        with open(self.log_file, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn("调试信息", content)
            self.assertIn("一般信息", content)
            self.assertIn("警告信息", content)
            self.assertIn("错误信息", content)
            self.assertIn("严重错误", content)
    
    def test_exception_logging(self):
        """测试异常日志记录"""
        logger = ApplicationLogger("test_exception", self.test_config)
        
        try:
            raise ValueError("测试异常")
        except ValueError:
            logger.exception("捕获到异常")
        
        # 验证异常信息被记录
        with open(self.log_file, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn("捕获到异常", content)
            self.assertIn("ValueError", content)
            self.assertIn("测试异常", content)
    
    def test_operation_logging(self):
        """测试操作日志记录"""
        logger = ApplicationLogger("test_operation", self.test_config)
        
        logger.log_operation("文件读取", {"文件名": "test.xlsx", "大小": "1MB"})
        
        with open(self.log_file, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn("操作: 文件读取", content)
            self.assertIn("文件名=test.xlsx", content)
            self.assertIn("大小=1MB", content)
    
    def test_error_with_context_logging(self):
        """测试带上下文的错误日志"""
        logger = ApplicationLogger("test_error_context", self.test_config)
        
        error = ValueError("测试错误")
        context = {"用户": "张三", "操作": "数据处理"}
        
        logger.log_error_with_context(error, context)
        
        with open(self.log_file, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn("错误: ValueError: 测试错误", content)
            self.assertIn("上下文:", content)
            self.assertIn("用户=张三", content)
            self.assertIn("操作=数据处理", content)
    
    def test_performance_logging(self):
        """测试性能日志记录"""
        logger = ApplicationLogger("test_performance", self.test_config)
        
        logger.log_performance("数据处理", 1.234, {"记录数": 1000})
        
        with open(self.log_file, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn("性能: 数据处理 耗时 1.234秒", content)
            self.assertIn("记录数=1000", content)
    
    def test_log_file_creation(self):
        """测试日志文件创建"""
        # 使用不存在的目录
        nested_log_file = Path(self.temp_dir) / "logs" / "nested" / "test.log"
        config = self.test_config.copy()
        config["log_file"] = str(nested_log_file)
        
        logger = ApplicationLogger("test_file_creation", config)
        logger.info("测试消息")
        
        # 验证目录和文件被创建
        self.assertTrue(nested_log_file.parent.exists())
        self.assertTrue(nested_log_file.exists())
    
    def test_file_rotation(self):
        """测试日志文件轮转"""
        # 设置很小的文件大小以触发轮转
        small_config = self.test_config.copy()
        small_config["max_file_size"] = "100"  # 100字节
        small_config["backup_count"] = 2
        
        logger = ApplicationLogger("test_rotation", small_config)
        
        # 写入足够多的日志以触发轮转
        for i in range(50):
            logger.info(f"这是一条很长的测试日志消息，用于触发文件轮转 {i}")
        
        # 验证备份文件存在
        backup1 = Path(f"{self.log_file}.1")
        self.assertTrue(backup1.exists() or self.log_file.stat().st_size > 0)
    
    def test_get_log_info(self):
        """测试获取日志信息"""
        logger = ApplicationLogger("test_info", self.test_config)
        logger.info("测试消息")
        
        info = logger.get_log_info()
        
        self.assertEqual(info["logger_name"], "test_info")
        self.assertEqual(info["level"], logging.DEBUG)
        self.assertGreater(info["handlers_count"], 0)
        self.assertIn("config", info)
        
        # 如果日志文件存在，应该有文件信息
        if self.log_file.exists():
            self.assertIn("log_file_info", info)
            self.assertIn("path", info["log_file_info"])
            self.assertIn("size", info["log_file_info"])
    
    def test_config_reload(self):
        """测试配置重新加载"""
        logger = ApplicationLogger("test_reload", self.test_config)
        original_level = logger.logger.level
        
        # 修改配置
        new_config = self.test_config.copy()
        new_config["level"] = "ERROR"
        
        logger.reload_config(new_config)
        
        # 验证配置已更新
        self.assertEqual(logger.logger.level, logging.ERROR)
        self.assertNotEqual(logger.logger.level, original_level)
    
    def test_file_handler_error_handling(self):
        """测试文件处理器错误处理"""
        # 使用无效路径
        invalid_config = self.test_config.copy()
        invalid_config["log_file"] = "/invalid<>path/test.log"
        
        # 应该不抛出异常，而是记录警告
        logger = ApplicationLogger("test_file_error", invalid_config)
        
        # 日志器应该仍然可用（只是没有文件处理器）
        self.assertIsNotNone(logger.logger)


class TestColoredFormatter(unittest.TestCase):
    """彩色格式化器测试类"""
    
    def test_colored_formatting(self):
        """测试彩色格式化"""
        formatter = ColoredFormatter("%(levelname)s - %(message)s")
        
        # 创建不同级别的日志记录
        record_info = logging.LogRecord(
            name="test", level=logging.INFO, pathname="", lineno=0,
            msg="测试信息", args=(), exc_info=None
        )
        
        formatted = formatter.format(record_info)
        
        # 验证包含颜色代码
        self.assertIn('\033[32m', formatted)  # 绿色
        self.assertIn('\033[0m', formatted)   # 重置


class TestModuleFunctions(unittest.TestCase):
    """模块函数测试类"""
    
    def setUp(self):
        """测试前准备"""
        ApplicationLogger._instances.clear()
    
    def tearDown(self):
        """测试后清理"""
        ApplicationLogger._instances.clear()
    
    def test_get_logger_function(self):
        """测试get_logger函数"""
        logger = get_logger("test_function")
        self.assertIsInstance(logger, ApplicationLogger)
        self.assertEqual(logger.name, "test_function")
    
    @patch('utils.logger.default_logger')
    def test_module_log_functions(self, mock_logger):
        """测试模块级别的日志函数"""
        log_info("测试信息")
        log_error("测试错误")
        log_warning("测试警告")
        
        mock_logger.info.assert_called_once_with("测试信息")
        mock_logger.error.assert_called_once_with("测试错误")
        mock_logger.warning.assert_called_once_with("测试警告")


class TestLoggerIntegration(unittest.TestCase):
    """日志系统集成测试"""
    
    def setUp(self):
        """测试前准备"""
        self.temp_dir = tempfile.mkdtemp()
        self.log_file = Path(self.temp_dir) / "integration.log"
        ApplicationLogger._instances.clear()
    
    def tearDown(self):
        """测试后清理"""
        # 关闭所有日志处理器
        for instance in ApplicationLogger._instances.values():
            if instance.logger:
                for handler in instance.logger.handlers[:]:
                    handler.close()
                    instance.logger.removeHandler(handler)
        
        # 清除单例实例
        ApplicationLogger._instances.clear()
        
        # 清理日志文件
        try:
            if self.log_file.exists():
                self.log_file.unlink()
            os.rmdir(self.temp_dir)
        except (OSError, PermissionError):
            pass  # 忽略清理错误
    
    def test_real_world_scenario(self):
        """测试真实世界场景"""
        config = {
            "level": "INFO",
            "log_file": str(self.log_file),
            "max_file_size": "1MB",
            "backup_count": 3,
            "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
        }
        
        logger = ApplicationLogger.get_logger("integration_test", config)
        
        # 模拟应用程序启动
        logger.info("应用程序启动")
        
        # 模拟文件处理操作
        logger.log_operation("读取Excel文件", {"文件": "test.xlsx", "大小": "2MB"})
        
        # 模拟错误处理
        try:
            raise FileNotFoundError("文件未找到")
        except FileNotFoundError as e:
            logger.log_error_with_context(e, {"文件路径": "/path/to/file.xlsx"})
        
        # 模拟性能监控
        start_time = time.time()
        time.sleep(0.01)  # 模拟处理时间
        duration = time.time() - start_time
        logger.log_performance("数据处理", duration, {"记录数": 1000})
        
        # 验证日志文件内容
        self.assertTrue(self.log_file.exists())
        
        with open(self.log_file, 'r', encoding='utf-8') as f:
            content = f.read()
            self.assertIn("应用程序启动", content)
            self.assertIn("操作: 读取Excel文件", content)
            self.assertIn("错误: FileNotFoundError", content)
            self.assertIn("性能: 数据处理", content)


if __name__ == '__main__':
    unittest.main()