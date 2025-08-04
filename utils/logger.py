"""
日志配置模块
负责配置和管理应用程序日志
"""
import logging
import logging.handlers
import os
import sys
from datetime import datetime
from pathlib import Path
from typing import Dict, Any, Optional, Union
from .config_loader import ConfigLoader


class LoggerError(Exception):
    """日志系统相关异常"""
    pass


class ApplicationLogger:
    """应用程序日志管理器"""
    
    _instances = {}  # 单例模式存储不同名称的日志器
    
    def __init__(self, name: str = "excel_position_query", config: Optional[Dict[str, Any]] = None):
        """
        初始化日志管理器
        
        Args:
            name: 日志记录器名称
            config: 日志配置字典
        """
        self.name = name
        self.config = config or self._load_config()
        self.logger = None
        self._setup_logger()
    
    @classmethod
    def get_logger(cls, name: str = "excel_position_query", config: Optional[Dict[str, Any]] = None) -> 'ApplicationLogger':
        """
        获取日志管理器实例（单例模式）
        
        Args:
            name: 日志记录器名称
            config: 日志配置字典
            
        Returns:
            日志管理器实例
        """
        if name not in cls._instances:
            cls._instances[name] = cls(name, config)
        return cls._instances[name]
    
    def _load_config(self) -> Dict[str, Any]:
        """
        加载日志配置
        
        Returns:
            日志配置字典
        """
        try:
            config_loader = ConfigLoader()
            return config_loader.get_logging_config()
        except Exception as e:
            # 如果配置加载失败，使用默认配置
            print(f"警告: 无法加载日志配置，使用默认配置: {e}")
            return {
                "level": "INFO",
                "log_file": "app.log",
                "max_file_size": "10MB",
                "backup_count": 5,
                "format": "%(asctime)s - %(name)s - %(levelname)s - %(message)s"
            }
    
    def _setup_logger(self) -> None:
        """设置日志记录器"""
        try:
            # 创建日志记录器
            self.logger = logging.getLogger(self.name)
            
            # 设置日志级别
            level = self.config.get("level", "INFO").upper()
            if hasattr(logging, level):
                self.logger.setLevel(getattr(logging, level))
            else:
                self.logger.setLevel(logging.INFO)
                self.logger.warning(f"无效的日志级别 '{level}'，使用默认级别 INFO")
            
            # 清除现有处理器（避免重复）
            self.logger.handlers.clear()
            
            # 创建格式化器
            log_format = self.config.get("format", "%(asctime)s - %(name)s - %(levelname)s - %(message)s")
            formatter = logging.Formatter(
                log_format,
                datefmt='%Y-%m-%d %H:%M:%S'
            )
            
            # 添加控制台处理器
            self._add_console_handler(formatter)
            
            # 添加文件处理器
            self._add_file_handler(formatter)
            
            # 防止日志传播到根日志器
            self.logger.propagate = False
            
        except Exception as e:
            raise LoggerError(f"设置日志记录器失败: {e}")
    
    def _add_console_handler(self, formatter: logging.Formatter) -> None:
        """
        添加控制台处理器
        
        Args:
            formatter: 日志格式化器
        """
        console_handler = logging.StreamHandler(sys.stdout)
        console_handler.setLevel(logging.INFO)
        console_handler.setFormatter(formatter)
        
        # 为不同级别的日志设置不同颜色（如果支持）
        if hasattr(sys.stdout, 'isatty') and sys.stdout.isatty():
            console_handler.setFormatter(ColoredFormatter(formatter._fmt))
        
        self.logger.addHandler(console_handler)
    
    def _add_file_handler(self, formatter: logging.Formatter) -> None:
        """
        添加文件处理器
        
        Args:
            formatter: 日志格式化器
        """
        try:
            log_file = self.config.get("log_file", "app.log")
            log_path = Path(log_file)
            
            # 确保日志目录存在
            log_path.parent.mkdir(parents=True, exist_ok=True)
            
            # 解析文件大小和备份数量
            max_size = self._parse_size(self.config.get("max_file_size", "10MB"))
            backup_count = self.config.get("backup_count", 5)
            
            # 创建轮转文件处理器
            file_handler = logging.handlers.RotatingFileHandler(
                log_file,
                maxBytes=max_size,
                backupCount=backup_count,
                encoding='utf-8'
            )
            
            file_handler.setLevel(getattr(logging, self.config.get("level", "INFO").upper()))
            file_handler.setFormatter(formatter)
            
            self.logger.addHandler(file_handler)
            
        except Exception as e:
            self.logger.warning(f"无法设置文件日志处理器: {e}")
    
    def _parse_size(self, size_str: str) -> int:
        """
        解析文件大小字符串
        
        Args:
            size_str: 大小字符串，如 "10MB"
            
        Returns:
            字节数
            
        Raises:
            ValueError: 无效的大小格式
        """
        try:
            size_str = str(size_str).upper().strip()
            
            if size_str.endswith('KB'):
                return int(float(size_str[:-2]) * 1024)
            elif size_str.endswith('MB'):
                return int(float(size_str[:-2]) * 1024 * 1024)
            elif size_str.endswith('GB'):
                return int(float(size_str[:-2]) * 1024 * 1024 * 1024)
            else:
                # 默认为字节
                return int(size_str)
        except (ValueError, TypeError) as e:
            raise ValueError(f"无效的文件大小格式: {size_str}") from e
    
    def debug(self, message: str, *args, **kwargs) -> None:
        """记录调试信息"""
        if self.logger:
            self.logger.debug(message, *args, **kwargs)
    
    def info(self, message: str, *args, **kwargs) -> None:
        """记录一般信息"""
        if self.logger:
            self.logger.info(message, *args, **kwargs)
    
    def warning(self, message: str, *args, **kwargs) -> None:
        """记录警告信息"""
        if self.logger:
            self.logger.warning(message, *args, **kwargs)
    
    def error(self, message: str, *args, **kwargs) -> None:
        """记录错误信息"""
        if self.logger:
            self.logger.error(message, *args, **kwargs)
    
    def critical(self, message: str, *args, **kwargs) -> None:
        """记录严重错误信息"""
        if self.logger:
            self.logger.critical(message, *args, **kwargs)
    
    def exception(self, message: str, *args, **kwargs) -> None:
        """记录异常信息（包含堆栈跟踪）"""
        if self.logger:
            self.logger.exception(message, *args, **kwargs)
    
    def log_operation(self, operation: str, details: Optional[Dict[str, Any]] = None) -> None:
        """
        记录操作日志
        
        Args:
            operation: 操作名称
            details: 操作详情
        """
        message = f"操作: {operation}"
        if details:
            detail_str = ", ".join([f"{k}={v}" for k, v in details.items()])
            message += f" - {detail_str}"
        
        self.info(message)
    
    def log_error_with_context(self, error: Exception, context: Optional[Dict[str, Any]] = None) -> None:
        """
        记录带上下文的错误信息
        
        Args:
            error: 异常对象
            context: 错误上下文信息
        """
        message = f"错误: {type(error).__name__}: {str(error)}"
        if context:
            context_str = ", ".join([f"{k}={v}" for k, v in context.items()])
            message += f" - 上下文: {context_str}"
        
        self.error(message)
        if hasattr(error, '__traceback__') and error.__traceback__:
            self.exception("详细错误信息:")
    
    def log_performance(self, operation: str, duration: float, details: Optional[Dict[str, Any]] = None) -> None:
        """
        记录性能日志
        
        Args:
            operation: 操作名称
            duration: 执行时间（秒）
            details: 额外详情
        """
        message = f"性能: {operation} 耗时 {duration:.3f}秒"
        if details:
            detail_str = ", ".join([f"{k}={v}" for k, v in details.items()])
            message += f" - {detail_str}"
        
        self.info(message)
    
    def get_log_info(self) -> Dict[str, Any]:
        """
        获取日志系统信息
        
        Returns:
            日志系统信息字典
        """
        info = {
            "logger_name": self.name,
            "level": self.logger.level if self.logger else None,
            "handlers_count": len(self.logger.handlers) if self.logger else 0,
            "config": self.config.copy()
        }
        
        # 获取日志文件信息
        log_file = self.config.get("log_file")
        if log_file and Path(log_file).exists():
            log_path = Path(log_file)
            info["log_file_info"] = {
                "path": str(log_path.absolute()),
                "size": log_path.stat().st_size,
                "modified": datetime.fromtimestamp(log_path.stat().st_mtime).isoformat()
            }
        
        return info
    
    def reload_config(self, new_config: Optional[Dict[str, Any]] = None) -> None:
        """
        重新加载配置
        
        Args:
            new_config: 新的配置字典，如果为None则从配置文件重新加载
        """
        if new_config:
            self.config = new_config
        else:
            self.config = self._load_config()
        
        self._setup_logger()
        self.info("日志配置已重新加载")


class ColoredFormatter(logging.Formatter):
    """彩色日志格式化器（用于控制台输出）"""
    
    # ANSI颜色代码
    COLORS = {
        'DEBUG': '\033[36m',    # 青色
        'INFO': '\033[32m',     # 绿色
        'WARNING': '\033[33m',  # 黄色
        'ERROR': '\033[31m',    # 红色
        'CRITICAL': '\033[35m', # 紫色
        'RESET': '\033[0m'      # 重置
    }
    
    def format(self, record):
        # 添加颜色
        if record.levelname in self.COLORS:
            record.levelname = f"{self.COLORS[record.levelname]}{record.levelname}{self.COLORS['RESET']}"
        
        return super().format(record)


class LoggerSetup:
    """日志设置类（向后兼容）"""
    
    @staticmethod
    def setup_logger(name: str = "excel_position_query", config: Dict[str, Any] = None) -> logging.Logger:
        """
        设置日志记录器（向后兼容方法）
        
        Args:
            name: 日志记录器名称
            config: 日志配置字典
            
        Returns:
            配置好的日志记录器
        """
        app_logger = ApplicationLogger.get_logger(name, config)
        return app_logger.logger


# 创建默认日志记录器实例
default_logger = ApplicationLogger.get_logger()

# 导出常用接口
def get_logger(name: str = "excel_position_query") -> ApplicationLogger:
    """获取日志记录器"""
    return ApplicationLogger.get_logger(name)

def log_info(message: str) -> None:
    """记录信息日志"""
    default_logger.info(message)

def log_error(message: str) -> None:
    """记录错误日志"""
    default_logger.error(message)

def log_warning(message: str) -> None:
    """记录警告日志"""
    default_logger.warning(message)