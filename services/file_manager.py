"""
文件管理器
负责处理文件保存、路径处理和文件名冲突
"""
import os
from pathlib import Path
from typing import Optional, Tuple
import shutil
from datetime import datetime


class FileManager:
    """文件管理器类"""
    
    def __init__(self):
        self.default_filename = "岗位最低分数汇总.xlsx"
    
    def get_save_path(self, user_specified_path: Optional[str] = None) -> str:
        """
        获取保存路径，处理用户指定路径或使用默认路径
        
        Args:
            user_specified_path: 用户指定的保存路径
            
        Returns:
            str: 最终的保存路径
        """
        if user_specified_path:
            return self._process_user_path(user_specified_path)
        else:
            # 使用当前目录作为默认保存位置
            return os.path.join(os.getcwd(), self.default_filename)
    
    def _process_user_path(self, user_path: str) -> str:
        """
        处理用户指定的路径
        
        Args:
            user_path: 用户指定的路径
            
        Returns:
            str: 处理后的完整路径
        """
        path = Path(user_path)
        
        # 如果用户只指定了目录，添加默认文件名
        if path.is_dir() or (not path.suffix and not path.exists()):
            path = path / self.default_filename
        
        # 确保文件扩展名为.xlsx
        if path.suffix.lower() != '.xlsx':
            path = path.with_suffix('.xlsx')
        
        return str(path.resolve())
    
    def handle_file_conflict(self, file_path: str, strategy: str = "auto_rename") -> str:
        """
        处理文件名冲突
        
        Args:
            file_path: 原始文件路径
            strategy: 冲突处理策略 ("auto_rename", "overwrite", "backup")
            
        Returns:
            str: 处理后的文件路径
        """
        if not os.path.exists(file_path):
            return file_path
        
        if strategy == "overwrite":
            return file_path
        elif strategy == "backup":
            return self._create_backup_and_return_original(file_path)
        else:  # auto_rename
            return self._generate_unique_filename(file_path)
    
    def _generate_unique_filename(self, file_path: str) -> str:
        """
        生成唯一的文件名
        
        Args:
            file_path: 原始文件路径
            
        Returns:
            str: 唯一的文件路径
        """
        path = Path(file_path)
        base_name = path.stem
        extension = path.suffix
        directory = path.parent
        
        counter = 1
        while True:
            new_name = f"{base_name}_{counter}{extension}"
            new_path = directory / new_name
            
            if not new_path.exists():
                return str(new_path)
            
            counter += 1
            
            # 防止无限循环
            if counter > 1000:
                timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
                new_name = f"{base_name}_{timestamp}{extension}"
                return str(directory / new_name)
    
    def _create_backup_and_return_original(self, file_path: str) -> str:
        """
        创建备份文件并返回原始路径
        
        Args:
            file_path: 原始文件路径
            
        Returns:
            str: 原始文件路径
        """
        if os.path.exists(file_path):
            path = Path(file_path)
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            backup_name = f"{path.stem}_backup_{timestamp}{path.suffix}"
            backup_path = path.parent / backup_name
            
            shutil.copy2(file_path, backup_path)
        
        return file_path
    
    def validate_save_path(self, file_path: str) -> Tuple[bool, str]:
        """
        验证保存路径的有效性
        
        Args:
            file_path: 要验证的文件路径
            
        Returns:
            Tuple[bool, str]: (是否有效, 错误信息)
        """
        try:
            path = Path(file_path)
            
            # 检查父目录是否存在
            if not path.parent.exists():
                try:
                    path.parent.mkdir(parents=True, exist_ok=True)
                except Exception as e:
                    return False, f"无法创建目录: {str(e)}"
            
            # 检查父目录是否可写
            if not os.access(path.parent, os.W_OK):
                return False, f"目录没有写入权限: {path.parent}"
            
            # 检查文件是否被占用（如果存在）
            if path.exists():
                try:
                    # 尝试以写入模式打开文件
                    with open(path, 'a'):
                        pass
                except PermissionError:
                    return False, f"文件被占用或没有写入权限: {path.name}"
            
            # 检查文件名是否有效
            if not self._is_valid_filename(path.name):
                return False, f"文件名包含无效字符: {path.name}"
            
            return True, ""
            
        except Exception as e:
            return False, f"路径验证失败: {str(e)}"
    
    def _is_valid_filename(self, filename: str) -> bool:
        """
        检查文件名是否有效
        
        Args:
            filename: 文件名
            
        Returns:
            bool: 是否有效
        """
        # Windows无效字符
        invalid_chars = '<>:"/\\|?*'
        
        if not filename or filename.strip() == "":
            return False
        
        # 检查无效字符
        for char in invalid_chars:
            if char in filename:
                return False
        
        # 检查保留名称（Windows）
        reserved_names = [
            'CON', 'PRN', 'AUX', 'NUL',
            'COM1', 'COM2', 'COM3', 'COM4', 'COM5', 'COM6', 'COM7', 'COM8', 'COM9',
            'LPT1', 'LPT2', 'LPT3', 'LPT4', 'LPT5', 'LPT6', 'LPT7', 'LPT8', 'LPT9'
        ]
        
        name_without_ext = Path(filename).stem.upper()
        if name_without_ext in reserved_names:
            return False
        
        return True
    
    def ensure_directory_exists(self, file_path: str) -> bool:
        """
        确保文件的父目录存在
        
        Args:
            file_path: 文件路径
            
        Returns:
            bool: 是否成功创建或目录已存在
        """
        try:
            path = Path(file_path)
            path.parent.mkdir(parents=True, exist_ok=True)
            return True
        except Exception:
            return False
    
    def get_file_info(self, file_path: str) -> dict:
        """
        获取文件信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            dict: 文件信息
        """
        path = Path(file_path)
        
        info = {
            "exists": path.exists(),
            "directory": str(path.parent),
            "filename": path.name,
            "extension": path.suffix,
            "absolute_path": str(path.resolve())
        }
        
        if path.exists():
            stat = path.stat()
            info.update({
                "size": stat.st_size,
                "modified_time": datetime.fromtimestamp(stat.st_mtime),
                "is_readonly": not os.access(path, os.W_OK)
            })
        
        return info