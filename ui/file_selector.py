"""
文件选择器组件
"""
import tkinter as tk
from tkinter import filedialog, messagebox
import os
from typing import Optional

from utils.logger import default_logger


class FileSelector:
    """文件选择器类"""
    
    def __init__(self, parent_window: Optional[tk.Tk] = None):
        """
        初始化文件选择器
        
        Args:
            parent_window: 父窗口，用于模态对话框
        """
        self.parent_window = parent_window
        default_logger.info("文件选择器初始化完成")
    
    def select_excel_file(self, title: str = "选择Excel文件") -> Optional[str]:
        """
        选择Excel文件
        
        Args:
            title: 对话框标题
            
        Returns:
            选择的文件路径，如果取消则返回None
        """
        try:
            file_path = filedialog.askopenfilename(
                parent=self.parent_window,
                title=title,
                filetypes=[
                    ("Excel文件", "*.xlsx *.xls"),
                    ("Excel 2007-365", "*.xlsx"),
                    ("Excel 97-2003", "*.xls"),
                    ("所有文件", "*.*")
                ],
                initialdir=os.getcwd()
            )
            
            if file_path:
                default_logger.info(f"用户选择文件: {file_path}")
                if self.validate_file_path(file_path):
                    return file_path
                else:
                    return None
            else:
                default_logger.info("用户取消了文件选择")
                return None
                
        except Exception as e:
            error_msg = f"文件选择过程中发生错误: {str(e)}"
            default_logger.error(error_msg)
            messagebox.showerror("错误", error_msg)
            return None
    
    def validate_file_path(self, file_path: str) -> bool:
        """
        验证文件路径
        
        Args:
            file_path: 文件路径
            
        Returns:
            验证是否通过
        """
        if not file_path:
            self._show_validation_error("文件路径不能为空")
            return False
        
        if not os.path.exists(file_path):
            self._show_validation_error(f"文件不存在: {file_path}")
            return False
        
        if not os.path.isfile(file_path):
            self._show_validation_error(f"路径不是文件: {file_path}")
            return False
        
        # 检查文件扩展名
        _, ext = os.path.splitext(file_path.lower())
        if ext not in ['.xlsx', '.xls']:
            self._show_validation_error(f"不支持的文件格式: {ext}。请选择Excel文件(.xlsx或.xls)")
            return False
        
        # 检查文件是否可读
        try:
            with open(file_path, 'rb') as f:
                # 尝试读取前几个字节来验证文件是否可访问
                f.read(10)
        except PermissionError:
            self._show_validation_error(f"没有权限访问文件: {file_path}")
            return False
        except Exception as e:
            self._show_validation_error(f"文件访问错误: {str(e)}")
            return False
        
        # 检查文件大小（避免选择过大的文件）
        try:
            file_size = os.path.getsize(file_path)
            max_size = 100 * 1024 * 1024  # 100MB
            if file_size > max_size:
                self._show_validation_error(f"文件过大: {file_size / (1024*1024):.1f}MB。最大支持100MB")
                return False
        except Exception as e:
            default_logger.warning(f"无法获取文件大小: {str(e)}")
        
        default_logger.info(f"文件验证通过: {file_path}")
        return True
    
    def _show_validation_error(self, message: str):
        """
        显示验证错误信息
        
        Args:
            message: 错误信息
        """
        default_logger.error(f"文件验证失败: {message}")
        messagebox.showerror("文件验证失败", message)
    
    def get_file_info(self, file_path: str) -> dict:
        """
        获取文件信息
        
        Args:
            file_path: 文件路径
            
        Returns:
            文件信息字典
        """
        if not os.path.exists(file_path):
            return {}
        
        try:
            stat = os.stat(file_path)
            return {
                'name': os.path.basename(file_path),
                'path': file_path,
                'size': stat.st_size,
                'size_mb': round(stat.st_size / (1024 * 1024), 2),
                'modified': stat.st_mtime,
                'extension': os.path.splitext(file_path)[1].lower()
            }
        except Exception as e:
            default_logger.error(f"获取文件信息失败: {str(e)}")
            return {}
    
    def select_position_file(self) -> Optional[str]:
        """
        选择职位表文件
        
        Returns:
            选择的文件路径
        """
        return self.select_excel_file("选择职位表Excel文件")
    
    def select_interview_file(self) -> Optional[str]:
        """
        选择面试人员名单文件
        
        Returns:
            选择的文件路径
        """
        return self.select_excel_file("选择面试人员名单Excel文件")
    
    def select_output_file(self, default_filename: str = "岗位最低分数汇总.xlsx") -> Optional[str]:
        """
        选择输出文件路径
        
        Args:
            default_filename: 默认文件名
            
        Returns:
            选择的文件路径
        """
        try:
            file_path = filedialog.asksaveasfilename(
                parent=self.parent_window,
                title="保存汇总报告",
                defaultextension=".xlsx",
                filetypes=[
                    ("Excel文件", "*.xlsx"),
                    ("所有文件", "*.*")
                ],
                initialdir=os.getcwd(),
                initialfile=default_filename
            )
            
            if file_path:
                default_logger.info(f"用户选择输出文件: {file_path}")
                return file_path
            else:
                default_logger.info("用户取消了输出文件选择")
                return None
                
        except Exception as e:
            error_msg = f"选择输出文件时发生错误: {str(e)}"
            default_logger.error(error_msg)
            messagebox.showerror("错误", error_msg)
            return None