"""
文件选择器测试
"""
import unittest
import tempfile
import os
from unittest.mock import patch, MagicMock
import sys

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.file_selector import FileSelector


class TestFileSelector(unittest.TestCase):
    """文件选择器测试类"""
    
    def setUp(self):
        """测试前设置"""
        self.file_selector = FileSelector()
        
        # 创建临时测试文件
        self.temp_dir = tempfile.mkdtemp()
        self.valid_excel_file = os.path.join(self.temp_dir, "test.xlsx")
        self.invalid_file = os.path.join(self.temp_dir, "test.txt")
        
        # 创建测试文件
        with open(self.valid_excel_file, 'wb') as f:
            # 写入一些Excel文件的基本头部字节
            f.write(b'PK\x03\x04')  # ZIP文件头（Excel是ZIP格式）
        
        with open(self.invalid_file, 'w') as f:
            f.write("test content")
    
    def tearDown(self):
        """测试后清理"""
        # 清理临时文件
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_initialization(self):
        """测试初始化"""
        file_selector = FileSelector()
        self.assertIsNotNone(file_selector)
        self.assertIsNone(file_selector.parent_window)
    
    def test_validate_file_path_valid_excel(self):
        """测试验证有效的Excel文件"""
        result = self.file_selector.validate_file_path(self.valid_excel_file)
        self.assertTrue(result)
    
    def test_validate_file_path_empty(self):
        """测试验证空路径"""
        with patch('tkinter.messagebox.showerror'):
            result = self.file_selector.validate_file_path("")
            self.assertFalse(result)
    
    def test_validate_file_path_nonexistent(self):
        """测试验证不存在的文件"""
        with patch('tkinter.messagebox.showerror'):
            result = self.file_selector.validate_file_path("/nonexistent/file.xlsx")
            self.assertFalse(result)
    
    def test_validate_file_path_invalid_extension(self):
        """测试验证无效扩展名"""
        with patch('tkinter.messagebox.showerror'):
            result = self.file_selector.validate_file_path(self.invalid_file)
            self.assertFalse(result)
    
    def test_get_file_info_valid_file(self):
        """测试获取文件信息"""
        info = self.file_selector.get_file_info(self.valid_excel_file)
        
        self.assertIsInstance(info, dict)
        self.assertEqual(info['name'], 'test.xlsx')
        self.assertEqual(info['path'], self.valid_excel_file)
        self.assertIn('size', info)
        self.assertIn('size_mb', info)
        self.assertEqual(info['extension'], '.xlsx')
    
    def test_get_file_info_nonexistent_file(self):
        """测试获取不存在文件的信息"""
        info = self.file_selector.get_file_info("/nonexistent/file.xlsx")
        self.assertEqual(info, {})
    
    @patch('tkinter.filedialog.askopenfilename')
    def test_select_excel_file_success(self, mock_dialog):
        """测试成功选择Excel文件"""
        mock_dialog.return_value = self.valid_excel_file
        
        result = self.file_selector.select_excel_file("测试标题")
        
        self.assertEqual(result, self.valid_excel_file)
        mock_dialog.assert_called_once()
    
    @patch('tkinter.filedialog.askopenfilename')
    def test_select_excel_file_cancel(self, mock_dialog):
        """测试取消选择文件"""
        mock_dialog.return_value = ""
        
        result = self.file_selector.select_excel_file("测试标题")
        
        self.assertIsNone(result)
    
    @patch('tkinter.filedialog.askopenfilename')
    def test_select_position_file(self, mock_dialog):
        """测试选择职位表文件"""
        mock_dialog.return_value = self.valid_excel_file
        
        result = self.file_selector.select_position_file()
        
        self.assertEqual(result, self.valid_excel_file)
        # 验证调用时使用了正确的标题
        args, kwargs = mock_dialog.call_args
        self.assertEqual(kwargs['title'], "选择职位表Excel文件")
    
    @patch('tkinter.filedialog.askopenfilename')
    def test_select_interview_file(self, mock_dialog):
        """测试选择面试人员名单文件"""
        mock_dialog.return_value = self.valid_excel_file
        
        result = self.file_selector.select_interview_file()
        
        self.assertEqual(result, self.valid_excel_file)
        # 验证调用时使用了正确的标题
        args, kwargs = mock_dialog.call_args
        self.assertEqual(kwargs['title'], "选择面试人员名单Excel文件")
    
    @patch('tkinter.filedialog.asksaveasfilename')
    def test_select_output_file(self, mock_dialog):
        """测试选择输出文件"""
        output_path = os.path.join(self.temp_dir, "output.xlsx")
        mock_dialog.return_value = output_path
        
        result = self.file_selector.select_output_file()
        
        self.assertEqual(result, output_path)
        # 验证调用时使用了正确的参数
        args, kwargs = mock_dialog.call_args
        self.assertEqual(kwargs['title'], "保存汇总报告")
        self.assertEqual(kwargs['initialfile'], "岗位最低分数汇总.xlsx")
    
    def test_validate_large_file(self):
        """测试验证大文件"""
        # 创建一个大文件用于测试
        large_file = os.path.join(self.temp_dir, "large.xlsx")
        with open(large_file, 'wb') as f:
            # 写入超过100MB的数据
            f.write(b'0' * (101 * 1024 * 1024))
        
        with patch('tkinter.messagebox.showerror'):
            result = self.file_selector.validate_file_path(large_file)
            self.assertFalse(result)


if __name__ == '__main__':
    unittest.main()