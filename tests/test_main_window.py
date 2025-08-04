"""
主窗口UI组件测试
"""
import unittest
import tkinter as tk
from unittest.mock import patch, MagicMock
import sys
import os

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from ui.main_window import MainWindow


class TestMainWindow(unittest.TestCase):
    """主窗口测试类"""
    
    def setUp(self):
        """测试前设置"""
        # 创建测试用的主窗口
        self.main_window = MainWindow()
        
    def tearDown(self):
        """测试后清理"""
        # 销毁窗口
        if hasattr(self, 'main_window'):
            self.main_window.destroy()
    
    def test_window_initialization(self):
        """测试窗口初始化"""
        # 检查窗口是否正确创建
        self.assertIsNotNone(self.main_window.root)
        self.assertEqual(self.main_window.root.title(), "Excel岗位分数查询工具")
        
        # 检查文件路径变量是否初始化
        self.assertIsInstance(self.main_window.position_file_path, tk.StringVar)
        self.assertIsInstance(self.main_window.interview_file_path, tk.StringVar)
        self.assertEqual(self.main_window.position_file_path.get(), "")
        self.assertEqual(self.main_window.interview_file_path.get(), "")
    
    def test_file_path_setting(self):
        """测试文件路径设置"""
        test_position_path = "/test/path/position.xlsx"
        test_interview_path = "/test/path/interview.xlsx"
        
        # 设置文件路径
        self.main_window.set_position_file_path(test_position_path)
        self.main_window.set_interview_file_path(test_interview_path)
        
        # 验证路径是否正确设置
        self.assertEqual(self.main_window.position_file_path.get(), test_position_path)
        self.assertEqual(self.main_window.interview_file_path.get(), test_interview_path)
    
    def test_start_button_state(self):
        """测试开始按钮状态控制"""
        # 初始状态应该是禁用的
        self.assertEqual(str(self.main_window.start_btn['state']), 'disabled')
        
        # 设置一个文件路径，按钮仍应禁用
        self.main_window.set_position_file_path("/test/position.xlsx")
        self.main_window._check_ready_to_process()
        self.assertEqual(str(self.main_window.start_btn['state']), 'disabled')
        
        # 设置两个文件路径，按钮应该启用
        self.main_window.set_interview_file_path("/test/interview.xlsx")
        self.main_window._check_ready_to_process()
        self.assertEqual(str(self.main_window.start_btn['state']), 'normal')
    
    def test_clear_selections(self):
        """测试清空选择功能"""
        # 先设置一些文件路径
        self.main_window.set_position_file_path("/test/position.xlsx")
        self.main_window.set_interview_file_path("/test/interview.xlsx")
        self.main_window._check_ready_to_process()
        
        # 清空选择
        self.main_window._clear_selections()
        
        # 验证路径已清空且按钮被禁用
        self.assertEqual(self.main_window.position_file_path.get(), "")
        self.assertEqual(self.main_window.interview_file_path.get(), "")
        self.assertEqual(str(self.main_window.start_btn['state']), 'disabled')
    
    def test_callback_assignment(self):
        """测试回调函数赋值"""
        mock_callback = MagicMock()
        
        # 设置回调函数
        self.main_window.on_position_file_select = mock_callback
        self.main_window.on_interview_file_select = mock_callback
        self.main_window.on_start_processing = mock_callback
        
        # 验证回调函数已正确设置
        self.assertEqual(self.main_window.on_position_file_select, mock_callback)
        self.assertEqual(self.main_window.on_interview_file_select, mock_callback)
        self.assertEqual(self.main_window.on_start_processing, mock_callback)
    
    @patch('tkinter.messagebox.showerror')
    def test_show_error(self, mock_showerror):
        """测试错误信息显示"""
        error_msg = "测试错误信息"
        self.main_window.show_error(error_msg)
        
        # 验证messagebox.showerror被调用
        mock_showerror.assert_called_once_with("错误", error_msg)
    
    @patch('tkinter.messagebox.showinfo')
    def test_show_info(self, mock_showinfo):
        """测试信息提示显示"""
        info_msg = "测试信息"
        self.main_window.show_info(info_msg)
        
        # 验证messagebox.showinfo被调用
        mock_showinfo.assert_called_once_with("信息", info_msg)
    
    def test_status_update(self):
        """测试状态更新功能"""
        test_message = "测试状态信息"
        
        # 更新状态
        self.main_window._update_status(test_message)
        
        # 验证状态文本是否包含消息
        status_content = self.main_window.status_text.get("1.0", tk.END)
        self.assertIn(test_message, status_content)
    
    def test_progress_display(self):
        """测试进度显示功能"""
        # 测试进度更新
        self.main_window.show_progress(50, "测试进度")
        
        # 验证进度条值
        self.assertEqual(self.main_window.progress_var.get(), 50)
        self.assertEqual(self.main_window.progress_label['text'], "50%")
        
        # 测试进度重置
        self.main_window.reset_progress()
        self.assertEqual(self.main_window.progress_var.get(), 0)
        self.assertEqual(self.main_window.progress_label['text'], "0%")
    
    def test_results_display(self):
        """测试结果显示功能"""
        test_results = {
            'total_positions': 10,
            'processed_positions': 8,
            'failed_positions': 2,
            'total_candidates': 50,
            'output_file': 'test_output.xlsx',
            'processing_time': 5.5
        }
        
        with patch('tkinter.messagebox.showinfo'):
            self.main_window.show_results(test_results)
        
        # 验证状态文本包含结果信息
        status_content = self.main_window.status_text.get("1.0", tk.END)
        self.assertIn("处理完成", status_content)
        self.assertIn("总岗位数: 10", status_content)
        self.assertIn("成功处理岗位数: 8", status_content)
    
    def test_processing_state(self):
        """测试处理状态设置"""
        # 测试设置为处理中状态
        self.main_window.set_processing_state(True)
        self.assertEqual(str(self.main_window.start_btn['state']), 'disabled')
        self.assertEqual(self.main_window.start_btn['text'], '处理中...')
        
        # 测试设置为非处理状态
        self.main_window.set_processing_state(False)
        self.assertEqual(str(self.main_window.start_btn['state']), 'normal')
        self.assertEqual(self.main_window.start_btn['text'], '开始处理')
    
    def test_clear_status(self):
        """测试清空状态功能"""
        # 先添加一些状态信息
        self.main_window._update_status("测试信息")
        self.main_window.show_progress(50)
        
        # 清空状态
        self.main_window.clear_status()
        
        # 验证状态已清空
        status_content = self.main_window.status_text.get("1.0", tk.END).strip()
        self.assertEqual(status_content, "")
        self.assertEqual(self.main_window.progress_var.get(), 0)


if __name__ == '__main__':
    # 运行测试
    unittest.main()