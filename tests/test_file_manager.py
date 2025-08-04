"""
文件管理器测试
"""
import unittest
import tempfile
import os
import shutil
from pathlib import Path
from datetime import datetime

from services.file_manager import FileManager


class TestFileManager(unittest.TestCase):
    """文件管理器测试类"""
    
    def setUp(self):
        """测试前准备"""
        self.file_manager = FileManager()
        self.temp_dir = tempfile.mkdtemp()
        self.test_file = os.path.join(self.temp_dir, "test.xlsx")
    
    def tearDown(self):
        """测试后清理"""
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_get_save_path_default(self):
        """测试获取默认保存路径"""
        path = self.file_manager.get_save_path()
        
        self.assertTrue(path.endswith("岗位最低分数汇总.xlsx"))
        self.assertTrue(os.path.isabs(path))
    
    def test_get_save_path_user_specified_file(self):
        """测试用户指定文件路径"""
        user_path = os.path.join(self.temp_dir, "custom_report.xlsx")
        path = self.file_manager.get_save_path(user_path)
        
        self.assertEqual(path, os.path.abspath(user_path))
    
    def test_get_save_path_user_specified_directory(self):
        """测试用户指定目录路径"""
        path = self.file_manager.get_save_path(self.temp_dir)
        
        expected_path = os.path.join(self.temp_dir, "岗位最低分数汇总.xlsx")
        self.assertEqual(path, os.path.abspath(expected_path))
    
    def test_get_save_path_add_extension(self):
        """测试自动添加文件扩展名"""
        user_path = os.path.join(self.temp_dir, "report")
        path = self.file_manager.get_save_path(user_path)
        
        self.assertTrue(path.endswith(".xlsx"))
    
    def test_handle_file_conflict_no_conflict(self):
        """测试无冲突情况"""
        non_existing_file = os.path.join(self.temp_dir, "new_file.xlsx")
        result_path = self.file_manager.handle_file_conflict(non_existing_file)
        
        self.assertEqual(result_path, non_existing_file)
    
    def test_handle_file_conflict_auto_rename(self):
        """测试自动重命名策略"""
        # 创建一个已存在的文件
        with open(self.test_file, 'w') as f:
            f.write("test")
        
        result_path = self.file_manager.handle_file_conflict(self.test_file, "auto_rename")
        
        self.assertNotEqual(result_path, self.test_file)
        self.assertTrue(result_path.endswith("_1.xlsx"))
    
    def test_handle_file_conflict_overwrite(self):
        """测试覆盖策略"""
        # 创建一个已存在的文件
        with open(self.test_file, 'w') as f:
            f.write("test")
        
        result_path = self.file_manager.handle_file_conflict(self.test_file, "overwrite")
        
        self.assertEqual(result_path, self.test_file)
    
    def test_handle_file_conflict_backup(self):
        """测试备份策略"""
        # 创建一个已存在的文件
        with open(self.test_file, 'w') as f:
            f.write("test")
        
        result_path = self.file_manager.handle_file_conflict(self.test_file, "backup")
        
        self.assertEqual(result_path, self.test_file)
        
        # 检查是否创建了备份文件
        backup_files = [f for f in os.listdir(self.temp_dir) if "backup" in f]
        self.assertTrue(len(backup_files) > 0)
    
    def test_generate_unique_filename_multiple_conflicts(self):
        """测试多个文件冲突的情况"""
        base_file = os.path.join(self.temp_dir, "report.xlsx")
        
        # 创建多个冲突文件
        for i in range(3):
            if i == 0:
                conflict_file = base_file
            else:
                conflict_file = os.path.join(self.temp_dir, f"report_{i}.xlsx")
            
            with open(conflict_file, 'w') as f:
                f.write(f"test{i}")
        
        result_path = self.file_manager._generate_unique_filename(base_file)
        
        self.assertTrue(result_path.endswith("_3.xlsx"))
        self.assertFalse(os.path.exists(result_path))
    
    def test_validate_save_path_valid(self):
        """测试有效路径验证"""
        valid_path = os.path.join(self.temp_dir, "valid_report.xlsx")
        
        is_valid, error_msg = self.file_manager.validate_save_path(valid_path)
        
        self.assertTrue(is_valid)
        self.assertEqual(error_msg, "")
    
    def test_validate_save_path_create_directory(self):
        """测试创建不存在的目录"""
        new_dir = os.path.join(self.temp_dir, "new_subdir")
        valid_path = os.path.join(new_dir, "report.xlsx")
        
        is_valid, error_msg = self.file_manager.validate_save_path(valid_path)
        
        self.assertTrue(is_valid)
        self.assertTrue(os.path.exists(new_dir))
    
    def test_validate_save_path_invalid_filename(self):
        """测试无效文件名"""
        invalid_path = os.path.join(self.temp_dir, "report<>.xlsx")
        
        is_valid, error_msg = self.file_manager.validate_save_path(invalid_path)
        
        self.assertFalse(is_valid)
        self.assertIn("无效字符", error_msg)
    
    def test_is_valid_filename_valid(self):
        """测试有效文件名"""
        valid_names = ["report.xlsx", "岗位分数_2024.xlsx", "test-file.xlsx"]
        
        for name in valid_names:
            with self.subTest(filename=name):
                self.assertTrue(self.file_manager._is_valid_filename(name))
    
    def test_is_valid_filename_invalid(self):
        """测试无效文件名"""
        invalid_names = ["", "report<>.xlsx", "CON.xlsx", "file|name.xlsx"]
        
        for name in invalid_names:
            with self.subTest(filename=name):
                self.assertFalse(self.file_manager._is_valid_filename(name))
    
    def test_ensure_directory_exists(self):
        """测试确保目录存在"""
        new_dir = os.path.join(self.temp_dir, "nested", "directory")
        file_path = os.path.join(new_dir, "file.xlsx")
        
        result = self.file_manager.ensure_directory_exists(file_path)
        
        self.assertTrue(result)
        self.assertTrue(os.path.exists(new_dir))
    
    def test_get_file_info_existing_file(self):
        """测试获取已存在文件的信息"""
        # 创建测试文件
        with open(self.test_file, 'w') as f:
            f.write("test content")
        
        info = self.file_manager.get_file_info(self.test_file)
        
        self.assertTrue(info["exists"])
        self.assertEqual(info["filename"], "test.xlsx")
        self.assertEqual(info["extension"], ".xlsx")
        self.assertIsInstance(info["size"], int)
        self.assertIsInstance(info["modified_time"], datetime)
    
    def test_get_file_info_non_existing_file(self):
        """测试获取不存在文件的信息"""
        non_existing_file = os.path.join(self.temp_dir, "non_existing.xlsx")
        
        info = self.file_manager.get_file_info(non_existing_file)
        
        self.assertFalse(info["exists"])
        self.assertEqual(info["filename"], "non_existing.xlsx")
        self.assertEqual(info["extension"], ".xlsx")
        self.assertNotIn("size", info)
        self.assertNotIn("modified_time", info)


if __name__ == '__main__':
    unittest.main()