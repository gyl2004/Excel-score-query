"""
ProcessingEngine集成测试
测试完整的业务处理流程
"""
import unittest
import tempfile
import os
import pandas as pd
from unittest.mock import Mock, patch, MagicMock
import sys
import logging

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.processing_engine import ProcessingEngine, ProcessingEngineError, ProcessingProgress, ProcessingResult
from models.data_models import PositionScoreResult


class TestProcessingEngine(unittest.TestCase):
    """ProcessingEngine测试类"""
    
    def setUp(self):
        """测试前准备"""
        # 创建临时目录
        self.temp_dir = tempfile.mkdtemp()
        
        # 创建测试用的Excel文件
        self.position_file = os.path.join(self.temp_dir, "test_positions.xlsx")
        self.interview_file = os.path.join(self.temp_dir, "test_interviews.xlsx")
        self.output_file = os.path.join(self.temp_dir, "test_output.xlsx")
        
        # 创建测试数据
        self._create_test_excel_files()
        
        # 创建进度回调mock
        self.progress_callback = Mock()
        
        # 创建ProcessingEngine实例
        self.engine = ProcessingEngine(progress_callback=self.progress_callback)
        
        # 配置日志
        logging.basicConfig(level=logging.INFO)
    
    def tearDown(self):
        """测试后清理"""
        # 清理临时文件
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def _create_test_excel_files(self):
        """创建测试用的Excel文件"""
        # 创建职位表测试数据
        position_data = {
            '岗位代码': ['P001', 'P002', 'P003', 'P004'],
            '岗位名称': ['软件工程师', '产品经理', '数据分析师', '测试工程师'],
            '部门': ['技术部', '产品部', '数据部', '质量部']
        }
        position_df = pd.DataFrame(position_data)
        position_df.to_excel(self.position_file, index=False)
        
        # 创建面试人员名单测试数据
        interview_data = {
            '姓名': ['张三', '李四', '王五', '赵六', '钱七', '孙八'],
            '岗位名称': ['软件工程师', '软件工程师', '产品经理', '数据分析师', '数据分析师', '未知岗位'],
            '分数': [85.5, 78.0, 92.5, 88.0, 76.5, 90.0]
        }
        interview_df = pd.DataFrame(interview_data)
        interview_df.to_excel(self.interview_file, index=False)
    
    def test_process_files_success(self):
        """测试成功的文件处理流程"""
        # 执行处理
        result = self.engine.process_files(
            position_file=self.position_file,
            interview_file=self.interview_file,
            output_path=self.output_file
        )
        
        # 验证结果
        self.assertTrue(result.success)
        self.assertIsInstance(result.results, list)
        self.assertGreater(len(result.results), 0)
        self.assertIsNotNone(result.statistics)
        self.assertIsNotNone(result.report_path)
        self.assertIsNotNone(result.processing_time)
        
        # 验证进度回调被调用
        self.assertTrue(self.progress_callback.called)
        self.assertEqual(self.progress_callback.call_count, 7)  # 7个处理步骤
        
        # 验证进度回调参数
        for call in self.progress_callback.call_args_list:
            progress = call[0][0]
            self.assertIsInstance(progress, ProcessingProgress)
            self.assertGreaterEqual(progress.step_number, 1)
            self.assertLessEqual(progress.step_number, 7)
            self.assertEqual(progress.total_steps, 7)
        
        # 验证结果数据
        for position_result in result.results:
            self.assertIsInstance(position_result, PositionScoreResult)
            self.assertIsNotNone(position_result.position_name)
            self.assertIn(position_result.status, ["正常", "无面试人员", "数据异常", "无法匹配"])
        
        # 验证报告文件被创建
        self.assertTrue(os.path.exists(result.report_path))
    
    def test_process_files_with_invalid_position_file(self):
        """测试无效职位表文件的处理"""
        invalid_file = os.path.join(self.temp_dir, "invalid.txt")
        with open(invalid_file, 'w') as f:
            f.write("invalid content")
        
        result = self.engine.process_files(
            position_file=invalid_file,
            interview_file=self.interview_file,
            output_path=self.output_file
        )
        
        # 验证结果
        self.assertFalse(result.success)
        self.assertIn("验证失败", result.message)
        self.assertEqual(len(result.results), 0)
        self.assertIsNotNone(result.errors)
        self.assertGreater(len(result.errors), 0)
    
    def test_process_files_with_invalid_interview_file(self):
        """测试无效面试名单文件的处理"""
        invalid_file = os.path.join(self.temp_dir, "invalid.txt")
        with open(invalid_file, 'w') as f:
            f.write("invalid content")
        
        result = self.engine.process_files(
            position_file=self.position_file,
            interview_file=invalid_file,
            output_path=self.output_file
        )
        
        # 验证结果
        self.assertFalse(result.success)
        self.assertIn("验证失败", result.message)
        self.assertEqual(len(result.results), 0)
    
    def test_process_files_with_nonexistent_files(self):
        """测试不存在的文件处理"""
        nonexistent_file = os.path.join(self.temp_dir, "nonexistent.xlsx")
        
        result = self.engine.process_files(
            position_file=nonexistent_file,
            interview_file=self.interview_file,
            output_path=self.output_file
        )
        
        # 验证结果
        self.assertFalse(result.success)
        self.assertIn("验证失败", result.message)
    
    def test_validate_input_files_success(self):
        """测试成功的输入文件验证"""
        result = self.engine.validate_input_files(self.position_file, self.interview_file)
        self.assertTrue(result)
    
    def test_validate_input_files_invalid_position_file(self):
        """测试无效职位表文件验证"""
        invalid_file = os.path.join(self.temp_dir, "invalid.txt")
        with open(invalid_file, 'w') as f:
            f.write("invalid content")
        
        result = self.engine.validate_input_files(invalid_file, self.interview_file)
        self.assertFalse(result)
    
    def test_validate_input_files_invalid_interview_file(self):
        """测试无效面试名单文件验证"""
        invalid_file = os.path.join(self.temp_dir, "invalid.txt")
        with open(invalid_file, 'w') as f:
            f.write("invalid content")
        
        result = self.engine.validate_input_files(self.position_file, invalid_file)
        self.assertFalse(result)
    
    def test_process_files_with_empty_position_file(self):
        """测试空职位表文件的处理"""
        empty_position_file = os.path.join(self.temp_dir, "empty_positions.xlsx")
        empty_df = pd.DataFrame()
        empty_df.to_excel(empty_position_file, index=False)
        
        result = self.engine.process_files(
            position_file=empty_position_file,
            interview_file=self.interview_file,
            output_path=self.output_file
        )
        
        # 验证结果
        self.assertFalse(result.success)
        self.assertIn("输入文件验证失败", result.message)
    
    def test_process_files_with_empty_interview_file(self):
        """测试空面试名单文件的处理"""
        empty_interview_file = os.path.join(self.temp_dir, "empty_interviews.xlsx")
        empty_df = pd.DataFrame()
        empty_df.to_excel(empty_interview_file, index=False)
        
        result = self.engine.process_files(
            position_file=self.position_file,
            interview_file=empty_interview_file,
            output_path=self.output_file
        )
        
        # 验证结果
        self.assertFalse(result.success)
        self.assertIn("输入文件验证失败", result.message)
    
    def test_process_files_with_mismatched_data(self):
        """测试完全不匹配的数据处理"""
        # 创建完全不匹配的面试名单
        mismatched_data = {
            '姓名': ['张三', '李四'],
            '岗位名称': ['销售经理', '市场专员'],  # 与职位表完全不匹配
            '分数': [85.5, 78.0]
        }
        mismatched_df = pd.DataFrame(mismatched_data)
        mismatched_file = os.path.join(self.temp_dir, "mismatched_interviews.xlsx")
        mismatched_df.to_excel(mismatched_file, index=False)
        
        result = self.engine.process_files(
            position_file=self.position_file,
            interview_file=mismatched_file,
            output_path=self.output_file
        )
        
        # 验证结果 - 应该成功但所有岗位都是无法匹配状态
        self.assertTrue(result.success)
        self.assertGreater(len(result.results), 0)
        self.assertGreater(len(result.warnings), 0)  # 应该有警告信息
        
        # 验证所有岗位都是无法匹配状态
        for position_result in result.results:
            self.assertEqual(position_result.status, "无法匹配")
            self.assertIsNone(position_result.min_score)
            self.assertEqual(position_result.candidate_count, 0)
    
    def test_process_files_with_default_output_path(self):
        """测试使用默认输出路径的处理"""
        result = self.engine.process_files(
            position_file=self.position_file,
            interview_file=self.interview_file
        )
        
        # 验证结果
        self.assertTrue(result.success)
        self.assertIsNotNone(result.report_path)
        self.assertTrue(os.path.exists(result.report_path))
    
    def test_get_processing_status(self):
        """测试获取处理状态"""
        # 初始状态
        status = self.engine.get_processing_status()
        self.assertEqual(status['current_step'], 0)
        self.assertEqual(status['total_steps'], 7)
        self.assertEqual(status['progress_percentage'], 0)
        self.assertFalse(status['is_processing'])
        
        # 开始处理后的状态（通过mock验证）
        with patch.object(self.engine, '_update_progress') as mock_update:
            self.engine.process_files(
                position_file=self.position_file,
                interview_file=self.interview_file,
                output_path=self.output_file
            )
            
            # 验证_update_progress被调用
            self.assertTrue(mock_update.called)
    
    def test_reset_processing_state(self):
        """测试重置处理状态"""
        # 先执行一次处理
        self.engine.process_files(
            position_file=self.position_file,
            interview_file=self.interview_file,
            output_path=self.output_file
        )
        
        # 重置状态
        self.engine.reset_processing_state()
        
        # 验证状态被重置
        status = self.engine.get_processing_status()
        self.assertEqual(status['current_step'], 0)
        self.assertIsNone(status['start_time'])
        self.assertEqual(len(status['errors']), 0)
        self.assertEqual(len(status['warnings']), 0)
    
    def test_progress_callback_exception_handling(self):
        """测试进度回调函数异常处理"""
        # 创建会抛出异常的回调函数
        def failing_callback(progress):
            raise Exception("Callback error")
        
        engine = ProcessingEngine(progress_callback=failing_callback)
        
        # 执行处理，应该不会因为回调异常而失败
        result = engine.process_files(
            position_file=self.position_file,
            interview_file=self.interview_file,
            output_path=self.output_file
        )
        
        # 验证处理仍然成功
        self.assertTrue(result.success)
    
    def test_statistics_generation(self):
        """测试统计信息生成"""
        result = self.engine.process_files(
            position_file=self.position_file,
            interview_file=self.interview_file,
            output_path=self.output_file
        )
        
        # 验证统计信息结构
        self.assertIn('processing_summary', result.statistics)
        self.assertIn('score_statistics', result.statistics)
        self.assertIn('candidate_statistics', result.statistics)
        self.assertIn('match_statistics', result.statistics)
        self.assertIn('quality_metrics', result.statistics)
        
        # 验证处理摘要
        summary = result.statistics['processing_summary']
        self.assertIn('total_positions', summary)
        self.assertIn('normal_positions', summary)
        self.assertIn('success_rate', summary)
        
        # 验证质量指标
        quality = result.statistics['quality_metrics']
        self.assertIn('data_completeness', quality)
        self.assertIn('match_rate', quality)
        self.assertIn('error_rate', quality)
    
    def test_error_handling_in_score_calculation(self):
        """测试分数计算中的错误处理"""
        # 创建包含无效分数的面试数据
        invalid_interview_data = {
            '姓名': ['张三', '李四', '王五'],
            '岗位名称': ['软件工程师', '软件工程师', '产品经理'],
            '分数': [85.5, None, 'invalid']  # 包含无效分数
        }
        invalid_interview_df = pd.DataFrame(invalid_interview_data)
        invalid_interview_file = os.path.join(self.temp_dir, "invalid_scores.xlsx")
        invalid_interview_df.to_excel(invalid_interview_file, index=False)
        
        result = self.engine.process_files(
            position_file=self.position_file,
            interview_file=invalid_interview_file,
            output_path=self.output_file
        )
        
        # 验证处理成功但有警告（无效分数被清理掉了）
        self.assertTrue(result.success)
        self.assertGreater(len(result.warnings), 0)  # 应该有关于无效分数的警告
        
        # 验证结果中软件工程师岗位只有一个有效分数
        software_engineer_result = None
        for r in result.results:
            if r.position_name == '软件工程师':
                software_engineer_result = r
                break
        
        if software_engineer_result:
            self.assertEqual(software_engineer_result.min_score, 85.5)  # 只有张三的分数有效
            self.assertEqual(software_engineer_result.candidate_count, 1)  # 只有一个有效候选人


class TestProcessingEngineIntegration(unittest.TestCase):
    """ProcessingEngine集成测试 - 测试与真实Excel文件的集成"""
    
    def setUp(self):
        """测试前准备"""
        self.temp_dir = tempfile.mkdtemp()
        self.engine = ProcessingEngine()
    
    def tearDown(self):
        """测试后清理"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def test_complex_excel_processing(self):
        """测试复杂Excel文件处理"""
        # 创建复杂的测试数据
        position_file = os.path.join(self.temp_dir, "complex_positions.xlsx")
        interview_file = os.path.join(self.temp_dir, "complex_interviews.xlsx")
        
        # 创建多sheet职位表
        with pd.ExcelWriter(position_file) as writer:
            # Sheet 1: 技术岗位
            tech_positions = pd.DataFrame({
                '岗位代码': ['T001', 'T002', 'T003'],
                '岗位名称': ['高级软件工程师', 'Python开发工程师', '前端开发工程师'],
                '部门': ['技术部', '技术部', '技术部']
            })
            tech_positions.to_excel(writer, sheet_name='技术岗位', index=False)
            
            # Sheet 2: 业务岗位
            business_positions = pd.DataFrame({
                '岗位代码': ['B001', 'B002'],
                '岗位名称': ['产品经理', '业务分析师'],
                '部门': ['产品部', '业务部']
            })
            business_positions.to_excel(writer, sheet_name='业务岗位', index=False)
        
        # 创建复杂的面试数据
        complex_interview_data = pd.DataFrame({
            '姓名': ['张三', '李四', '王五', '赵六', '钱七', '孙八', '周九', '吴十'],
            '岗位名称': ['高级软件工程师', '高级软件工程师', 'Python开发工程师', 
                      '前端开发工程师', '产品经理', '产品经理', '业务分析师', '数据科学家'],
            '分数': [88.5, 92.0, 85.5, 78.0, 95.5, 87.0, 82.5, 90.0]
        })
        complex_interview_data.to_excel(interview_file, index=False)
        
        # 执行处理
        result = self.engine.process_files(
            position_file=position_file,
            interview_file=interview_file
        )
        
        # 验证结果
        self.assertTrue(result.success)
        self.assertGreater(len(result.results), 0)
        
        # 验证多sheet处理
        position_names = [r.position_name for r in result.results]
        self.assertIn('高级软件工程师', position_names)
        self.assertIn('Python开发工程师', position_names)
        self.assertIn('前端开发工程师', position_names)
        self.assertIn('产品经理', position_names)
        self.assertIn('业务分析师', position_names)
        
        # 验证分数计算
        for result_item in result.results:
            if result_item.position_name == '高级软件工程师':
                self.assertEqual(result_item.min_score, 88.5)  # 张三和李四的最低分
                self.assertEqual(result_item.candidate_count, 2)
            elif result_item.position_name == '产品经理':
                self.assertEqual(result_item.min_score, 87.0)  # 钱七和孙八的最低分
                self.assertEqual(result_item.candidate_count, 2)


if __name__ == '__main__':
    # 运行测试
    unittest.main(verbosity=2)