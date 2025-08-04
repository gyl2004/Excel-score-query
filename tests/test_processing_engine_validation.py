"""
ProcessingEngine输入验证和预处理测试
测试增强的验证和预处理功能
"""
import unittest
import tempfile
import os
import pandas as pd
from unittest.mock import Mock, patch
import sys

# 添加项目根目录到Python路径
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from services.processing_engine import ProcessingEngine, ProcessingEngineError


class TestProcessingEngineValidation(unittest.TestCase):
    """ProcessingEngine输入验证测试类"""
    
    def setUp(self):
        """测试前准备"""
        self.temp_dir = tempfile.mkdtemp()
        self.engine = ProcessingEngine()
        
        # 创建基础测试文件
        self.valid_position_file = os.path.join(self.temp_dir, "valid_positions.xlsx")
        self.valid_interview_file = os.path.join(self.temp_dir, "valid_interviews.xlsx")
        self._create_valid_test_files()
    
    def tearDown(self):
        """测试后清理"""
        import shutil
        shutil.rmtree(self.temp_dir, ignore_errors=True)
    
    def _create_valid_test_files(self):
        """创建有效的测试文件"""
        # 创建有效的职位表
        position_data = pd.DataFrame({
            '岗位代码': ['P001', 'P002'],
            '岗位名称': ['软件工程师', '产品经理'],
            '部门': ['技术部', '产品部']
        })
        position_data.to_excel(self.valid_position_file, index=False)
        
        # 创建有效的面试名单
        interview_data = pd.DataFrame({
            '姓名': ['张三', '李四'],
            '岗位名称': ['软件工程师', '产品经理'],
            '分数': [85.5, 92.0]
        })
        interview_data.to_excel(self.valid_interview_file, index=False)
    
    def test_pre_validate_file_paths_success(self):
        """测试成功的文件路径预验证"""
        result = self.engine._pre_validate_file_paths(
            self.valid_position_file, 
            self.valid_interview_file
        )
        self.assertTrue(result)
    
    def test_pre_validate_file_paths_empty_position_file(self):
        """测试空职位表文件路径"""
        result = self.engine._pre_validate_file_paths("", self.valid_interview_file)
        self.assertFalse(result)
        self.assertIn("职位表文件路径不能为空", self.engine._processing_errors)
    
    def test_pre_validate_file_paths_empty_interview_file(self):
        """测试空面试名单文件路径"""
        result = self.engine._pre_validate_file_paths(self.valid_position_file, "")
        self.assertFalse(result)
        self.assertIn("面试人员名单文件路径不能为空", self.engine._processing_errors)
    
    def test_pre_validate_file_paths_same_files(self):
        """测试相同的文件路径"""
        result = self.engine._pre_validate_file_paths(
            self.valid_position_file, 
            self.valid_position_file
        )
        self.assertFalse(result)
        self.assertTrue(any("不能是同一个文件" in error for error in self.engine._processing_errors))
    
    def test_validate_position_file_success(self):
        """测试成功的职位表文件验证"""
        result = self.engine._validate_position_file(self.valid_position_file)
        self.assertTrue(result)
    
    def test_validate_position_file_nonexistent(self):
        """测试不存在的职位表文件"""
        nonexistent_file = os.path.join(self.temp_dir, "nonexistent.xlsx")
        result = self.engine._validate_position_file(nonexistent_file)
        self.assertFalse(result)
        self.assertTrue(any("验证失败" in error for error in self.engine._processing_errors))
    
    def test_validate_position_file_no_position_columns(self):
        """测试没有职位相关列的文件"""
        invalid_file = os.path.join(self.temp_dir, "invalid_positions.xlsx")
        invalid_data = pd.DataFrame({
            '姓名': ['张三', '李四'],
            '年龄': [25, 30]
        })
        invalid_data.to_excel(invalid_file, index=False)
        
        result = self.engine._validate_position_file(invalid_file)
        self.assertFalse(result)
        self.assertTrue(any("没有包含有效职位数据的工作表" in error for error in self.engine._processing_errors))
    
    def test_validate_interview_file_success(self):
        """测试成功的面试名单文件验证"""
        result = self.engine._validate_interview_file(self.valid_interview_file)
        self.assertTrue(result)
    
    def test_validate_interview_file_missing_columns(self):
        """测试缺少必需列的面试名单文件"""
        invalid_file = os.path.join(self.temp_dir, "invalid_interviews.xlsx")
        invalid_data = pd.DataFrame({
            '姓名': ['张三', '李四'],
            '年龄': [25, 30]  # 缺少岗位和分数列
        })
        invalid_data.to_excel(invalid_file, index=False)
        
        result = self.engine._validate_interview_file(invalid_file)
        self.assertFalse(result)
        self.assertTrue(any("缺少必需的列类型" in error for error in self.engine._processing_errors))
    
    def test_validate_interview_file_empty_data(self):
        """测试空数据的面试名单文件"""
        empty_file = os.path.join(self.temp_dir, "empty_interviews.xlsx")
        empty_data = pd.DataFrame(columns=['姓名', '岗位名称', '分数'])
        empty_data.to_excel(empty_file, index=False)
        
        result = self.engine._validate_interview_file(empty_file)
        self.assertFalse(result)
        self.assertTrue(any("没有数据" in error for error in self.engine._processing_errors))
    
    def test_validate_file_compatibility_success(self):
        """测试成功的文件兼容性验证"""
        result = self.engine._validate_file_compatibility(
            self.valid_position_file, 
            self.valid_interview_file
        )
        self.assertTrue(result)
    
    def test_validate_file_compatibility_no_matches(self):
        """测试完全不匹配的文件兼容性"""
        # 创建完全不匹配的面试文件
        mismatched_file = os.path.join(self.temp_dir, "mismatched_interviews.xlsx")
        mismatched_data = pd.DataFrame({
            '姓名': ['张三', '李四'],
            '岗位名称': ['销售经理', '市场专员'],  # 完全不匹配
            '分数': [85.5, 92.0]
        })
        mismatched_data.to_excel(mismatched_file, index=False)
        
        result = self.engine._validate_file_compatibility(
            self.valid_position_file, 
            mismatched_file
        )
        self.assertTrue(result)  # 兼容性验证不会失败，只会产生警告
        self.assertGreater(len(self.engine._processing_warnings), 0)
    
    def test_calculate_similarity(self):
        """测试字符串相似度计算"""
        # 完全相同
        self.assertEqual(self.engine._calculate_similarity("软件工程师", "软件工程师"), 1.0)
        
        # 部分相似
        similarity = self.engine._calculate_similarity("软件工程师", "高级软件工程师")
        self.assertGreater(similarity, 0.5)
        self.assertLess(similarity, 1.0)
        
        # 完全不同
        similarity = self.engine._calculate_similarity("软件工程师", "销售经理")
        self.assertLess(similarity, 0.5)


class TestProcessingEnginePreprocessing(unittest.TestCase):
    """ProcessingEngine数据预处理测试类"""
    
    def setUp(self):
        """测试前准备"""
        self.engine = ProcessingEngine()
    
    def test_standardize_text_normal(self):
        """测试正常文本标准化"""
        result = self.engine._standardize_text("  软件工程师  ")
        self.assertEqual(result, "软件工程师")
    
    def test_standardize_text_with_special_chars(self):
        """测试包含特殊字符的文本标准化"""
        result = self.engine._standardize_text("软件@工程师#")
        self.assertEqual(result, "软件工程师")
    
    def test_standardize_text_multiple_spaces(self):
        """测试多个空格的文本标准化"""
        result = self.engine._standardize_text("软件   工程师")
        self.assertEqual(result, "软件 工程师")
    
    def test_standardize_text_none_value(self):
        """测试None值的文本标准化"""
        result = self.engine._standardize_text(None)
        self.assertEqual(result, "")
    
    def test_standardize_text_nan_string(self):
        """测试'nan'字符串的文本标准化"""
        result = self.engine._standardize_text("nan")
        self.assertEqual(result, "")
    
    def test_standardize_score_valid(self):
        """测试有效分数标准化"""
        self.assertEqual(self.engine._standardize_score(85.5), 85.5)
        self.assertEqual(self.engine._standardize_score("92.3"), 92.3)
        self.assertEqual(self.engine._standardize_score(100), 100.0)
        self.assertEqual(self.engine._standardize_score(0), 0.0)
    
    def test_standardize_score_invalid(self):
        """测试无效分数标准化"""
        self.assertIsNone(self.engine._standardize_score(None))
        self.assertIsNone(self.engine._standardize_score("invalid"))
        self.assertIsNone(self.engine._standardize_score(-10))
        self.assertIsNone(self.engine._standardize_score(150))
        self.assertIsNone(self.engine._standardize_score("nan"))
    
    def test_standardize_score_precision(self):
        """测试分数精度标准化"""
        result = self.engine._standardize_score(85.123456)
        self.assertEqual(result, 85.12)
    
    def test_preprocess_position_data(self):
        """测试职位数据预处理"""
        raw_data = [
            {
                'position_code': '  P001  ',
                'position_name': '软件@工程师',
                'department': '技术部',
                'sheet_name': 'Sheet1',
                'row_index': 1
            },
            {
                'position_code': 'P002',
                'position_name': '软件工程师',  # 重复的职位名称
                'department': '技术部',
                'sheet_name': 'Sheet1',
                'row_index': 2
            },
            {
                'position_code': '',
                'position_name': '',  # 空的职位名称，应该被过滤
                'department': '技术部',
                'sheet_name': 'Sheet1',
                'row_index': 3
            }
        ]
        
        result = self.engine._preprocess_position_data(raw_data)
        
        # 验证结果
        self.assertEqual(len(result), 1)  # 只保留一个有效且不重复的职位
        self.assertEqual(result[0]['position_code'], 'P001')
        self.assertEqual(result[0]['position_name'], '软件工程师')
        self.assertGreater(len(self.engine._processing_warnings), 0)  # 应该有重复警告
    
    def test_preprocess_interview_data(self):
        """测试面试人员数据预处理"""
        raw_data = [
            {
                'name': '  张三  ',
                'position_name': '软件@工程师',
                'score': 85.5,
                'is_qualified': True,
                'row_index': 1
            },
            {
                'name': '张三',
                'position_name': '软件工程师',
                'score': 90.0,  # 同一人同一岗位的更高分数
                'is_qualified': True,
                'row_index': 2
            },
            {
                'name': '李四',
                'position_name': '产品经理',
                'score': None,  # 无效分数，应该被过滤
                'is_qualified': False,
                'row_index': 3
            }
        ]
        
        result = self.engine._preprocess_interview_data(raw_data)
        
        # 验证结果
        self.assertEqual(len(result), 1)  # 只保留一个有效记录（张三的更高分数）
        self.assertEqual(result[0]['name'], '张三')
        self.assertEqual(result[0]['position_name'], '软件工程师')
        self.assertEqual(result[0]['score'], 90.0)
        self.assertGreater(len(self.engine._processing_warnings), 0)  # 应该有重复和无效数据警告


class TestProcessingEngineResultValidation(unittest.TestCase):
    """ProcessingEngine结果验证测试类"""
    
    def setUp(self):
        """测试前准备"""
        self.engine = ProcessingEngine()
    
    def test_validate_processing_results_success(self):
        """测试成功的结果验证"""
        from models.data_models import PositionScoreResult
        
        results = [
            PositionScoreResult(
                position_code='P001',
                position_name='软件工程师',
                department='技术部',
                min_score=85.5,
                candidate_count=2,
                status='正常',
                notes='匹配成功'
            ),
            PositionScoreResult(
                position_code='P002',
                position_name='产品经理',
                department='产品部',
                min_score=92.0,
                candidate_count=1,
                status='正常',
                notes='匹配成功'
            )
        ]
        
        validation_report = self.engine.validate_processing_results(results)
        
        # 验证结果
        self.assertTrue(validation_report['is_valid'])
        self.assertEqual(validation_report['total_positions'], 2)
        self.assertEqual(len(validation_report['validation_errors']), 0)
        self.assertEqual(validation_report['quality_metrics']['normal_positions'], 2)
        self.assertEqual(validation_report['quality_metrics']['success_rate'], '100.0%')
    
    def test_validate_processing_results_empty(self):
        """测试空结果验证"""
        validation_report = self.engine.validate_processing_results([])
        
        # 验证结果
        self.assertFalse(validation_report['is_valid'])
        self.assertIn("处理结果为空", validation_report['validation_errors'])
    
    def test_validate_processing_results_with_errors(self):
        """测试包含错误的结果验证"""
        from models.data_models import PositionScoreResult
        
        results = [
            PositionScoreResult(
                position_code='P001',
                position_name='',  # 空的职位名称
                department='技术部',
                min_score=85.5,
                candidate_count=-1,  # 无效的候选人数量
                status='正常',
                notes='匹配成功'
            ),
            PositionScoreResult(
                position_code='P002',
                position_name='产品经理',
                department='产品部',
                min_score=None,
                candidate_count=0,
                status='无法匹配',
                notes='无法匹配'
            )
        ]
        
        validation_report = self.engine.validate_processing_results(results)
        
        # 验证结果
        self.assertFalse(validation_report['is_valid'])
        self.assertGreater(len(validation_report['validation_errors']), 0)
        self.assertGreater(len(validation_report['validation_warnings']), 0)
        self.assertGreater(len(validation_report['recommendations']), 0)
    
    def test_validate_processing_results_low_success_rate(self):
        """测试低成功率的结果验证"""
        from models.data_models import PositionScoreResult
        
        results = [
            PositionScoreResult(
                position_code='P001',
                position_name='软件工程师',
                department='技术部',
                min_score=85.5,
                candidate_count=1,
                status='正常',
                notes='匹配成功'
            ),
            PositionScoreResult(
                position_code='P002',
                position_name='产品经理',
                department='产品部',
                min_score=None,
                candidate_count=0,
                status='无法匹配',
                notes='无法匹配'
            ),
            PositionScoreResult(
                position_code='P003',
                position_name='数据分析师',
                department='数据部',
                min_score=None,
                candidate_count=0,
                status='数据异常',
                notes='数据异常'
            )
        ]
        
        validation_report = self.engine.validate_processing_results(results)
        
        # 验证结果
        self.assertTrue(validation_report['is_valid'])  # 有警告但仍然有效
        self.assertGreater(len(validation_report['validation_warnings']), 0)
        self.assertIn("成功处理率较低", validation_report['validation_warnings'][0])
        self.assertGreater(len(validation_report['recommendations']), 0)


if __name__ == '__main__':
    # 运行测试
    unittest.main(verbosity=2)