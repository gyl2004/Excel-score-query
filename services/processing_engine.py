"""
业务逻辑引擎
整合文件读取、数据匹配、分数计算和报告生成的主要业务流程控制器
"""
import logging
import os
from typing import Dict, List, Optional, Callable, Tuple, Any
from dataclasses import dataclass
from datetime import datetime

from models.data_models import Position, InterviewCandidate, PositionScoreResult
from services.excel_reader import ExcelReader, ExcelProcessingError
from services.data_matcher import DataMatcher, DataMatchingError
from services.data_validator import DataValidator, ValidationError
from services.report_generator import ReportGenerator


class ProcessingEngineError(Exception):
    """业务处理引擎相关异常"""
    pass


@dataclass
class ProcessingProgress:
    """处理进度数据类"""
    current_step: str
    step_number: int
    total_steps: int
    progress_percentage: float
    message: str
    details: Optional[Dict] = None


@dataclass
class ProcessingResult:
    """处理结果数据类"""
    success: bool
    message: str
    results: List[PositionScoreResult]
    statistics: Dict[str, Any]
    report_path: Optional[str] = None
    processing_time: Optional[float] = None
    errors: List[str] = None
    warnings: List[str] = None


class ProcessingEngine:
    """业务逻辑处理引擎"""
    
    def __init__(self, progress_callback: Optional[Callable[[ProcessingProgress], None]] = None):
        """
        初始化处理引擎
        
        Args:
            progress_callback: 进度回调函数，用于更新处理进度
        """
        self.logger = logging.getLogger(__name__)
        self.progress_callback = progress_callback
        
        # 初始化各个组件
        self.excel_reader = ExcelReader()
        self.data_matcher = DataMatcher()
        self.data_validator = DataValidator()
        self.report_generator = ReportGenerator()
        
        # 处理状态
        self._current_step = 0
        self._total_steps = 7  # 总共7个处理步骤
        self._start_time = None
        self._processing_errors = []
        self._processing_warnings = []
        
        self.logger.info("ProcessingEngine 初始化完成")
    
    def process_files(self, position_file: str, interview_file: str, 
                     column_mappings: Optional[dict] = None,
                     output_path: Optional[str] = None) -> ProcessingResult:
        """
        处理Excel文件的主要业务流程
        
        Args:
            position_file: 职位表文件路径
            interview_file: 面试人员名单文件路径
            column_mappings: 列映射配置，格式为 {岗位表列名: 面试表列名}
            output_path: 输出报告文件路径，如果为None则使用默认路径
            
        Returns:
            ProcessingResult: 处理结果
            
        Raises:
            ProcessingEngineError: 处理过程中发生错误时抛出异常
        """
        self._start_time = datetime.now()
        self._processing_errors.clear()
        self._processing_warnings.clear()
        
        try:
            self.logger.info(f"开始处理文件 - 职位表: {position_file}, 面试名单: {interview_file}")
            
            # 步骤1: 验证输入文件
            self._update_progress(1, "验证输入文件", "检查文件路径和格式...")
            if not self.validate_input_files(position_file, interview_file):
                raise ProcessingEngineError("输入文件验证失败")
            
            # 步骤2: 读取职位表文件
            self._update_progress(2, "读取职位表", "正在读取职位表数据...")
            position_data = self._read_position_file(position_file)
            
            # 步骤3: 读取面试人员名单文件
            self._update_progress(3, "读取面试名单", "正在读取面试人员数据...")
            interview_data = self._read_interview_file(interview_file)
            
            # 步骤4: 数据匹配
            self._update_progress(4, "匹配岗位数据", "正在匹配职位表和面试名单...")
            match_results = self._match_position_data(position_data, interview_data, column_mappings, position_file, interview_file)
            
            # 步骤5: 处理岗位分数
            self._update_progress(5, "处理岗位分数", "正在整理各岗位面试人员分数...")
            score_results = self._calculate_min_scores(match_results, interview_data)
            
            # 步骤6: 生成报告
            self._update_progress(6, "生成报告", "正在生成Excel报告...")
            report_path = self._generate_report(score_results, output_path)
            
            # 步骤7: 验证处理结果
            self._update_progress(7, "验证结果", "正在验证处理结果质量...")
            validation_report = self.validate_processing_results(score_results)
            
            # 记录验证警告
            if validation_report.get('validation_warnings'):
                self._processing_warnings.extend(validation_report['validation_warnings'])
            
            # 如果验证失败，记录错误但不中断流程
            if not validation_report.get('is_valid', True):
                for error in validation_report.get('validation_errors', []):
                    self._processing_warnings.append(f"结果验证警告: {error}")
            
            # 计算处理时间
            processing_time = (datetime.now() - self._start_time).total_seconds()
            
            # 生成统计信息
            statistics = self._generate_statistics(score_results, match_results)
            statistics['validation_report'] = validation_report
            
            result = ProcessingResult(
                success=True,
                message="文件处理成功完成",
                results=score_results,
                statistics=statistics,
                report_path=report_path,
                processing_time=processing_time,
                errors=self._processing_errors.copy(),
                warnings=self._processing_warnings.copy()
            )
            
            self.logger.info(f"文件处理完成，耗时: {processing_time:.2f}秒，"
                           f"处理岗位: {len(score_results)}个，报告路径: {report_path}")
            
            return result
            
        except Exception as e:
            processing_time = (datetime.now() - self._start_time).total_seconds() if self._start_time else 0
            error_msg = f"文件处理失败: {str(e)}"
            self.logger.error(error_msg)
            
            return ProcessingResult(
                success=False,
                message=error_msg,
                results=[],
                statistics={},
                processing_time=processing_time,
                errors=self._processing_errors + [str(e)],
                warnings=self._processing_warnings.copy()
            )
    
    def validate_input_files(self, position_file: str, interview_file: str) -> bool:
        """
        验证输入文件的有效性
        
        Args:
            position_file: 职位表文件路径
            interview_file: 面试人员名单文件路径
            
        Returns:
            bool: 文件是否有效
        """
        try:
            self.logger.info("开始验证输入文件")
            
            # 预验证文件路径
            if not self._pre_validate_file_paths(position_file, interview_file):
                return False
            
            # 验证职位表文件
            if not self._validate_position_file(position_file):
                return False
            
            # 验证面试人员名单文件
            if not self._validate_interview_file(interview_file):
                return False
            
            # 验证文件兼容性
            if not self._validate_file_compatibility(position_file, interview_file):
                return False
            
            self.logger.info("输入文件验证完成")
            return True
            
        except Exception as e:
            error_msg = f"验证输入文件时发生错误: {str(e)}"
            self._processing_errors.append(error_msg)
            self.logger.error(error_msg)
            return False
    
    def _pre_validate_file_paths(self, position_file: str, interview_file: str) -> bool:
        """
        预验证文件路径
        
        Args:
            position_file: 职位表文件路径
            interview_file: 面试人员名单文件路径
            
        Returns:
            bool: 路径是否有效
        """
        try:
            # 检查文件路径是否为空
            if not position_file or not position_file.strip():
                error_msg = "职位表文件路径不能为空"
                self._processing_errors.append(error_msg)
                self.logger.error(error_msg)
                return False
            
            if not interview_file or not interview_file.strip():
                error_msg = "面试人员名单文件路径不能为空"
                self._processing_errors.append(error_msg)
                self.logger.error(error_msg)
                return False
            
            # 检查文件路径是否相同
            if os.path.abspath(position_file) == os.path.abspath(interview_file):
                error_msg = "职位表文件和面试人员名单文件不能是同一个文件"
                self._processing_errors.append(error_msg)
                self.logger.error(error_msg)
                return False
            
            # 检查文件大小（避免处理过大的文件）
            max_file_size = 50 * 1024 * 1024  # 50MB
            
            if os.path.exists(position_file):
                position_size = os.path.getsize(position_file)
                if position_size > max_file_size:
                    warning_msg = f"职位表文件较大 ({position_size / 1024 / 1024:.1f}MB)，处理可能较慢"
                    self._processing_warnings.append(warning_msg)
                    self.logger.warning(warning_msg)
            
            if os.path.exists(interview_file):
                interview_size = os.path.getsize(interview_file)
                if interview_size > max_file_size:
                    warning_msg = f"面试人员名单文件较大 ({interview_size / 1024 / 1024:.1f}MB)，处理可能较慢"
                    self._processing_warnings.append(warning_msg)
                    self.logger.warning(warning_msg)
            
            return True
            
        except Exception as e:
            error_msg = f"预验证文件路径时发生错误: {str(e)}"
            self._processing_errors.append(error_msg)
            self.logger.error(error_msg)
            return False
    
    def _validate_position_file(self, position_file: str) -> bool:
        """
        验证职位表文件
        
        Args:
            position_file: 职位表文件路径
            
        Returns:
            bool: 文件是否有效
        """
        try:
            # 基础文件验证
            self.excel_reader.validate_file_path(position_file)
            self.excel_reader.check_excel_format(position_file)
            
            # 预读取验证数据结构
            try:
                sheet_names = self.excel_reader.get_sheet_names(position_file)
                if not sheet_names:
                    error_msg = "职位表文件没有工作表"
                    self._processing_errors.append(error_msg)
                    self.logger.error(error_msg)
                    return False
                
                # 检查至少有一个工作表包含有效数据
                valid_sheets = 0
                for sheet_name in sheet_names:
                    try:
                        df = self.excel_reader.read_excel_sheet(position_file, sheet_name)
                        if not df.empty:
                            # 简单检查是否包含职位相关列
                            columns_str = ' '.join([str(col).lower() for col in df.columns])
                            if any(keyword in columns_str for keyword in ['岗位', '职位', 'position']):
                                valid_sheets += 1
                    except Exception as e:
                        self.logger.warning(f"跳过无效工作表 '{sheet_name}': {str(e)}")
                
                if valid_sheets == 0:
                    error_msg = "职位表文件没有包含有效职位数据的工作表"
                    self._processing_errors.append(error_msg)
                    self.logger.error(error_msg)
                    return False
                
                self.logger.info(f"职位表文件验证通过: {position_file}，包含 {valid_sheets} 个有效工作表")
                
            except Exception as e:
                error_msg = f"职位表文件数据结构验证失败: {str(e)}"
                self._processing_errors.append(error_msg)
                self.logger.error(error_msg)
                return False
            
            return True
            
        except ExcelProcessingError as e:
            error_msg = f"职位表文件验证失败: {str(e)}"
            self._processing_errors.append(error_msg)
            self.logger.error(error_msg)
            return False
    
    def _validate_interview_file(self, interview_file: str) -> bool:
        """
        验证面试人员名单文件
        
        Args:
            interview_file: 面试人员名单文件路径
            
        Returns:
            bool: 文件是否有效
        """
        try:
            # 基础文件验证
            self.excel_reader.validate_file_path(interview_file)
            self.excel_reader.check_excel_format(interview_file)
            
            # 预读取验证数据结构
            try:
                df = self.excel_reader.read_excel_sheet(interview_file)
                if df.empty:
                    error_msg = "面试人员名单文件没有数据"
                    self._processing_errors.append(error_msg)
                    self.logger.error(error_msg)
                    return False
                
                # 检查必需的列
                columns_str = ' '.join([str(col).lower() for col in df.columns])
                required_keywords = {
                    'name': ['姓名', '名字', '面试人员', 'name'],
                    'position': ['岗位', '职位', 'position'],
                    'score': ['分数', '成绩', 'score']
                }
                
                missing_types = []
                for col_type, keywords in required_keywords.items():
                    if not any(keyword in columns_str for keyword in keywords):
                        missing_types.append(col_type)
                
                if missing_types:
                    error_msg = f"面试人员名单文件缺少必需的列类型: {', '.join(missing_types)}"
                    self._processing_errors.append(error_msg)
                    self.logger.error(error_msg)
                    return False
                
                # 检查数据行数
                if len(df) == 0:
                    error_msg = "面试人员名单文件没有数据行"
                    self._processing_errors.append(error_msg)
                    self.logger.error(error_msg)
                    return False
                
                # 检查数据质量
                non_empty_rows = df.dropna(how='all')
                if len(non_empty_rows) == 0:
                    error_msg = "面试人员名单文件所有数据行都为空"
                    self._processing_errors.append(error_msg)
                    self.logger.error(error_msg)
                    return False
                
                if len(non_empty_rows) < len(df) * 0.5:
                    warning_msg = f"面试人员名单文件有较多空行 ({len(df) - len(non_empty_rows)} 行)"
                    self._processing_warnings.append(warning_msg)
                    self.logger.warning(warning_msg)
                
                self.logger.info(f"面试人员名单文件验证通过: {interview_file}，包含 {len(non_empty_rows)} 行有效数据")
                
            except Exception as e:
                error_msg = f"面试人员名单文件数据结构验证失败: {str(e)}"
                self._processing_errors.append(error_msg)
                self.logger.error(error_msg)
                return False
            
            return True
            
        except ExcelProcessingError as e:
            error_msg = f"面试人员名单文件验证失败: {str(e)}"
            self._processing_errors.append(error_msg)
            self.logger.error(error_msg)
            return False
    
    def _validate_file_compatibility(self, position_file: str, interview_file: str) -> bool:
        """
        验证文件兼容性
        
        Args:
            position_file: 职位表文件路径
            interview_file: 面试人员名单文件路径
            
        Returns:
            bool: 文件是否兼容
        """
        try:
            self.logger.info("开始验证文件兼容性")
            
            # 快速预览数据以检查兼容性
            try:
                # 读取职位表的第一个工作表的前几行
                position_sheets = self.excel_reader.get_sheet_names(position_file)
                position_preview = self.excel_reader.read_excel_sheet(position_file, position_sheets[0])
                if len(position_preview) > 10:
                    position_preview = position_preview.head(10)
                
                # 读取面试名单的前几行
                interview_preview = self.excel_reader.read_excel_sheet(interview_file)
                if len(interview_preview) > 10:
                    interview_preview = interview_preview.head(10)
                
                # 提取职位名称样本
                position_names = set()
                for col in position_preview.columns:
                    if any(keyword in str(col).lower() for keyword in ['岗位', '职位', 'position']):
                        names = position_preview[col].dropna().astype(str).str.strip()
                        position_names.update(names.tolist())
                        break
                
                # 提取面试岗位名称样本
                interview_positions = set()
                for col in interview_preview.columns:
                    if any(keyword in str(col).lower() for keyword in ['岗位', '职位', 'position']):
                        names = interview_preview[col].dropna().astype(str).str.strip()
                        interview_positions.update(names.tolist())
                        break
                
                # 检查是否有潜在的匹配
                if position_names and interview_positions:
                    # 精确匹配检查
                    exact_matches = position_names.intersection(interview_positions)
                    
                    # 模糊匹配检查
                    fuzzy_matches = 0
                    for pos_name in position_names:
                        for int_name in interview_positions:
                            if self._calculate_similarity(pos_name, int_name) > 0.6:
                                fuzzy_matches += 1
                                break
                    
                    total_potential_matches = len(exact_matches) + fuzzy_matches
                    
                    if total_potential_matches == 0:
                        warning_msg = "职位表和面试名单的岗位名称可能完全不匹配，请检查数据一致性"
                        self._processing_warnings.append(warning_msg)
                        self.logger.warning(warning_msg)
                    else:
                        self.logger.info(f"发现 {len(exact_matches)} 个精确匹配，{fuzzy_matches} 个潜在模糊匹配")
                
            except Exception as e:
                warning_msg = f"文件兼容性检查时发生错误，将继续处理: {str(e)}"
                self._processing_warnings.append(warning_msg)
                self.logger.warning(warning_msg)
            
            self.logger.info("文件兼容性验证完成")
            return True
            
        except Exception as e:
            error_msg = f"验证文件兼容性时发生错误: {str(e)}"
            self._processing_errors.append(error_msg)
            self.logger.error(error_msg)
            return False
    
    def _calculate_similarity(self, str1: str, str2: str) -> float:
        """
        计算两个字符串的相似度
        
        Args:
            str1: 字符串1
            str2: 字符串2
            
        Returns:
            float: 相似度 (0-1)
        """
        try:
            from difflib import SequenceMatcher
            return SequenceMatcher(None, str1.lower(), str2.lower()).ratio()
        except Exception:
            return 0.0
    
    def _read_position_file(self, position_file: str) -> List[Dict]:
        """
        读取职位表文件
        
        Args:
            position_file: 职位表文件路径
            
        Returns:
            List[Dict]: 职位数据列表
            
        Raises:
            ProcessingEngineError: 读取失败时抛出异常
        """
        try:
            # 读取职位表数据
            position_sheets = self.excel_reader.read_position_file(position_file)
            
            # 提取职位信息
            position_data = self.excel_reader.extract_position_info(position_sheets)
            
            if not position_data:
                raise ProcessingEngineError("职位表中没有有效的职位数据")
            
            # 数据预处理
            position_data = self._preprocess_position_data(position_data)
            
            # 验证职位数据完整性
            validation_result = self.data_validator.validate_position_data_integrity(position_data)
            
            if not validation_result.is_valid:
                error_msg = f"职位数据验证失败: {'; '.join(validation_result.errors)}"
                raise ProcessingEngineError(error_msg)
            
            # 记录警告信息
            if validation_result.warnings:
                self._processing_warnings.extend(validation_result.warnings)
            
            self.logger.info(f"成功读取职位数据，共 {len(position_data)} 个职位")
            return position_data
            
        except (ExcelProcessingError, ValidationError) as e:
            raise ProcessingEngineError(f"读取职位表文件失败: {str(e)}")
        except Exception as e:
            raise ProcessingEngineError(f"读取职位表文件时发生未知错误: {str(e)}")
    
    def _read_interview_file(self, interview_file: str) -> List[Dict]:
        """
        读取面试人员名单文件
        
        Args:
            interview_file: 面试人员名单文件路径
            
        Returns:
            List[Dict]: 面试人员数据列表
            
        Raises:
            ProcessingEngineError: 读取失败时抛出异常
        """
        try:
            # 读取面试人员数据
            interview_df = self.excel_reader.read_interview_file(interview_file)
            
            # 提取面试信息
            interview_data = self.excel_reader.extract_interview_info(interview_df)
            
            if not interview_data:
                raise ProcessingEngineError("面试人员名单中没有有效的面试数据")
            
            # 数据预处理
            interview_data = self._preprocess_interview_data(interview_data)
            
            # 验证面试数据完整性
            validation_result = self.data_validator.validate_interview_data_integrity(interview_data)
            
            if not validation_result.is_valid:
                error_msg = f"面试数据验证失败: {'; '.join(validation_result.errors)}"
                raise ProcessingEngineError(error_msg)
            
            # 记录警告信息
            if validation_result.warnings:
                self._processing_warnings.extend(validation_result.warnings)
            
            self.logger.info(f"成功读取面试数据，共 {len(interview_data)} 个面试人员")
            return interview_data
            
        except (ExcelProcessingError, ValidationError) as e:
            raise ProcessingEngineError(f"读取面试人员名单文件失败: {str(e)}")
        except Exception as e:
            raise ProcessingEngineError(f"读取面试人员名单文件时发生未知错误: {str(e)}")
    
    def _match_position_data(self, position_data: List[Dict], interview_data: List[Dict], column_mappings: Optional[dict] = None, position_file: Optional[str] = None, interview_file: Optional[str] = None) -> Dict:
        """
        匹配职位数据和面试数据
        
        Args:
            position_data: 职位数据列表
            interview_data: 面试人员数据列表
            column_mappings: 列映射配置
            position_file: 职位表文件路径（可配置匹配器需要）
            interview_file: 面试名单文件路径（可配置匹配器需要）
            
        Returns:
            Dict: 匹配结果
            
        Raises:
            ProcessingEngineError: 匹配失败时抛出异常
        """
        try:
            # 验证数据一致性
            consistency_result = self.data_validator.validate_data_consistency(position_data, interview_data)
            
            # 记录一致性警告，但不因为没有匹配而失败
            if consistency_result.warnings:
                self._processing_warnings.extend(consistency_result.warnings)
            
            # 即使没有完全匹配的岗位，也继续执行匹配过程
            if not consistency_result.is_valid:
                # 记录警告但不抛出异常
                warning_msg = f"数据一致性较低: {'; '.join(consistency_result.errors)}"
                self._processing_warnings.append(warning_msg)
                self.logger.warning(warning_msg)
            
            # 执行岗位匹配
            if column_mappings:
                # 使用可配置的数据匹配器
                from .configurable_data_matcher import ConfigurableDataMatcher
                import pandas as pd
                
                # 检查是否提供了文件路径
                if not position_file or not interview_file:
                    raise ProcessingEngineError("使用可配置匹配器时需要提供文件路径")
                
                # 直接读取原始DataFrame，避免数据转换导致的列名问题
                self.logger.info("重新读取原始Excel文件以获取正确的DataFrame格式")
                position_df = self.excel_reader.read_excel_sheet(position_file)
                interview_df = self.excel_reader.read_excel_sheet(interview_file)
                
                self.logger.info(f"岗位表列名: {list(position_df.columns)}")
                self.logger.info(f"面试表列名: {list(interview_df.columns)}")
                
                # 创建可配置匹配器
                fuzzy_threshold = 0.8  # 默认模糊匹配阈值
                configurable_matcher = ConfigurableDataMatcher(column_mappings, fuzzy_threshold)
                
                # 执行匹配
                configurable_results = configurable_matcher.match_data(position_df, interview_df)
                
                # 转换结果格式以兼容现有代码
                match_results = self._convert_configurable_results(configurable_results)
            else:
                # 使用原有的数据匹配器
                match_results = self.data_matcher.match_positions(position_data, interview_data)
            
            # 验证匹配结果质量
            match_validation = self.data_validator.validate_match_results(match_results)
            
            if match_validation.warnings:
                self._processing_warnings.extend(match_validation.warnings)
            
            # 记录匹配统计信息
            statistics = match_results.get('statistics', {})
            self.logger.info(f"岗位匹配完成，匹配率: {statistics.get('match_rate', 0):.1%}，"
                           f"成功匹配: {statistics.get('matched_positions', 0)} 个，"
                           f"未匹配: {statistics.get('unmatched_positions', 0)} 个")
            
            return match_results
            
        except (DataMatchingError, ValidationError) as e:
            raise ProcessingEngineError(f"岗位数据匹配失败: {str(e)}")
        except Exception as e:
            raise ProcessingEngineError(f"岗位数据匹配时发生未知错误: {str(e)}")
    
    def _convert_configurable_results(self, configurable_results: dict) -> dict:
        """
        转换可配置匹配器的结果格式以兼容现有代码
        
        Args:
            configurable_results: 可配置匹配器的结果
            
        Returns:
            转换后的匹配结果
        """
        try:
            from .data_matcher import PositionMapping
            
            match_results = configurable_results['match_results']
            statistics = configurable_results['statistics']
            
            # 转换为PositionMapping对象列表
            mappings = []
            unmatched_positions = []
            
            for result in match_results:
                if result.matched:
                    # 创建PositionMapping对象
                    position_data = result.position_row
                    mapping = PositionMapping(
                        position_code=str(position_data.get('职位代码', position_data.get('岗位代码', ''))),
                        position_name=position_data.get('招考职位', position_data.get('岗位名称', '')),
                        department=position_data.get('用人司局', ''),  # 用人司局
                        department_name=position_data.get('部门名称', ''),  # 部门名称
                        recruit_count=int(position_data.get('招考人数', 0)) if position_data.get('招考人数') else 0,
                        sheet_name=position_data.get('sheet_name', ''),
                        interview_position=result.interview_rows[0].get('招考职位', '') if result.interview_rows else '',
                        match_confidence=result.match_score,
                        candidate_count=len(result.interview_rows)
                    )
                    mappings.append(mapping)
                else:
                    unmatched_positions.append(result.position_row)
            
            # 构建兼容的结果格式
            return {
                'mappings': mappings,
                'unmatched_positions': unmatched_positions,
                'unmatched_interviews': [],  # 可配置匹配器不直接提供这个信息
                'statistics': {
                    'total_positions': statistics['total_positions'],
                    'matched_positions': statistics['matched_positions'],
                    'unmatched_positions': statistics['unmatched_positions'],
                    'match_rate': statistics['match_rate']
                }
            }
            
        except Exception as e:
            self.logger.error(f"转换匹配结果格式时发生错误: {e}")
            # 返回空结果而不是抛出异常
            return {
                'mappings': [],
                'unmatched_positions': [],
                'unmatched_interviews': [],
                'statistics': {
                    'total_positions': 0,
                    'matched_positions': 0,
                    'unmatched_positions': 0,
                    'match_rate': 0.0
                }
            }
    
    def _calculate_min_scores(self, match_results: Dict, interview_data: List[Dict]) -> List[PositionScoreResult]:
        """
        获取各岗位面试人员的分数列表（以岗位表为主）
        
        Args:
            match_results: 匹配结果
            interview_data: 面试人员数据列表
            
        Returns:
            List[PositionScoreResult]: 岗位分数结果列表
            
        Raises:
            ProcessingEngineError: 处理失败时抛出异常
        """
        try:
            results = []
            mappings = match_results.get('mappings', [])
            unmatched_positions = match_results.get('unmatched_positions', [])
            
            # 处理所有岗位（包括匹配和未匹配的）
            all_positions = []
            
            # 添加匹配成功的岗位
            for mapping in mappings:
                all_positions.append({
                    'type': 'matched',
                    'data': mapping
                })
            
            # 添加未匹配的岗位
            for position in unmatched_positions:
                all_positions.append({
                    'type': 'unmatched',
                    'data': position
                })
            
            # 为每个岗位（无论是否匹配）生成结果
            for position_info in all_positions:
                try:
                    if position_info['type'] == 'matched':
                        # 处理匹配成功的岗位
                        mapping = position_info['data']
                        result = self._process_matched_position(mapping, interview_data)
                    else:
                        # 处理未匹配的岗位
                        position = position_info['data']
                        result = self._process_unmatched_position(position)
                    
                    results.append(result)
                    
                except Exception as e:
                    # 处理单个岗位计算错误，不中断整个流程
                    position_name = "未知岗位"
                    try:
                        if position_info['type'] == 'matched':
                            position_name = position_info['data'].position_name
                        else:
                            position_name = position_info['data'].get('招考职位', '未知岗位')
                    except:
                        pass
                    
                    error_msg = f"处理岗位 '{position_name}' 时发生错误: {str(e)}"
                    self._processing_warnings.append(error_msg)
                    self.logger.warning(error_msg)
            
            self.logger.info(f"岗位分数处理完成，共处理 {len(results)} 个岗位")
            return results
            
        except Exception as e:
            raise ProcessingEngineError(f"处理岗位分数时发生错误: {str(e)}")
    
    def _process_matched_position(self, mapping, interview_data: List[Dict]) -> 'PositionScoreResult':
        """
        处理匹配成功的岗位
        
        Args:
            mapping: 岗位映射信息
            interview_data: 面试人员数据列表
            
        Returns:
            PositionScoreResult: 岗位处理结果
        """
        # 获取该岗位的所有面试人员
        position_candidates = [
            interview for interview in interview_data
            if interview.get('position_name') == mapping.interview_position
        ]
        
        if not position_candidates:
            # 没有面试人员的情况
            return PositionScoreResult(
                position_code=mapping.position_code,
                position_name=mapping.position_name,
                department=mapping.department,
                department_name=mapping.department_name,
                recruit_count=mapping.recruit_count,
                candidate_count=0,
                min_score=None,
                status="无面试人员",
                notes=f"该岗位在面试名单中没有找到面试人员",
                all_scores=[]
            )
        
        # 获取所有面试人员的分数
        scores = []
        candidate_details = []
        
        for candidate in position_candidates:
            score = candidate.get('score')
            name = candidate.get('name', '未知')
            
            # 安全地处理分数数据
            if score is not None:
                try:
                    # 确保分数是数字类型
                    if isinstance(score, (int, float)):
                        score_value = float(score)
                    else:
                        # 如果是字符串，尝试转换
                        score_str = str(score).strip() if hasattr(score, 'strip') else str(score)
                        score_value = float(score_str)
                    
                    scores.append(score_value)
                    candidate_details.append(f"{name}: {score_value}")
                except (ValueError, TypeError, AttributeError) as e:
                    self.logger.warning(f"跳过无效分数: {score} (类型: {type(score)}) - {e}")
                    continue
        
        if not scores:
            return PositionScoreResult(
                position_code=mapping.position_code,
                position_name=mapping.position_name,
                department=mapping.department,
                department_name=mapping.department_name,
                recruit_count=mapping.recruit_count,
                candidate_count=len(position_candidates),
                min_score=None,
                status="数据异常",
                notes="面试人员存在但没有有效分数数据",
                all_scores=[]
            )
        else:
            # 不计算最低分，而是列出所有分数
            sorted_scores = sorted(scores, reverse=True)
            scores_summary = f"共{len(scores)}人，分数: {sorted_scores}"
            return PositionScoreResult(
                position_code=mapping.position_code,
                position_name=mapping.position_name,
                department=mapping.department,
                department_name=mapping.department_name,
                recruit_count=mapping.recruit_count,
                candidate_count=len(scores),
                min_score=None,  # 不设置最低分
                status="正常",
                notes=scores_summary,
                all_scores=sorted_scores  # 存储所有分数
            )
    
    def _process_unmatched_position(self, position: Dict) -> 'PositionScoreResult':
        """
        处理未匹配的岗位
        
        Args:
            position: 岗位数据字典
            
        Returns:
            PositionScoreResult: 岗位处理结果
        """
        return PositionScoreResult(
            position_code=str(position.get('职位代码', position.get('岗位代码', ''))),
            position_name=position.get('招考职位', position.get('岗位名称', '')),
            department=position.get('用人司局', ''),
            department_name=position.get('部门名称', ''),
            recruit_count=int(position.get('招考人数', 0)) if position.get('招考人数') else 0,
            candidate_count=0,
            min_score=None,
            status="无法匹配",
            notes="在面试名单中找不到对应的岗位",
            all_scores=[]
        )
    
    def _generate_report(self, results: List[PositionScoreResult], output_path: Optional[str]) -> str:
        """
        生成Excel报告
        
        Args:
            results: 岗位分数结果列表
            output_path: 输出路径
            
        Returns:
            str: 实际保存的报告文件路径
            
        Raises:
            ProcessingEngineError: 生成报告失败时抛出异常
        """
        try:
            # 如果没有指定输出路径，使用默认路径
            if not output_path:
                output_path = self.report_generator.get_recommended_save_path()
            
            # 验证输出路径
            is_valid, error_msg = self.report_generator.validate_output_path(output_path)
            if not is_valid:
                raise ProcessingEngineError(f"输出路径无效: {error_msg}")
            
            # 生成报告
            success, actual_path = self.report_generator.generate_report(results, output_path)
            
            if not success:
                raise ProcessingEngineError("报告生成失败")
            
            self.logger.info(f"报告生成成功，保存路径: {actual_path}")
            return actual_path
            
        except Exception as e:
            if isinstance(e, ProcessingEngineError):
                raise
            else:
                raise ProcessingEngineError(f"生成报告时发生错误: {str(e)}")
    
    def _generate_statistics(self, results: List[PositionScoreResult], match_results: Dict) -> Dict[str, Any]:
        """
        生成处理统计信息
        
        Args:
            results: 岗位分数结果列表
            match_results: 匹配结果
            
        Returns:
            Dict[str, Any]: 统计信息
        """
        try:
            # 基本统计
            total_positions = len(results)
            normal_positions = len([r for r in results if r.status == "正常"])
            no_interview_positions = len([r for r in results if r.status == "无面试人员"])
            unmatched_positions = len([r for r in results if r.status == "无法匹配"])
            error_positions = len([r for r in results if r.status == "数据异常"])
            
            # 分数统计
            valid_scores = [r.min_score for r in results if r.min_score is not None]
            score_stats = {}
            if valid_scores:
                score_stats = {
                    'min_score': min(valid_scores),
                    'max_score': max(valid_scores),
                    'avg_score': sum(valid_scores) / len(valid_scores),
                    'score_count': len(valid_scores)
                }
            
            # 面试人员统计
            total_candidates = sum(r.candidate_count for r in results)
            positions_with_candidates = len([r for r in results if r.candidate_count > 0])
            
            # 匹配统计
            match_stats = match_results.get('statistics', {})
            
            statistics = {
                'processing_summary': {
                    'total_positions': total_positions,
                    'normal_positions': normal_positions,
                    'no_interview_positions': no_interview_positions,
                    'unmatched_positions': unmatched_positions,
                    'error_positions': error_positions,
                    'success_rate': f"{(normal_positions / total_positions * 100):.1f}%" if total_positions > 0 else "0%"
                },
                'score_statistics': score_stats,
                'candidate_statistics': {
                    'total_candidates': total_candidates,
                    'positions_with_candidates': positions_with_candidates,
                    'avg_candidates_per_position': total_candidates / positions_with_candidates if positions_with_candidates > 0 else 0
                },
                'match_statistics': match_stats,
                'quality_metrics': {
                    'data_completeness': f"{(normal_positions / total_positions * 100):.1f}%" if total_positions > 0 else "0%",
                    'match_rate': f"{(match_stats.get('match_rate', 0) * 100):.1f}%",
                    'error_rate': f"{(error_positions / total_positions * 100):.1f}%" if total_positions > 0 else "0%"
                }
            }
            
            return statistics
            
        except Exception as e:
            self.logger.warning(f"生成统计信息时发生错误: {str(e)}")
            return {'error': f"统计信息生成失败: {str(e)}"}
    
    def _update_progress(self, step_number: int, step_name: str, message: str, details: Optional[Dict] = None):
        """
        更新处理进度
        
        Args:
            step_number: 当前步骤编号
            step_name: 步骤名称
            message: 进度消息
            details: 详细信息
        """
        self._current_step = step_number
        progress_percentage = (step_number / self._total_steps) * 100
        
        progress = ProcessingProgress(
            current_step=step_name,
            step_number=step_number,
            total_steps=self._total_steps,
            progress_percentage=progress_percentage,
            message=message,
            details=details
        )
        
        self.logger.info(f"处理进度 [{step_number}/{self._total_steps}] {step_name}: {message}")
        
        # 调用进度回调函数
        if self.progress_callback:
            try:
                self.progress_callback(progress)
            except Exception as e:
                self.logger.warning(f"进度回调函数执行失败: {str(e)}")
    
    def get_processing_status(self) -> Dict[str, Any]:
        """
        获取当前处理状态
        
        Returns:
            Dict[str, Any]: 处理状态信息
        """
        return {
            'current_step': self._current_step,
            'total_steps': self._total_steps,
            'progress_percentage': (self._current_step / self._total_steps) * 100,
            'is_processing': self._start_time is not None,
            'start_time': self._start_time.isoformat() if self._start_time else None,
            'errors': self._processing_errors.copy(),
            'warnings': self._processing_warnings.copy()
        }
    
    def _preprocess_position_data(self, position_data: List[Dict]) -> List[Dict]:
        """
        预处理职位数据
        
        Args:
            position_data: 原始职位数据列表
            
        Returns:
            List[Dict]: 预处理后的职位数据列表
        """
        try:
            self.logger.info("开始预处理职位数据")
            processed_data = []
            
            for position in position_data:
                try:
                    # 数据标准化
                    processed_position = {
                        'position_code': self._standardize_text(position.get('position_code', '')),
                        'position_name': self._standardize_text(position.get('position_name', '')),
                        'department': self._standardize_text(position.get('department', '')),
                        'sheet_name': position.get('sheet_name', ''),
                        'row_index': position.get('row_index', 0)
                    }
                    
                    # 数据清理
                    if processed_position['position_name']:  # 只保留有效的职位名称
                        # 保留所有有效职位，不进行去重（职位代码可能对应多个职位，需要结合其他条件筛选）
                        processed_data.append(processed_position)
                    
                except Exception as e:
                    warning_msg = f"预处理职位数据时跳过无效记录: {str(e)}"
                    self._processing_warnings.append(warning_msg)
                    self.logger.warning(warning_msg)
            
            self.logger.info(f"职位数据预处理完成，保留 {len(processed_data)} 个有效职位")
            return processed_data
            
        except Exception as e:
            raise ProcessingEngineError(f"预处理职位数据时发生错误: {str(e)}")
    
    def _preprocess_interview_data(self, interview_data: List[Dict]) -> List[Dict]:
        """
        预处理面试人员数据
        
        Args:
            interview_data: 原始面试人员数据列表
            
        Returns:
            List[Dict]: 预处理后的面试人员数据列表
        """
        try:
            self.logger.info("开始预处理面试人员数据")
            processed_data = []
            
            for interview in interview_data:
                try:
                    # 数据标准化
                    processed_interview = {
                        'name': self._standardize_text(interview.get('name', '')),
                        'position_name': self._standardize_text(interview.get('position_name', '')),
                        'score': self._standardize_score(interview.get('score')),
                        'is_qualified': interview.get('is_qualified', False),
                        'row_index': interview.get('row_index', 0)
                    }
                    
                    # 数据验证和清理
                    if (processed_interview['name'] and 
                        processed_interview['position_name'] and 
                        processed_interview['score'] is not None):
                        
                        # 保留所有有效面试人员，不进行去重（一个职位可能有多个人进入面试）
                        processed_data.append(processed_interview)
                    
                except Exception as e:
                    warning_msg = f"预处理面试人员数据时跳过无效记录: {str(e)}"
                    self._processing_warnings.append(warning_msg)
                    self.logger.warning(warning_msg)
            
            self.logger.info(f"面试人员数据预处理完成，保留 {len(processed_data)} 个有效面试人员")
            return processed_data
            
        except Exception as e:
            raise ProcessingEngineError(f"预处理面试人员数据时发生错误: {str(e)}")
    
    def _standardize_text(self, text: Any) -> str:
        """
        标准化文本数据
        
        Args:
            text: 原始文本
            
        Returns:
            str: 标准化后的文本
        """
        if text is None:
            return ""
        
        try:
            # 转换为字符串并清理
            text_str = str(text).strip()
            
            # 移除多余的空白字符
            import re
            text_str = re.sub(r'\s+', ' ', text_str)
            
            # 移除特殊字符（保留中文、英文、数字和常用标点）
            text_str = re.sub(r'[^\w\u4e00-\u9fff\s\-\(\)（）]', '', text_str)
            
            # 处理常见的编码问题
            if text_str.lower() in ['nan', 'null', 'none', '']:
                return ""
            
            return text_str.strip()
            
        except Exception:
            return ""
    
    def _standardize_score(self, score: Any) -> Optional[float]:
        """
        标准化分数数据
        
        Args:
            score: 原始分数
            
        Returns:
            Optional[float]: 标准化后的分数，无效时返回None
        """
        if score is None:
            return None
        
        try:
            # 尝试转换为浮点数
            if isinstance(score, str):
                score = score.strip()
                if score.lower() in ['nan', 'null', 'none', '']:
                    return None
            
            score_float = float(score)
            
            # 验证分数范围（只检查非负数）
            if score_float >= 0:
                return round(score_float, 2)  # 保留两位小数
            else:
                return None
                
        except (ValueError, TypeError):
            return None
    
    def validate_processing_results(self, results: List[PositionScoreResult]) -> Dict[str, Any]:
        """
        验证处理结果的质量
        
        Args:
            results: 处理结果列表
            
        Returns:
            Dict[str, Any]: 验证结果报告
        """
        try:
            self.logger.info("开始验证处理结果")
            
            validation_report = {
                'is_valid': True,
                'total_positions': len(results),
                'validation_errors': [],
                'validation_warnings': [],
                'quality_metrics': {},
                'recommendations': []
            }
            
            if not results:
                validation_report['is_valid'] = False
                validation_report['validation_errors'].append("处理结果为空")
                return validation_report
            
            # 统计各种状态的岗位
            status_counts = {}
            score_values = []
            
            for result in results:
                status = result.status
                status_counts[status] = status_counts.get(status, 0) + 1
                
                if result.min_score is not None:
                    score_values.append(result.min_score)
                
                # 验证单个结果的完整性
                if not result.position_name:
                    validation_report['validation_errors'].append(f"发现空的岗位名称")
                
                if result.candidate_count < 0:
                    validation_report['validation_errors'].append(f"岗位 '{result.position_name}' 的候选人数量无效: {result.candidate_count}")
            
            # 计算质量指标
            normal_count = status_counts.get('正常', 0)
            success_rate = normal_count / len(results) if results else 0
            
            validation_report['quality_metrics'] = {
                'success_rate': f"{success_rate:.1%}",
                'normal_positions': normal_count,
                'no_interview_positions': status_counts.get('无面试人员', 0),
                'unmatched_positions': status_counts.get('无法匹配', 0),
                'error_positions': status_counts.get('数据异常', 0),
                'avg_score': sum(score_values) / len(score_values) if score_values else 0,
                'min_score': min(score_values) if score_values else None,
                'max_score': max(score_values) if score_values else None
            }
            
            # 生成警告和建议
            if success_rate < 0.5:
                validation_report['validation_warnings'].append(f"成功处理率较低 ({success_rate:.1%})，建议检查数据质量")
                validation_report['recommendations'].append("检查职位表和面试名单的岗位名称一致性")
            
            if status_counts.get('无法匹配', 0) > 0:
                validation_report['validation_warnings'].append(f"有 {status_counts['无法匹配']} 个岗位无法匹配")
                validation_report['recommendations'].append("考虑使用更宽松的匹配条件或手动映射")
            
            if status_counts.get('数据异常', 0) > 0:
                validation_report['validation_warnings'].append(f"有 {status_counts['数据异常']} 个岗位存在数据异常")
                validation_report['recommendations'].append("检查面试数据的完整性和格式")
            
            # 如果有错误，标记为无效
            if validation_report['validation_errors']:
                validation_report['is_valid'] = False
            
            self.logger.info(f"处理结果验证完成，成功率: {success_rate:.1%}")
            return validation_report
            
        except Exception as e:
            error_msg = f"验证处理结果时发生错误: {str(e)}"
            self.logger.error(error_msg)
            return {
                'is_valid': False,
                'validation_errors': [error_msg],
                'validation_warnings': [],
                'quality_metrics': {},
                'recommendations': ['重新检查数据和处理流程']
            }
    
    def reset_processing_state(self):
        """重置处理状态"""
        self._current_step = 0
        self._start_time = None
        self._processing_errors.clear()
        self._processing_warnings.clear()
        self.data_validator.clear_validation_results()
        self.logger.info("处理状态已重置")