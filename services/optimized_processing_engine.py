"""
优化版业务逻辑引擎
针对大数据量场景进行性能优化
"""
import logging
import os
import time
from typing import Dict, List, Optional, Callable, Tuple, Any
from dataclasses import dataclass
from datetime import datetime
import pandas as pd
import numpy as np

from models.data_models import Position, InterviewCandidate, PositionScoreResult
from services.excel_reader import ExcelReader, ExcelProcessingError
from services.data_validator import DataValidator, ValidationError
from services.report_generator import ReportGenerator


class OptimizedProcessingEngine:
    """优化版业务逻辑处理引擎"""
    
    def __init__(self, progress_callback: Optional[Callable] = None):
        """初始化处理引擎"""
        self.logger = logging.getLogger(__name__)
        self.progress_callback = progress_callback
        
        # 初始化组件
        self.excel_reader = ExcelReader()
        self.data_validator = DataValidator()
        self.report_generator = ReportGenerator()
        
        # 缓存数据，避免重复读取
        self._position_cache = None
        self._interview_cache = None
        self._interview_index = None  # 用于快速查找的索引
        
        self.logger.info("OptimizedProcessingEngine 初始化完成")
    
    def process_files_optimized(self, position_file: str, interview_file: str, 
                               column_mappings: Optional[dict] = None,
                               output_path: Optional[str] = None,
                               selected_columns: Optional[List[str]] = None) -> Dict[str, Any]:
        """
        优化版文件处理流程
        
        Args:
            position_file: 职位表文件路径
            interview_file: 面试人员名单文件路径
            column_mappings: 列映射配置
            output_path: 输出报告文件路径
            
        Returns:
            处理结果字典
        """
        start_time = time.time()
        
        try:
            self.logger.info("开始优化版文件处理流程")
            
            # 步骤1: 快速读取和缓存数据
            self._update_progress("读取数据文件", 10)
            position_df, interview_df = self._load_data_optimized(position_file, interview_file)
            
            # 步骤2: 构建快速查找索引
            self._update_progress("构建数据索引", 30)
            self._build_interview_index(interview_df, column_mappings)
            
            # 步骤3: 批量匹配处理
            self._update_progress("执行批量匹配", 60)
            match_results = self._batch_match_optimized(position_df, interview_df, column_mappings)
            
            # 步骤4: 生成结果
            self._update_progress("生成处理结果", 90)
            results = self._generate_optimized_results(match_results, output_path, selected_columns)
            
            processing_time = time.time() - start_time
            results['processing_time'] = processing_time
            
            self._update_progress("处理完成", 100)
            self.logger.info(f"优化版处理完成，耗时: {processing_time:.2f}秒")
            
            return results
            
        except Exception as e:
            error_msg = f"优化版处理失败: {str(e)}"
            self.logger.error(error_msg)
            return {
                'success': False,
                'message': error_msg,
                'processing_time': time.time() - start_time
            }
    
    def _load_data_optimized(self, position_file: str, interview_file: str) -> Tuple[pd.DataFrame, pd.DataFrame]:
        """
        优化版数据加载，使用缓存避免重复读取
        """
        # 检查缓存
        if self._position_cache is None:
            self.logger.info("读取职位表数据...")
            self._position_cache = self.excel_reader.read_excel_sheet(position_file)
            self.logger.info(f"职位表读取完成: {len(self._position_cache)} 行")
        
        if self._interview_cache is None:
            self.logger.info("读取面试表数据...")
            # 使用更高效的读取方式
            self._interview_cache = pd.read_excel(
                interview_file, 
                engine='openpyxl',
                dtype=str  # 统一使用字符串类型，避免类型转换开销
            )
            self.logger.info(f"面试表读取完成: {len(self._interview_cache)} 行")
        
        return self._position_cache.copy(), self._interview_cache.copy()
    
    def _build_interview_index(self, interview_df: pd.DataFrame, column_mappings: dict):
        """
        构建面试数据的快速查找索引
        """
        if self._interview_index is not None:
            return
        
        self.logger.info("构建面试数据索引...")
        
        # 为每个映射列创建索引
        self._interview_index = {}
        
        for pos_col, int_col in column_mappings.items():
            if int_col in interview_df.columns:
                # 创建分组索引，用于快速查找
                grouped = interview_df.groupby(int_col)
                self._interview_index[int_col] = {
                    name: group.index.tolist() 
                    for name, group in grouped
                }
                self.logger.info(f"为列 '{int_col}' 创建索引，包含 {len(self._interview_index[int_col])} 个唯一值")
        
        self.logger.info("数据索引构建完成")
    
    def _batch_match_optimized(self, position_df: pd.DataFrame, interview_df: pd.DataFrame, 
                              column_mappings: dict) -> List[Dict]:
        """
        优化版批量匹配处理
        """
        results = []
        total_positions = len(position_df)
        
        self.logger.info(f"开始批量匹配，共 {total_positions} 个岗位")
        
        # 批量处理，减少日志输出
        batch_size = max(10, total_positions // 20)  # 动态调整批次大小
        
        for i in range(0, total_positions, batch_size):
            batch_end = min(i + batch_size, total_positions)
            batch_positions = position_df.iloc[i:batch_end]
            
            # 批量匹配
            batch_results = self._match_batch_positions(batch_positions, interview_df, column_mappings)
            results.extend(batch_results)
            
            # 更新进度（减少频率）
            progress = 60 + (batch_end / total_positions) * 30
            self._update_progress(f"匹配进度: {batch_end}/{total_positions}", progress)
        
        self.logger.info(f"批量匹配完成，处理了 {len(results)} 个岗位")
        return results
    
    def _match_batch_positions(self, batch_positions: pd.DataFrame, interview_df: pd.DataFrame, 
                              column_mappings: dict) -> List[Dict]:
        """
        匹配一批岗位
        """
        batch_results = []
        
        for _, pos_row in batch_positions.iterrows():
            # 使用索引进行快速查找
            matched_indices = self._find_matches_with_index(pos_row, column_mappings)
            
            if matched_indices:
                # 获取匹配的面试数据
                matched_interviews = interview_df.iloc[matched_indices]
                
                # 计算最低分数
                score_col = self._find_score_column(interview_df)
                if score_col:
                    scores = pd.to_numeric(matched_interviews[score_col], errors='coerce')
                    min_score = scores.min() if not scores.isna().all() else 0
                else:
                    min_score = 0
                
                result = {
                    'position_name': pos_row.get('招考职位', ''),
                    'position_code': pos_row.get('职位代码', ''),
                    'department': pos_row.get('用人司局', ''),
                    'min_score': min_score,
                    'candidate_count': len(matched_interviews),
                    'matched': True,
                    'original_position_data': pos_row.to_dict()  # 保存完整的原始数据
                }
            else:
                result = {
                    'position_name': pos_row.get('招考职位', ''),
                    'position_code': pos_row.get('职位代码', ''),
                    'department': pos_row.get('用人司局', ''),
                    'min_score': 0,
                    'candidate_count': 0,
                    'matched': False,
                    'original_position_data': pos_row.to_dict()  # 保存完整的原始数据
                }
            
            batch_results.append(result)
        
        return batch_results
    
    def _find_matches_with_index(self, pos_row: pd.Series, column_mappings: dict) -> List[int]:
        """
        使用索引快速查找匹配的面试记录
        """
        if not self._interview_index:
            return []
        
        # 获取所有匹配条件的交集
        matched_indices = None
        
        for pos_col, int_col in column_mappings.items():
            pos_value = str(pos_row.get(pos_col, ''))
            
            if not pos_value or pos_value == 'nan':
                continue
            
            # 从索引中查找匹配的记录
            if int_col in self._interview_index:
                col_matches = self._interview_index[int_col].get(pos_value, [])
                
                if matched_indices is None:
                    matched_indices = set(col_matches)
                else:
                    matched_indices = matched_indices.intersection(set(col_matches))
                
                # 如果交集为空，直接返回
                if not matched_indices:
                    return []
        
        return list(matched_indices) if matched_indices else []
    
    def _find_score_column(self, df: pd.DataFrame) -> Optional[str]:
        """
        查找分数列
        """
        score_keywords = ['最低面试分数', '分数', '成绩', 'score']
        
        for col in df.columns:
            col_str = str(col).strip()
            if any(keyword in col_str for keyword in score_keywords):
                return col
        
        return None
    
    def _generate_optimized_results(self, match_results: List[Dict], output_path: Optional[str], 
                                   selected_columns: Optional[List[str]] = None) -> Dict:
        """
        生成优化版处理结果
        """
        # 统计信息
        total_positions = len(match_results)
        matched_positions = len([r for r in match_results if r['matched']])
        total_candidates = sum(r['candidate_count'] for r in match_results if r['matched'])
        
        # 生成Excel报告
        if output_path:
            self._create_optimized_excel(match_results, output_path, selected_columns)
        
        return {
            'success': True,
            'message': '优化版处理完成',
            'total_positions': total_positions,
            'matched_positions': matched_positions,
            'unmatched_positions': total_positions - matched_positions,
            'total_candidates': total_candidates,
            'match_rate': matched_positions / total_positions if total_positions > 0 else 0,
            'results': match_results,
            'output_file': output_path
        }
    
    def _create_optimized_excel(self, results: List[Dict], output_path: str, 
                               selected_columns: Optional[List[str]] = None):
        """
        创建优化版Excel报告
        """
        # 转换为DataFrame
        df_data = []
        for result in results:
            # 构建完整的数据行 - 需要包含所有原始职位数据
            full_row_data = {}
            
            # 如果result中包含原始职位数据，使用它
            if 'original_position_data' in result:
                # 添加所有原始职位数据
                for key, value in result['original_position_data'].items():
                    if value is None or (isinstance(value, float) and pd.isna(value)):
                        processed_value = ''
                    else:
                        processed_value = str(value).strip()
                    full_row_data[key] = processed_value
            else:
                # 回退到基本数据
                full_row_data.update({
                    '招考职位': result['position_name'],
                    '职位代码': str(result['position_code']),
                    '用人司局': result['department']
                })
            
            # 添加处理结果数据
            full_row_data['最低面试分数'] = result['min_score']
            full_row_data['面试人数'] = result['candidate_count']
            full_row_data['状态'] = '成功' if result['matched'] else '未匹配'
            
            # 确保职位代码以文本格式保存
            position_code = str(result['position_code'])
            if position_code and position_code != 'nan':
                if position_code.replace('.', '').replace('E', '').replace('+', '').replace('-', '').isdigit():
                    position_code = "'" + position_code
            full_row_data['职位代码'] = position_code
            
            # 根据用户选择的列过滤数据
            if selected_columns:
                filtered_row = {}
                for col in selected_columns:
                    value = full_row_data.get(col, '')
                    if value is None or (isinstance(value, float) and pd.isna(value)):
                        value = ''
                    filtered_row[col] = value
                df_data.append(filtered_row)
            else:
                # 使用默认列
                df_data.append(full_row_data)
        
        df = pd.DataFrame(df_data)
        
        # 确保职位代码列为文本类型
        df['职位代码'] = df['职位代码'].astype(str)
        
        # 使用openpyxl引擎进行更精确的格式控制
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='岗位分数汇总', index=False)
            
            # 获取工作表对象，设置职位代码列为文本格式
            worksheet = writer.sheets['岗位分数汇总']
            
            # 设置职位代码列为文本格式
            from openpyxl.utils import get_column_letter
            
            # 找到职位代码列的位置
            headers = list(df.columns)
            if '职位代码' in headers:
                col_idx = headers.index('职位代码') + 1  # Excel列从1开始
                col_letter = get_column_letter(col_idx)
                
                # 设置整列为文本格式
                for row in range(1, len(df) + 2):  # +2因为包含标题行
                    cell = worksheet[f'{col_letter}{row}']
                    cell.number_format = '@'  # '@'表示文本格式
                    
                    # 如果是数据行且包含职位代码，确保显示为完整数字
                    if row > 1:  # 跳过标题行
                        cell_value = str(cell.value)
                        if cell_value.startswith("'"):
                            cell.value = cell_value[1:]  # 移除前导单引号
            
            # 自动调整列宽
            for column in worksheet.columns:
                max_length = 0
                column_letter = column[0].column_letter
                for cell in column:
                    try:
                        if len(str(cell.value)) > max_length:
                            max_length = len(str(cell.value))
                    except:
                        pass
                adjusted_width = min(max_length + 2, 30)
                worksheet.column_dimensions[column_letter].width = adjusted_width
        
        self.logger.info(f"Excel报告已生成: {output_path}")
    
    def _update_progress(self, message: str, percentage: float):
        """更新进度"""
        if self.progress_callback:
            self.progress_callback(message, percentage)
        
        # 减少日志输出频率
        if percentage % 20 == 0 or percentage >= 100:
            self.logger.info(f"进度: {percentage:.0f}% - {message}")
    
    def clear_cache(self):
        """清理缓存"""
        self._position_cache = None
        self._interview_cache = None
        self._interview_index = None
        self.logger.info("缓存已清理")