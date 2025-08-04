"""
快速数据匹配器
针对大数据量场景优化的高性能匹配器
"""
import logging
import time
from typing import Dict, List, Optional, Set, Any
import pandas as pd
import numpy as np
from dataclasses import dataclass


@dataclass
class FastMatchResult:
    """快速匹配结果"""
    position_row: Dict[str, Any]
    interview_rows: List[Dict[str, Any]]
    match_score: float
    matched: bool


class FastDataMatcher:
    """快速数据匹配器"""
    
    def __init__(self, column_mappings: Dict[str, str]):
        """
        初始化快速数据匹配器
        
        Args:
            column_mappings: 列映射关系，格式为 {岗位表列名: 面试表列名}
        """
        self.logger = logging.getLogger(__name__)
        self.column_mappings = column_mappings
        self.interview_indices = {}  # 预构建的索引
        
        self.logger.info(f"初始化快速数据匹配器，列映射: {len(column_mappings)} 个")
    
    def build_indices(self, interview_df: pd.DataFrame):
        """
        预构建面试数据的索引，用于快速查找
        
        Args:
            interview_df: 面试数据DataFrame
        """
        start_time = time.time()
        self.logger.info("开始构建面试数据索引...")
        
        # 为每个映射的面试表列构建索引
        for pos_col, int_col in self.column_mappings.items():
            if int_col in interview_df.columns:
                # 使用字典索引，将值映射到行索引列表
                self.interview_indices[int_col] = {}
                
                # 分组构建索引
                for idx, value in enumerate(interview_df[int_col]):
                    str_value = str(value).strip() if pd.notna(value) else ""
                    if str_value:
                        if str_value not in self.interview_indices[int_col]:
                            self.interview_indices[int_col][str_value] = []
                        self.interview_indices[int_col][str_value].append(idx)
                
                unique_values = len(self.interview_indices[int_col])
                self.logger.info(f"列 '{int_col}' 索引构建完成: {unique_values} 个唯一值")
        
        build_time = time.time() - start_time
        self.logger.info(f"索引构建完成，耗时: {build_time:.2f}秒")
    
    def match_data_fast(self, position_df: pd.DataFrame, interview_df: pd.DataFrame) -> Dict[str, Any]:
        """
        快速匹配数据
        
        Args:
            position_df: 岗位表DataFrame
            interview_df: 面试表DataFrame
            
        Returns:
            匹配结果字典
        """
        start_time = time.time()
        
        # 如果索引未构建，先构建索引
        if not self.interview_indices:
            self.build_indices(interview_df)
        
        self.logger.info(f"开始快速匹配，岗位: {len(position_df)} 个，面试: {len(interview_df)} 个")
        
        match_results = []
        matched_count = 0
        
        # 批量处理岗位
        batch_size = 50  # 批量大小
        total_batches = (len(position_df) + batch_size - 1) // batch_size
        
        for batch_idx in range(total_batches):
            start_idx = batch_idx * batch_size
            end_idx = min(start_idx + batch_size, len(position_df))
            batch_positions = position_df.iloc[start_idx:end_idx]
            
            # 处理当前批次
            for _, pos_row in batch_positions.iterrows():
                result = self._match_single_position_fast(pos_row, interview_df)
                match_results.append(result)
                
                if result.matched:
                    matched_count += 1
            
            # 每处理10个批次输出一次进度
            if (batch_idx + 1) % 10 == 0 or batch_idx == total_batches - 1:
                progress = (batch_idx + 1) / total_batches * 100
                self.logger.info(f"匹配进度: {progress:.1f}% ({end_idx}/{len(position_df)})")
        
        match_time = time.time() - start_time
        
        # 生成统计信息
        statistics = {
            'total_positions': len(position_df),
            'matched_positions': matched_count,
            'unmatched_positions': len(position_df) - matched_count,
            'match_rate': matched_count / len(position_df) if len(position_df) > 0 else 0,
            'total_candidates': sum(len(r.interview_rows) for r in match_results if r.matched),
            'match_time': match_time
        }
        
        self.logger.info(f"快速匹配完成，耗时: {match_time:.2f}秒，匹配率: {statistics['match_rate']:.2%}")
        
        return {
            'match_results': match_results,
            'statistics': statistics
        }
    
    def _match_single_position_fast(self, pos_row: pd.Series, interview_df: pd.DataFrame) -> FastMatchResult:
        """
        快速匹配单个岗位
        
        Args:
            pos_row: 岗位行数据
            interview_df: 面试数据DataFrame
            
        Returns:
            匹配结果
        """
        # 获取所有匹配条件的交集
        candidate_indices = None
        
        for pos_col, int_col in self.column_mappings.items():
            pos_value = str(pos_row.get(pos_col, '')).strip()
            
            if not pos_value or pos_value == 'nan':
                continue
            
            # 从预构建的索引中查找
            if int_col in self.interview_indices:
                col_indices = self.interview_indices[int_col].get(pos_value, [])
                
                if candidate_indices is None:
                    candidate_indices = set(col_indices)
                else:
                    candidate_indices = candidate_indices.intersection(set(col_indices))
                
                # 如果交集为空，直接返回未匹配
                if not candidate_indices:
                    break
        
        # 构建结果
        if candidate_indices:
            # 获取匹配的面试记录
            matched_interviews = interview_df.iloc[list(candidate_indices)]
            interview_rows = matched_interviews.to_dict('records')
            
            return FastMatchResult(
                position_row=pos_row.to_dict(),
                interview_rows=interview_rows,
                match_score=1.0,  # 精确匹配
                matched=True
            )
        else:
            return FastMatchResult(
                position_row=pos_row.to_dict(),
                interview_rows=[],
                match_score=0.0,
                matched=False
            )
    
    def get_match_statistics(self, match_results: List[FastMatchResult]) -> Dict[str, Any]:
        """
        获取匹配统计信息
        
        Args:
            match_results: 匹配结果列表
            
        Returns:
            统计信息字典
        """
        total_positions = len(match_results)
        matched_positions = len([r for r in match_results if r.matched])
        total_candidates = sum(len(r.interview_rows) for r in match_results if r.matched)
        
        # 按面试人数分组统计
        candidate_counts = [len(r.interview_rows) for r in match_results if r.matched]
        
        statistics = {
            'total_positions': total_positions,
            'matched_positions': matched_positions,
            'unmatched_positions': total_positions - matched_positions,
            'match_rate': matched_positions / total_positions if total_positions > 0 else 0,
            'total_candidates': total_candidates,
            'avg_candidates_per_position': total_candidates / matched_positions if matched_positions > 0 else 0,
            'max_candidates_per_position': max(candidate_counts) if candidate_counts else 0,
            'min_candidates_per_position': min(candidate_counts) if candidate_counts else 0
        }
        
        return statistics
    
    def export_results_fast(self, match_results: List[FastMatchResult], output_path: str, 
                           selected_columns: Optional[List[str]] = None):
        """
        快速导出匹配结果
        
        Args:
            match_results: 匹配结果列表
            output_path: 输出文件路径
            selected_columns: 用户选择的输出列，如果为None则使用默认列
        """
        start_time = time.time()
        
        # 准备导出数据
        export_data = []
        
        for result in match_results:
            pos_data = result.position_row
            
            if result.matched:
                # 计算最低分数
                scores = []
                for interview_row in result.interview_rows:
                    score_value = interview_row.get('最低面试分数', interview_row.get('分数', 0))
                    try:
                        scores.append(float(score_value))
                    except (ValueError, TypeError):
                        continue
                
                min_score = min(scores) if scores else 0
                candidate_count = len(result.interview_rows)
                status = '成功'
            else:
                min_score = 0
                candidate_count = 0
                status = '未匹配'
            
            # 构建完整的数据行
            full_row_data = {}
            
            # 添加职位表的所有数据
            for key, value in pos_data.items():
                # 处理空值和NaN
                if value is None or (isinstance(value, float) and pd.isna(value)):
                    processed_value = ''
                else:
                    processed_value = str(value).strip()
                    
                # 确保职位代码以文本格式保存
                if key == '职位代码':
                    if processed_value and processed_value != 'nan':
                        if processed_value.replace('.', '').replace('E', '').replace('+', '').replace('-', '').isdigit():
                            processed_value = "'" + processed_value
                    full_row_data[key] = processed_value
                else:
                    full_row_data[key] = processed_value
            
            # 添加处理结果数据
            full_row_data['最低面试分数'] = min_score
            full_row_data['面试人数'] = candidate_count
            full_row_data['状态'] = status
            
            # 根据用户选择的列过滤数据
            if selected_columns:
                filtered_row = {}
                for col in selected_columns:
                    # 获取列数据，如果不存在则使用空字符串
                    value = full_row_data.get(col, '')
                    # 确保值不是None或NaN
                    if value is None or (isinstance(value, float) and pd.isna(value)):
                        value = ''
                    filtered_row[col] = value
                export_data.append(filtered_row)
            else:
                # 使用默认列
                default_columns = ['招考职位', '职位代码', '用人司局', '最低面试分数', '面试人数', '状态']
                filtered_row = {}
                for col in default_columns:
                    value = full_row_data.get(col, '')
                    if value is None or (isinstance(value, float) and pd.isna(value)):
                        value = ''
                    filtered_row[col] = value
                export_data.append(filtered_row)
        
        # 创建DataFrame并导出
        df = pd.DataFrame(export_data)
        
        # 确保职位代码列为文本类型
        df['职位代码'] = df['职位代码'].astype(str)
        
        # 使用openpyxl引擎进行更精确的格式控制
        with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
            df.to_excel(writer, sheet_name='岗位分数汇总', index=False)
            
            # 获取工作表对象，设置职位代码列为文本格式
            worksheet = writer.sheets['岗位分数汇总']
            
            # 设置职位代码列（B列）为文本格式
            from openpyxl.styles import NamedStyle
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
            
            # 添加统计信息
            stats = self.get_match_statistics(match_results)
            stats_df = pd.DataFrame([stats])
            stats_df.to_excel(writer, sheet_name='统计信息', index=False)
        
        export_time = time.time() - start_time
        self.logger.info(f"结果导出完成，耗时: {export_time:.2f}秒，文件: {output_path}")
    
    def clear_indices(self):
        """清理索引，释放内存"""
        self.interview_indices.clear()
        self.logger.info("索引已清理")


def create_fast_matcher(column_mappings: Dict[str, str]) -> FastDataMatcher:
    """
    创建快速数据匹配器
    
    Args:
        column_mappings: 列映射关系
        
    Returns:
        快速数据匹配器实例
    """
    return FastDataMatcher(column_mappings)