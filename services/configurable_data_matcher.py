"""
可配置数据匹配器
支持用户自定义列映射的数据匹配功能
"""
import logging
from typing import Dict, List, Optional, Tuple, Set, Any
from difflib import SequenceMatcher
import re
from dataclasses import dataclass
import pandas as pd

from .data_validator import ValidationError


class ConfigurableDataMatchingError(Exception):
    """可配置数据匹配相关异常"""
    pass


@dataclass
class ColumnMapping:
    """列映射配置"""
    position_column: str  # 岗位表列名
    interview_column: str  # 面试表列名
    weight: float = 1.0  # 匹配权重
    match_type: str = "exact"  # 匹配类型: exact, fuzzy, contains


@dataclass
class ConfigurableMatchResult:
    """可配置匹配结果"""
    position_row: Dict[str, Any]
    interview_rows: List[Dict[str, Any]]
    match_score: float
    match_details: Dict[str, Any]
    matched: bool


class ConfigurableDataMatcher:
    """可配置数据匹配器类"""
    
    def __init__(self, column_mappings: Dict[str, str], fuzzy_threshold: float = 0.8):
        """
        初始化可配置数据匹配器
        
        Args:
            column_mappings: 列映射关系，格式为 {岗位表列名: 面试表列名}
            fuzzy_threshold: 模糊匹配阈值
        """
        self.logger = logging.getLogger(__name__)
        self.column_mappings = column_mappings
        self.fuzzy_threshold = fuzzy_threshold
        self.match_results: List[ConfigurableMatchResult] = []
        
        # 验证列映射配置
        self._validate_column_mappings()
        
        self.logger.info(f"初始化可配置数据匹配器，列映射: {column_mappings}")
    
    def _validate_column_mappings(self):
        """验证列映射配置"""
        if not self.column_mappings:
            raise ConfigurableDataMatchingError("列映射配置不能为空")
        
        for pos_col, int_col in self.column_mappings.items():
            if not pos_col or not int_col:
                raise ConfigurableDataMatchingError(f"列映射配置无效: '{pos_col}' -> '{int_col}'")
    
    def match_data(self, position_df: pd.DataFrame, interview_df: pd.DataFrame) -> Dict[str, Any]:
        """
        根据配置的列映射匹配数据
        
        Args:
            position_df: 岗位表DataFrame
            interview_df: 面试人员表DataFrame
            
        Returns:
            匹配结果字典
        """
        try:
            self.logger.info(f"开始匹配数据，岗位表: {len(position_df)} 行，面试表: {len(interview_df)} 行")
            
            # 记录详细的列映射配置
            self.logger.info("=== 列映射配置详情 ===")
            print("=== 列映射配置详情 ===")
            for pos_col, int_col in self.column_mappings.items():
                log_msg = f"岗位表列 '{pos_col}' -> 面试表列 '{int_col}'"
                self.logger.info(log_msg)
                print(log_msg)
            
            # 记录数据表的列信息
            self.logger.info("=== 岗位表列信息 ===")
            print("=== 岗位表列信息 ===")
            log_msg = f"岗位表列名: {list(position_df.columns)}"
            self.logger.info(log_msg)
            print(log_msg)
            
            self.logger.info("=== 面试表列信息 ===")
            print("=== 面试表列信息 ===")
            log_msg = f"面试表列名: {list(interview_df.columns)}"
            self.logger.info(log_msg)
            print(log_msg)
            
            # 验证列是否存在
            self._validate_columns_exist(position_df, interview_df)
            
            # 清空之前的结果
            self.match_results.clear()
            
            # 为每个岗位表行寻找匹配
            total_positions = len(position_df)
            self.logger.info(f"开始逐行匹配，总共 {total_positions} 个岗位")
            print(f"开始逐行匹配，总共 {total_positions} 个岗位")
            
            for pos_idx, pos_row in position_df.iterrows():
                # 记录进度
                if pos_idx % 10 == 0 or pos_idx == total_positions - 1:
                    progress = (pos_idx + 1) / total_positions * 100
                    progress_msg = f"匹配进度: {progress:.1f}% ({pos_idx + 1}/{total_positions})"
                    self.logger.info(progress_msg)
                    print(progress_msg)
                
                match_result = self._find_matches_for_position(pos_row, interview_df)
                self.match_results.append(match_result)
            
            # 生成匹配统计
            return self._generate_match_statistics()
            
        except Exception as e:
            if isinstance(e, ConfigurableDataMatchingError):
                raise
            else:
                raise ConfigurableDataMatchingError(f"数据匹配过程中发生错误: {str(e)}")
    
    def _validate_columns_exist(self, position_df: pd.DataFrame, interview_df: pd.DataFrame):
        """验证配置的列是否在DataFrame中存在"""
        # 检查岗位表列
        missing_pos_cols = []
        for pos_col in self.column_mappings.keys():
            if pos_col not in position_df.columns:
                missing_pos_cols.append(pos_col)
        
        if missing_pos_cols:
            raise ConfigurableDataMatchingError(f"岗位表中缺少以下列: {missing_pos_cols}")
        
        # 检查面试表列
        missing_int_cols = []
        for int_col in self.column_mappings.values():
            if int_col not in interview_df.columns:
                missing_int_cols.append(int_col)
        
        if missing_int_cols:
            raise ConfigurableDataMatchingError(f"面试人员表中缺少以下列: {missing_int_cols}")
    
    def _find_matches_for_position(self, pos_row: pd.Series, interview_df: pd.DataFrame) -> ConfigurableMatchResult:
        """
        为单个岗位行寻找匹配的面试记录（优化版本）
        
        Args:
            pos_row: 岗位表中的一行
            interview_df: 面试人员表DataFrame
            
        Returns:
            匹配结果
        """
        # 优化：首先尝试精确匹配，避免遍历所有行
        matched_rows = []
        match_score = 0.0
        match_details = {}
        
        # 构建匹配条件
        match_conditions = []
        for pos_col, int_col in self.column_mappings.items():
            pos_value = str(pos_row[pos_col]) if pd.notna(pos_row[pos_col]) else ""
            if pos_value:  # 只对非空值进行匹配
                match_conditions.append((int_col, pos_value))
        
        if not match_conditions:
            # 如果没有有效的匹配条件，返回未匹配
            self.logger.warning("没有有效的匹配条件，所有配置的列值都为空")
            return ConfigurableMatchResult(
                position_row=pos_row.to_dict(),
                interview_rows=[],
                match_score=0.0,
                match_details={'no_conditions': '所有配置的列值都为空'},
                matched=False
            )
        
        # 优化：使用pandas的向量化操作进行快速筛选
        try:
            # 使用所有配置的条件进行精确匹配
            mask = pd.Series([True] * len(interview_df))
            exact_match_details = {}
            
            # 获取岗位信息用于调试
            pos_name = pos_row.get('招考职位', 'N/A')
            pos_dept = pos_row.get('用人司局', 'N/A')
            pos_code = pos_row.get('职位代码', 'N/A')
            
            # 记录到日志文件
            self.logger.info(f"\n=== 开始匹配岗位 ===")
            self.logger.info(f"岗位名称: {pos_name}")
            self.logger.info(f"用人司局: {pos_dept}")
            self.logger.info(f"职位代码: {pos_code}")
            self.logger.info(f"配置的匹配条件数量: {len(match_conditions)}")
            self.logger.info(f"面试表总记录数: {len(interview_df)}")
            
            # 同时输出到控制台
            print(f"\n=== 开始匹配岗位 ===")
            print(f"岗位名称: {pos_name}")
            print(f"用人司局: {pos_dept}")
            print(f"职位代码: {pos_code}")
            print(f"配置的匹配条件数量: {len(match_conditions)}")
            print(f"面试表总记录数: {len(interview_df)}")
            
            # 对所有配置的列进行AND匹配
            for i, (int_col, pos_value) in enumerate(match_conditions):
                if int_col in interview_df.columns:
                    # 精确匹配
                    column_mask = interview_df[int_col].astype(str) == pos_value
                    matches_before = mask.sum()
                    mask = mask & column_mask
                    matches_after = mask.sum()
                    
                    exact_match_details[f"exact_{int_col}"] = {
                        'target_value': pos_value,
                        'match_type': 'exact',
                        'matches_found': column_mask.sum(),
                        'cumulative_matches': matches_after
                    }
                    
                    # 记录详细的匹配过程
                    log_msg = f"条件 {i+1}: {int_col} = '{pos_value}'"
                    self.logger.info(log_msg)
                    print(log_msg)
                    
                    log_msg = f"  该条件匹配的记录数: {column_mask.sum()}"
                    self.logger.info(log_msg)
                    print(log_msg)
                    
                    log_msg = f"  累积匹配记录数: {matches_before} -> {matches_after}"
                    self.logger.info(log_msg)
                    print(log_msg)
                    
                    # 如果该条件没有匹配到任何记录，显示面试表中该列的所有唯一值
                    if column_mask.sum() == 0:
                        unique_values = interview_df[int_col].dropna().astype(str).unique()[:10]  # 只显示前10个
                        log_msg = f"  面试表中 {int_col} 列的唯一值样例: {list(unique_values)}"
                        self.logger.warning(log_msg)
                        print(log_msg)
                    
                else:
                    log_msg = f"条件 {i+1}: {int_col} - 列不存在于面试表中"
                    self.logger.error(log_msg)
                    print(log_msg)
                    
                    # 显示面试表的所有列名
                    log_msg = f"  面试表的所有列名: {list(interview_df.columns)}"
                    self.logger.error(log_msg)
                    print(log_msg)
            
            exact_matches = interview_df[mask]
            log_msg = f"最终精确匹配结果: {len(exact_matches)} 个"
            self.logger.info(log_msg)
            print(log_msg)
            
            # 如果有匹配结果，显示前几个
            if len(exact_matches) > 0:
                log_msg = "匹配的面试人员样例:"
                self.logger.info(log_msg)
                print(log_msg)
                
                for idx, (_, row) in enumerate(exact_matches.head(3).iterrows()):
                    name = row.get('姓名', 'N/A')
                    score = row.get('最低面试分数', row.get('分数', 'N/A'))
                    dept = row.get('用人司局', 'N/A')
                    pos = row.get('招考职位', 'N/A')
                    
                    log_msg = f"  {idx+1}: {name} - 分数: {score} - 司局: {dept} - 岗位: {pos}"
                    self.logger.info(log_msg)
                    print(log_msg)
            else:
                log_msg = "没有找到匹配的面试人员"
                self.logger.warning(log_msg)
                print(log_msg)
                
                # 分析为什么没有匹配
                self.logger.warning("=== 匹配失败分析 ===")
                print("=== 匹配失败分析 ===")
                
                # 逐个条件分析
                for i, (int_col, pos_value) in enumerate(match_conditions):
                    if int_col in interview_df.columns:
                        # 查找最相似的值
                        column_values = interview_df[int_col].dropna().astype(str).unique()
                        similar_values = []
                        
                        for val in column_values:
                            if pos_value.lower() in val.lower() or val.lower() in pos_value.lower():
                                similar_values.append(val)
                        
                        if similar_values:
                            log_msg = f"条件 {i+1} ({int_col}='{pos_value}') 可能的相似值: {similar_values[:5]}"
                            self.logger.warning(log_msg)
                            print(log_msg)
                        else:
                            log_msg = f"条件 {i+1} ({int_col}='{pos_value}') 没有找到相似值"
                            self.logger.warning(log_msg)
                            print(log_msg)
            
            if len(exact_matches) > 0:
                # 找到精确匹配
                matched_rows = exact_matches.to_dict('records')
                match_score = 1.0
                match_details = exact_match_details
                self.logger.info(f"精确匹配成功，找到 {len(matched_rows)} 个匹配")
            else:
                # 不使用模糊匹配，只接受精确匹配
                self.logger.warning("精确匹配失败，不进行模糊匹配")
                matched_rows = []
                match_score = 0.0
                match_details = {
                    'no_match': '所有配置的条件都必须精确匹配',
                    'conditions_checked': len(match_conditions),
                    'match_conditions': match_conditions
                }
        
        except Exception as e:
            error_msg = f"匹配过程中发生错误: {e}"
            self.logger.error(error_msg)
            print(error_msg)
            # 发生错误时返回未匹配结果
            matched_rows = []
            match_score = 0.0
            match_details = {'error': str(e)}
        
        # 记录最终结果
        result_msg = f"岗位 '{pos_name}' 匹配完成: {'成功' if len(matched_rows) > 0 else '失败'}"
        if len(matched_rows) > 0:
            self.logger.info(result_msg)
        else:
            self.logger.warning(result_msg)
        print(result_msg)
        
        return ConfigurableMatchResult(
            position_row=pos_row.to_dict(),
            interview_rows=matched_rows,
            match_score=match_score,
            match_details=match_details,
            matched=len(matched_rows) > 0
        )
    
    def _calculate_match_score(self, pos_row: pd.Series, int_row: pd.Series) -> Tuple[float, Dict[str, Any]]:
        """
        计算两行数据的匹配分数
        
        Args:
            pos_row: 岗位表行
            int_row: 面试表行
            
        Returns:
            匹配分数和详细信息
        """
        total_score = 0.0
        total_weight = 0.0
        details = {}
        
        for pos_col, int_col in self.column_mappings.items():
            pos_value = str(pos_row[pos_col]) if pd.notna(pos_row[pos_col]) else ""
            int_value = str(int_row[int_col]) if pd.notna(int_row[int_col]) else ""
            
            # 计算单列匹配分数
            column_score = self._calculate_column_match_score(pos_value, int_value)
            
            # 权重（目前所有列权重相等）
            weight = 1.0
            
            total_score += column_score * weight
            total_weight += weight
            
            details[f"{pos_col}_vs_{int_col}"] = {
                'pos_value': pos_value,
                'int_value': int_value,
                'score': column_score,
                'weight': weight
            }
        
        # 计算加权平均分数
        final_score = total_score / total_weight if total_weight > 0 else 0.0
        
        return final_score, details
    
    def _calculate_column_match_score(self, value1: str, value2: str) -> float:
        """
        计算两个列值的匹配分数
        
        Args:
            value1: 值1
            value2: 值2
            
        Returns:
            匹配分数 (0.0 - 1.0)
        """
        if not value1 or not value2:
            return 0.0
        
        # 1. 精确匹配
        if value1 == value2:
            return 1.0
        
        # 2. 忽略大小写匹配
        if value1.lower() == value2.lower():
            return 0.95
        
        # 3. 去除空格后匹配
        clean1 = re.sub(r'\s+', '', value1)
        clean2 = re.sub(r'\s+', '', value2)
        if clean1 == clean2:
            return 0.9
        
        # 4. 包含关系匹配
        if value1 in value2 or value2 in value1:
            return 0.8
        
        # 5. 模糊匹配
        similarity = SequenceMatcher(None, value1, value2).ratio()
        
        # 6. 数字匹配（如果都是数字）
        if value1.isdigit() and value2.isdigit():
            if value1 == value2:
                return 1.0
            else:
                return 0.0  # 数字不匹配则完全不匹配
        
        return similarity
    
    def _generate_match_statistics(self) -> Dict[str, Any]:
        """生成匹配统计信息"""
        total_positions = len(self.match_results)
        matched_positions = len([r for r in self.match_results if r.matched])
        unmatched_positions = total_positions - matched_positions
        
        # 计算总的面试人员数量
        total_candidates = sum(len(r.interview_rows) for r in self.match_results if r.matched)
        
        # 按匹配分数分组
        high_confidence = len([r for r in self.match_results if r.matched and r.match_score >= 0.9])
        medium_confidence = len([r for r in self.match_results if r.matched and 0.7 <= r.match_score < 0.9])
        low_confidence = len([r for r in self.match_results if r.matched and r.match_score < 0.7])
        
        statistics = {
            'total_positions': total_positions,
            'matched_positions': matched_positions,
            'unmatched_positions': unmatched_positions,
            'match_rate': matched_positions / total_positions if total_positions > 0 else 0.0,
            'total_candidates': total_candidates,
            'high_confidence_matches': high_confidence,
            'medium_confidence_matches': medium_confidence,
            'low_confidence_matches': low_confidence,
            'average_match_score': sum(r.match_score for r in self.match_results if r.matched) / matched_positions if matched_positions > 0 else 0.0
        }
        
        # 详细的统计日志
        self.logger.info("=== 最终匹配统计 ===")
        print("=== 最终匹配统计 ===")
        
        stats_msg = f"总岗位数: {total_positions}"
        self.logger.info(stats_msg)
        print(stats_msg)
        
        stats_msg = f"匹配成功: {matched_positions} 个"
        self.logger.info(stats_msg)
        print(stats_msg)
        
        stats_msg = f"匹配失败: {unmatched_positions} 个"
        self.logger.info(stats_msg)
        print(stats_msg)
        
        stats_msg = f"匹配率: {statistics['match_rate']:.2%}"
        self.logger.info(stats_msg)
        print(stats_msg)
        
        stats_msg = f"匹配到的面试人员总数: {total_candidates} 人"
        self.logger.info(stats_msg)
        print(stats_msg)
        
        # 显示未匹配岗位的详细信息
        if unmatched_positions > 0:
            self.logger.warning("=== 未匹配岗位详情 ===")
            print("=== 未匹配岗位详情 ===")
            
            unmatched_count = 0
            for result in self.match_results:
                if not result.matched:
                    unmatched_count += 1
                    if unmatched_count <= 5:  # 只显示前5个未匹配的岗位
                        pos_data = result.position_row
                        pos_name = pos_data.get('招考职位', 'N/A')
                        pos_dept = pos_data.get('用人司局', 'N/A')
                        pos_code = pos_data.get('职位代码', 'N/A')
                        
                        unmatched_msg = f"未匹配岗位 {unmatched_count}: {pos_name} (司局: {pos_dept}, 代码: {pos_code})"
                        self.logger.warning(unmatched_msg)
                        print(unmatched_msg)
                        
                        # 显示匹配失败的原因
                        if 'match_conditions' in result.match_details:
                            conditions = result.match_details['match_conditions']
                            reason_msg = f"  匹配条件: {conditions}"
                            self.logger.warning(reason_msg)
                            print(reason_msg)
            
            if unmatched_count > 5:
                remaining_msg = f"  ... 还有 {unmatched_count - 5} 个未匹配岗位"
                self.logger.warning(remaining_msg)
                print(remaining_msg)
        
        # 显示匹配成功岗位的样例
        if matched_positions > 0:
            self.logger.info("=== 匹配成功岗位样例 ===")
            print("=== 匹配成功岗位样例 ===")
            
            matched_count = 0
            for result in self.match_results:
                if result.matched:
                    matched_count += 1
                    if matched_count <= 3:  # 只显示前3个匹配成功的岗位
                        pos_data = result.position_row
                        pos_name = pos_data.get('招考职位', 'N/A')
                        pos_dept = pos_data.get('用人司局', 'N/A')
                        pos_code = pos_data.get('职位代码', 'N/A')
                        candidate_count = len(result.interview_rows)
                        
                        matched_msg = f"匹配成功岗位 {matched_count}: {pos_name} (司局: {pos_dept}, 代码: {pos_code}, 面试人数: {candidate_count})"
                        self.logger.info(matched_msg)
                        print(matched_msg)
        
        return {
            'match_results': self.match_results,
            'statistics': statistics,
            'column_mappings': self.column_mappings
        }
    
    def get_matched_data_for_export(self) -> List[Dict[str, Any]]:
        """
        获取用于导出的匹配数据
        
        Returns:
            导出数据列表
        """
        export_data = []
        
        for result in self.match_results:
            if result.matched:
                pos_data = result.position_row
                
                # 计算最低分数
                scores = []
                for int_row in result.interview_rows:
                    score_value = int_row.get('最低面试分数', int_row.get('分数', int_row.get('score', 0)))
                    try:
                        scores.append(float(score_value))
                    except (ValueError, TypeError):
                        continue
                
                min_score = min(scores) if scores else 0.0
                candidate_count = len(result.interview_rows)
                
                export_row = {
                    '岗位名称': pos_data.get('招考职位', pos_data.get('岗位名称', '')),
                    '岗位代码': pos_data.get('职位代码', pos_data.get('岗位代码', '')),
                    '用人单位': pos_data.get('用人司局', pos_data.get('用人单位', '')),
                    '部门代码': pos_data.get('部门代码', ''),
                    '最低分数': min_score,
                    '面试人数': candidate_count,
                    '匹配分数': result.match_score,
                    '匹配状态': '已匹配'
                }
                
                export_data.append(export_row)
            else:
                # 未匹配的岗位
                pos_data = result.position_row
                export_row = {
                    '岗位名称': pos_data.get('招考职位', pos_data.get('岗位名称', '')),
                    '岗位代码': pos_data.get('职位代码', pos_data.get('岗位代码', '')),
                    '用人单位': pos_data.get('用人司局', pos_data.get('用人单位', '')),
                    '部门代码': pos_data.get('部门代码', ''),
                    '最低分数': 0,
                    '面试人数': 0,
                    '匹配分数': 0.0,
                    '匹配状态': '未匹配'
                }
                
                export_data.append(export_row)
        
        return export_data
    
    def get_unmatched_positions(self) -> List[Dict[str, Any]]:
        """获取未匹配的岗位"""
        return [r.position_row for r in self.match_results if not r.matched]
    
    def get_match_details(self) -> List[Dict[str, Any]]:
        """获取详细的匹配信息"""
        details = []
        
        for i, result in enumerate(self.match_results):
            detail = {
                'position_index': i,
                'position_data': result.position_row,
                'matched': result.matched,
                'match_score': result.match_score,
                'candidate_count': len(result.interview_rows),
                'match_details': result.match_details
            }
            
            if result.matched:
                detail['interview_data'] = result.interview_rows
            
            details.append(detail)
        
        return details
    
    def export_match_report(self, output_path: str) -> str:
        """
        导出匹配报告
        
        Args:
            output_path: 输出文件路径
            
        Returns:
            实际输出文件路径
        """
        try:
            export_data = self.get_matched_data_for_export()
            
            if not export_data:
                raise ConfigurableDataMatchingError("没有数据可以导出")
            
            # 创建DataFrame并导出
            df = pd.DataFrame(export_data)
            
            # 确保输出路径以.xlsx结尾
            if not output_path.endswith('.xlsx'):
                output_path += '.xlsx'
            
            with pd.ExcelWriter(output_path, engine='openpyxl') as writer:
                # 主要结果
                df.to_excel(writer, sheet_name='匹配结果', index=False)
                
                # 统计信息
                stats = self._generate_match_statistics()['statistics']
                stats_df = pd.DataFrame([stats])
                stats_df.to_excel(writer, sheet_name='统计信息', index=False)
                
                # 未匹配岗位
                unmatched = self.get_unmatched_positions()
                if unmatched:
                    unmatched_df = pd.DataFrame(unmatched)
                    unmatched_df.to_excel(writer, sheet_name='未匹配岗位', index=False)
                
                # 列映射配置
                mapping_data = [{'岗位表列': k, '面试表列': v} for k, v in self.column_mappings.items()]
                mapping_df = pd.DataFrame(mapping_data)
                mapping_df.to_excel(writer, sheet_name='列映射配置', index=False)
            
            self.logger.info(f"匹配报告已导出到: {output_path}")
            return output_path
            
        except Exception as e:
            raise ConfigurableDataMatchingError(f"导出匹配报告失败: {str(e)}")


def create_configurable_matcher(column_mappings: Dict[str, str], 
                               fuzzy_threshold: float = 0.8) -> ConfigurableDataMatcher:
    """
    创建可配置数据匹配器
    
    Args:
        column_mappings: 列映射关系
        fuzzy_threshold: 模糊匹配阈值
        
    Returns:
        可配置数据匹配器实例
    """
    return ConfigurableDataMatcher(column_mappings, fuzzy_threshold)