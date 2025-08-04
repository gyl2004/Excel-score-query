"""
数据验证器
提供数据完整性检查和验证功能
"""
import logging
from typing import Dict, List, Optional, Tuple, Any
import pandas as pd
from dataclasses import dataclass


class ValidationError(Exception):
    """数据验证相关异常"""
    pass


class DataIntegrityError(Exception):
    """数据完整性相关异常"""
    pass


@dataclass
class ValidationResult:
    """验证结果数据类"""
    is_valid: bool
    errors: List[str]
    warnings: List[str]
    valid_count: int
    invalid_count: int
    details: Dict[str, Any]


@dataclass
class UnmatchedRecord:
    """无法匹配的记录数据类"""
    record_type: str  # "position" or "interview"
    record_data: Dict
    reason: str
    suggestions: List[str]


class DataValidator:
    """数据验证器类"""
    
    def __init__(self):
        """初始化数据验证器"""
        self.logger = logging.getLogger(__name__)
        self._validation_results: List[ValidationResult] = []
        self._unmatched_records: List[UnmatchedRecord] = []
        
    def validate_position_data_integrity(self, position_data: List[Dict]) -> ValidationResult:
        """
        验证职位数据的完整性
        
        Args:
            position_data: 职位数据列表
            
        Returns:
            ValidationResult: 验证结果
            
        Raises:
            ValidationError: 验证失败时抛出异常
        """
        errors = []
        warnings = []
        valid_count = 0
        invalid_count = 0
        details = {
            'duplicate_positions': [],
            'missing_codes': [],
            'empty_departments': [],
            'invalid_names': []
        }
        
        try:
            self.logger.info(f"开始验证职位数据完整性，共 {len(position_data)} 条记录")
            
            if not position_data:
                errors.append("职位数据为空")
                return ValidationResult(False, errors, warnings, 0, 0, details)
            
            # 检查重复职位
            position_names = []
            position_codes = []
            
            for i, position in enumerate(position_data):
                try:
                    # 验证基本结构
                    if not isinstance(position, dict):
                        errors.append(f"第 {i+1} 条记录不是有效的字典格式")
                        invalid_count += 1
                        continue
                    
                    # 验证必需字段
                    position_name = position.get('position_name', '').strip()
                    if not position_name:
                        errors.append(f"第 {i+1} 条记录缺少有效的岗位名称")
                        details['invalid_names'].append(i+1)
                        invalid_count += 1
                        continue
                    
                    # 检查岗位名称重复
                    if position_name in position_names:
                        warnings.append(f"发现重复的岗位名称: {position_name}")
                        details['duplicate_positions'].append(position_name)
                    else:
                        position_names.append(position_name)
                    
                    # 检查岗位代码
                    position_code = position.get('position_code', '').strip()
                    if not position_code:
                        warnings.append(f"岗位 '{position_name}' 缺少岗位代码")
                        details['missing_codes'].append(position_name)
                    elif position_code in position_codes:
                        warnings.append(f"发现重复的岗位代码: {position_code}")
                    else:
                        position_codes.append(position_code)
                    
                    # 检查部门信息
                    department = position.get('department', '').strip()
                    if not department:
                        warnings.append(f"岗位 '{position_name}' 缺少部门信息")
                        details['empty_departments'].append(position_name)
                    
                    valid_count += 1
                    
                except Exception as e:
                    errors.append(f"验证第 {i+1} 条记录时发生错误: {str(e)}")
                    invalid_count += 1
            
            is_valid = len(errors) == 0 and valid_count > 0
            
            self.logger.info(f"职位数据验证完成，有效记录: {valid_count}，无效记录: {invalid_count}，"
                           f"错误: {len(errors)}，警告: {len(warnings)}")
            
            result = ValidationResult(is_valid, errors, warnings, valid_count, invalid_count, details)
            self._validation_results.append(result)
            
            return result
            
        except Exception as e:
            raise ValidationError(f"验证职位数据完整性时发生错误: {str(e)}")
    
    def validate_interview_data_integrity(self, interview_data: List[Dict]) -> ValidationResult:
        """
        验证面试人员数据的完整性
        
        Args:
            interview_data: 面试人员数据列表
            
        Returns:
            ValidationResult: 验证结果
            
        Raises:
            ValidationError: 验证失败时抛出异常
        """
        errors = []
        warnings = []
        valid_count = 0
        invalid_count = 0
        details = {
            'duplicate_candidates': [],
            'invalid_scores': [],
            'missing_positions': [],
            'score_range_violations': []
        }
        
        try:
            self.logger.info(f"开始验证面试人员数据完整性，共 {len(interview_data)} 条记录")
            
            if not interview_data:
                errors.append("面试人员数据为空")
                return ValidationResult(False, errors, warnings, 0, 0, details)
            
            # 检查重复面试人员
            candidate_keys = []  # (姓名, 岗位) 组合
            
            for i, interview in enumerate(interview_data):
                try:
                    # 验证基本结构
                    if not isinstance(interview, dict):
                        errors.append(f"第 {i+1} 条记录不是有效的字典格式")
                        invalid_count += 1
                        continue
                    
                    # 验证必需字段
                    name = interview.get('name', '').strip()
                    if not name:
                        errors.append(f"第 {i+1} 条记录缺少有效的姓名")
                        invalid_count += 1
                        continue
                    
                    position_name = interview.get('position_name', '').strip()
                    if not position_name:
                        errors.append(f"第 {i+1} 条记录缺少有效的岗位名称")
                        details['missing_positions'].append(f"{name} (第{i+1}条)")
                        invalid_count += 1
                        continue
                    
                    # 检查重复面试人员
                    candidate_key = (name, position_name)
                    if candidate_key in candidate_keys:
                        warnings.append(f"发现重复的面试人员: {name} - {position_name}")
                        details['duplicate_candidates'].append(f"{name} - {position_name}")
                    else:
                        candidate_keys.append(candidate_key)
                    
                    # 验证分数
                    score = interview.get('score')
                    if score is None:
                        errors.append(f"第 {i+1} 条记录缺少分数信息")
                        details['invalid_scores'].append(f"{name} - {position_name}")
                        invalid_count += 1
                        continue
                    
                    try:
                        score_float = float(score)
                        if score_float < 0:
                            errors.append(f"第 {i+1} 条记录分数为负数: {score_float}")
                            details['score_range_violations'].append(f"{name} - {position_name}: {score_float}")
                            invalid_count += 1
                            continue
                    except (ValueError, TypeError):
                        errors.append(f"第 {i+1} 条记录分数格式无效: {score}")
                        details['invalid_scores'].append(f"{name} - {position_name}: {score}")
                        invalid_count += 1
                        continue
                    
                    valid_count += 1
                    
                except Exception as e:
                    errors.append(f"验证第 {i+1} 条记录时发生错误: {str(e)}")
                    invalid_count += 1
            
            is_valid = len(errors) == 0 and valid_count > 0
            
            self.logger.info(f"面试人员数据验证完成，有效记录: {valid_count}，无效记录: {invalid_count}，"
                           f"错误: {len(errors)}，警告: {len(warnings)}")
            
            result = ValidationResult(is_valid, errors, warnings, valid_count, invalid_count, details)
            self._validation_results.append(result)
            
            return result
            
        except Exception as e:
            raise ValidationError(f"验证面试人员数据完整性时发生错误: {str(e)}")
    
    def validate_match_results(self, match_results: Dict) -> ValidationResult:
        """
        验证匹配结果的质量
        
        Args:
            match_results: 匹配结果字典
            
        Returns:
            ValidationResult: 验证结果
            
        Raises:
            ValidationError: 验证失败时抛出异常
        """
        errors = []
        warnings = []
        valid_count = 0
        invalid_count = 0
        details = {
            'low_confidence_matches': [],
            'unmatched_positions': [],
            'unmatched_interviews': [],
            'match_statistics': {}
        }
        
        try:
            self.logger.info("开始验证匹配结果质量")
            
            # 验证匹配结果结构
            required_keys = ['mappings', 'unmatched_positions', 'unmatched_interviews', 'statistics']
            for key in required_keys:
                if key not in match_results:
                    errors.append(f"匹配结果缺少必需的键: {key}")
            
            if errors:
                return ValidationResult(False, errors, warnings, 0, 0, details)
            
            mappings = match_results['mappings']
            unmatched_positions = match_results['unmatched_positions']
            unmatched_interviews = match_results['unmatched_interviews']
            statistics = match_results['statistics']
            
            # 验证映射质量
            low_confidence_threshold = 0.7
            for mapping in mappings:
                if hasattr(mapping, 'match_confidence'):
                    if mapping.match_confidence < low_confidence_threshold:
                        warnings.append(f"低置信度匹配: {mapping.position_name} -> {mapping.interview_position} "
                                      f"(置信度: {mapping.match_confidence:.2f})")
                        details['low_confidence_matches'].append({
                            'position': mapping.position_name,
                            'interview_position': mapping.interview_position,
                            'confidence': mapping.match_confidence
                        })
                    valid_count += 1
            
            # 记录未匹配的记录
            for position in unmatched_positions:
                unmatched_record = UnmatchedRecord(
                    record_type="position",
                    record_data=position,
                    reason="无法在面试名单中找到匹配的岗位",
                    suggestions=[
                        "检查岗位名称是否正确",
                        "确认该岗位是否有面试人员",
                        "考虑使用模糊匹配或手动映射"
                    ]
                )
                self._unmatched_records.append(unmatched_record)
                details['unmatched_positions'].append(position.get('position_name', '未知岗位'))
                invalid_count += 1
            
            for interview_position in unmatched_interviews:
                unmatched_record = UnmatchedRecord(
                    record_type="interview",
                    record_data={'position_name': interview_position},
                    reason="无法在职位表中找到匹配的岗位",
                    suggestions=[
                        "检查面试岗位名称是否正确",
                        "确认该岗位是否在职位表中",
                        "考虑更新职位表或面试名单"
                    ]
                )
                self._unmatched_records.append(unmatched_record)
                details['unmatched_interviews'].append(interview_position)
                invalid_count += 1
            
            # 验证统计信息
            details['match_statistics'] = statistics
            match_rate = statistics.get('match_rate', 0)
            
            if match_rate < 0.5:
                warnings.append(f"匹配率较低: {match_rate:.1%}，建议检查数据质量")
            elif match_rate < 0.8:
                warnings.append(f"匹配率中等: {match_rate:.1%}，可能需要优化匹配算法")
            
            is_valid = len(errors) == 0
            
            self.logger.info(f"匹配结果验证完成，有效匹配: {valid_count}，未匹配: {invalid_count}，"
                           f"匹配率: {match_rate:.1%}")
            
            result = ValidationResult(is_valid, errors, warnings, valid_count, invalid_count, details)
            self._validation_results.append(result)
            
            return result
            
        except Exception as e:
            raise ValidationError(f"验证匹配结果时发生错误: {str(e)}")
    
    def generate_unmatched_report(self) -> Dict:
        """
        生成无法匹配记录的报告
        
        Returns:
            Dict: 无法匹配记录的详细报告
        """
        report = {
            'total_unmatched': len(self._unmatched_records),
            'unmatched_positions': [],
            'unmatched_interviews': [],
            'summary': {
                'position_count': 0,
                'interview_count': 0
            },
            'suggestions': {
                'data_quality': [
                    "检查岗位名称的一致性",
                    "确保职位表和面试名单使用相同的命名规范",
                    "移除多余的空格和特殊字符"
                ],
                'matching_improvement': [
                    "考虑降低模糊匹配阈值",
                    "添加更多的岗位代码映射",
                    "使用手动映射处理特殊情况"
                ]
            }
        }
        
        for record in self._unmatched_records:
            if record.record_type == "position":
                report['unmatched_positions'].append({
                    'position_name': record.record_data.get('position_name', ''),
                    'position_code': record.record_data.get('position_code', ''),
                    'department': record.record_data.get('department', ''),
                    'reason': record.reason,
                    'suggestions': record.suggestions
                })
                report['summary']['position_count'] += 1
            else:
                report['unmatched_interviews'].append({
                    'interview_position': record.record_data.get('position_name', ''),
                    'reason': record.reason,
                    'suggestions': record.suggestions
                })
                report['summary']['interview_count'] += 1
        
        self.logger.info(f"生成无法匹配记录报告，职位: {report['summary']['position_count']} 个，"
                        f"面试岗位: {report['summary']['interview_count']} 个")
        
        return report
    
    def validate_data_consistency(self, position_data: List[Dict], interview_data: List[Dict]) -> ValidationResult:
        """
        验证职位数据和面试数据的一致性
        
        Args:
            position_data: 职位数据列表
            interview_data: 面试人员数据列表
            
        Returns:
            ValidationResult: 验证结果
        """
        errors = []
        warnings = []
        valid_count = 0
        invalid_count = 0
        details = {
            'position_names': set(),
            'interview_positions': set(),
            'common_positions': set(),
            'position_only': set(),
            'interview_only': set()
        }
        
        try:
            self.logger.info("开始验证数据一致性")
            
            # 提取职位名称
            for position in position_data:
                position_name = position.get('position_name', '').strip()
                if position_name:
                    details['position_names'].add(position_name)
            
            # 提取面试岗位名称
            for interview in interview_data:
                position_name = interview.get('position_name', '').strip()
                if position_name:
                    details['interview_positions'].add(position_name)
            
            # 分析一致性
            details['common_positions'] = details['position_names'].intersection(details['interview_positions'])
            details['position_only'] = details['position_names'] - details['interview_positions']
            details['interview_only'] = details['interview_positions'] - details['position_names']
            
            # 生成报告
            total_positions = len(details['position_names'])
            total_interviews = len(details['interview_positions'])
            common_count = len(details['common_positions'])
            
            if common_count == 0:
                errors.append("职位表和面试名单没有任何匹配的岗位名称")
            else:
                valid_count = common_count
                
                consistency_rate = common_count / max(total_positions, total_interviews)
                if consistency_rate < 0.3:
                    warnings.append(f"数据一致性较低: {consistency_rate:.1%}")
                elif consistency_rate < 0.7:
                    warnings.append(f"数据一致性中等: {consistency_rate:.1%}")
            
            if details['position_only']:
                warnings.append(f"职位表中有 {len(details['position_only'])} 个岗位在面试名单中找不到")
                invalid_count += len(details['position_only'])
            
            if details['interview_only']:
                warnings.append(f"面试名单中有 {len(details['interview_only'])} 个岗位在职位表中找不到")
                invalid_count += len(details['interview_only'])
            
            is_valid = len(errors) == 0
            
            # 转换set为list以便序列化
            details['position_names'] = list(details['position_names'])
            details['interview_positions'] = list(details['interview_positions'])
            details['common_positions'] = list(details['common_positions'])
            details['position_only'] = list(details['position_only'])
            details['interview_only'] = list(details['interview_only'])
            
            self.logger.info(f"数据一致性验证完成，共同岗位: {common_count}，"
                           f"仅职位表: {len(details['position_only'])}，"
                           f"仅面试名单: {len(details['interview_only'])}")
            
            result = ValidationResult(is_valid, errors, warnings, valid_count, invalid_count, details)
            self._validation_results.append(result)
            
            return result
            
        except Exception as e:
            raise ValidationError(f"验证数据一致性时发生错误: {str(e)}")
    
    def get_validation_summary(self) -> Dict:
        """
        获取所有验证结果的汇总
        
        Returns:
            Dict: 验证结果汇总
        """
        summary = {
            'total_validations': len(self._validation_results),
            'passed_validations': 0,
            'failed_validations': 0,
            'total_errors': 0,
            'total_warnings': 0,
            'total_valid_records': 0,
            'total_invalid_records': 0,
            'validation_details': []
        }
        
        for result in self._validation_results:
            if result.is_valid:
                summary['passed_validations'] += 1
            else:
                summary['failed_validations'] += 1
            
            summary['total_errors'] += len(result.errors)
            summary['total_warnings'] += len(result.warnings)
            summary['total_valid_records'] += result.valid_count
            summary['total_invalid_records'] += result.invalid_count
            
            summary['validation_details'].append({
                'is_valid': result.is_valid,
                'error_count': len(result.errors),
                'warning_count': len(result.warnings),
                'valid_count': result.valid_count,
                'invalid_count': result.invalid_count
            })
        
        return summary
    
    def clear_validation_results(self):
        """清空验证结果"""
        self._validation_results.clear()
        self._unmatched_records.clear()
        self.logger.info("已清空验证结果")
    
    def export_validation_report(self, file_path: str) -> bool:
        """
        导出验证报告到文件
        
        Args:
            file_path: 导出文件路径
            
        Returns:
            bool: 导出是否成功
        """
        try:
            import json
            from datetime import datetime
            
            report = {
                'generated_at': datetime.now().isoformat(),
                'summary': self.get_validation_summary(),
                'unmatched_report': self.generate_unmatched_report(),
                'detailed_results': []
            }
            
            for result in self._validation_results:
                report['detailed_results'].append({
                    'is_valid': result.is_valid,
                    'errors': result.errors,
                    'warnings': result.warnings,
                    'valid_count': result.valid_count,
                    'invalid_count': result.invalid_count,
                    'details': result.details
                })
            
            with open(file_path, 'w', encoding='utf-8') as f:
                json.dump(report, f, ensure_ascii=False, indent=2)
            
            self.logger.info(f"验证报告已导出到: {file_path}")
            return True
            
        except Exception as e:
            self.logger.error(f"导出验证报告失败: {str(e)}")
            return False