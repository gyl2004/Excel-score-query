"""
数据匹配器
提供岗位数据匹配和验证功能
"""
import logging
from typing import Dict, List, Optional, Tuple, Set
from difflib import SequenceMatcher
import re
from dataclasses import dataclass


from .data_validator import ValidationError


class DataMatchingError(Exception):
    """数据匹配相关异常"""
    pass


@dataclass
class MatchResult:
    """匹配结果数据类"""
    position_name: str
    interview_position: str
    match_type: str  # "exact", "fuzzy", "code"
    confidence: float
    matched: bool


@dataclass
class PositionMapping:
    """岗位映射关系数据类"""
    position_code: str
    position_name: str
    department: str  # 用人司局
    department_name: str  # 部门名称
    recruit_count: int  # 招考人数
    sheet_name: str
    interview_position: str
    match_confidence: float
    candidate_count: int


class DataMatcher:
    """数据匹配器类"""
    
    def __init__(self, fuzzy_threshold: float = 0.8):
        """
        初始化数据匹配器
        
        Args:
            fuzzy_threshold: 模糊匹配的相似度阈值，默认0.8
        """
        self.logger = logging.getLogger(__name__)
        self.fuzzy_threshold = fuzzy_threshold
        self._position_mappings: List[PositionMapping] = []
        self._unmatched_positions: List[Dict] = []
        self._unmatched_interviews: List[str] = []
        
    def match_positions(self, position_data: List[Dict], interview_data: List[Dict]) -> Dict:
        """
        匹配职位表和面试人员名单表中的岗位信息
        
        Args:
            position_data: 职位数据列表
            interview_data: 面试人员数据列表
            
        Returns:
            Dict: 匹配结果，包含匹配成功的映射关系和统计信息
            
        Raises:
            DataMatchingError: 匹配过程中发生错误时抛出异常
        """
        try:
            self.logger.info(f"开始匹配岗位数据，职位表: {len(position_data)} 个岗位，面试名单: {len(interview_data)} 个面试人员")
            
            # 清空之前的匹配结果
            self._position_mappings.clear()
            self._unmatched_positions.clear()
            self._unmatched_interviews.clear()
            
            # 验证输入数据
            self._validate_input_data(position_data, interview_data)
            
            # 获取面试名单中的唯一岗位
            interview_positions = self._extract_interview_positions(interview_data)
            
            # 为每个职位表中的岗位寻找匹配
            for position in position_data:
                match_result = self._find_best_match(position, interview_positions, interview_data)
                
                if match_result.matched:
                    # 创建岗位映射
                    mapping = PositionMapping(
                        position_code=position.get('position_code', ''),
                        position_name=position['position_name'],
                        department=position.get('department', ''),
                        sheet_name=position.get('sheet_name', ''),
                        interview_position=match_result.interview_position,
                        match_confidence=match_result.confidence,
                        candidate_count=self._count_candidates(match_result.interview_position, interview_data)
                    )
                    self._position_mappings.append(mapping)
                    
                    self.logger.debug(f"匹配成功: {position['position_name']} -> {match_result.interview_position} "
                                    f"({match_result.match_type}, 置信度: {match_result.confidence:.2f})")
                else:
                    self._unmatched_positions.append(position)
                    self.logger.warning(f"无法匹配岗位: {position['position_name']}")
            
            # 查找面试名单中未匹配的岗位
            matched_interview_positions = {mapping.interview_position for mapping in self._position_mappings}
            self._unmatched_interviews = [pos for pos in interview_positions if pos not in matched_interview_positions]
            
            # 生成匹配结果统计
            result = self._generate_match_result()
            
            self.logger.info(f"岗位匹配完成，成功匹配: {len(self._position_mappings)} 个，"
                           f"未匹配职位: {len(self._unmatched_positions)} 个，"
                           f"未匹配面试岗位: {len(self._unmatched_interviews)} 个")
            
            return result
            
        except Exception as e:
            if isinstance(e, (DataMatchingError, ValidationError)):
                raise
            else:
                raise DataMatchingError(f"岗位匹配过程中发生错误: {str(e)}")
    
    def _validate_input_data(self, position_data: List[Dict], interview_data: List[Dict]) -> None:
        """
        验证输入数据的完整性
        
        Args:
            position_data: 职位数据列表
            interview_data: 面试人员数据列表
            
        Raises:
            ValidationError: 数据验证失败时抛出异常
        """
        if not position_data:
            raise ValidationError("职位数据不能为空")
        
        if not interview_data:
            raise ValidationError("面试人员数据不能为空")
        
        # 验证职位数据结构
        for i, position in enumerate(position_data):
            if not isinstance(position, dict):
                raise ValidationError(f"职位数据第 {i+1} 项不是字典类型")
            
            if 'position_name' not in position or not position['position_name']:
                raise ValidationError(f"职位数据第 {i+1} 项缺少有效的岗位名称")
        
        # 验证面试人员数据结构
        for i, interview in enumerate(interview_data):
            if not isinstance(interview, dict):
                raise ValidationError(f"面试人员数据第 {i+1} 项不是字典类型")
            
            required_fields = ['name', 'position_name', 'score']
            for field in required_fields:
                if field not in interview or interview[field] is None:
                    raise ValidationError(f"面试人员数据第 {i+1} 项缺少必需字段: {field}")
        
        self.logger.debug("输入数据验证通过")
    
    def _extract_interview_positions(self, interview_data: List[Dict]) -> List[str]:
        """
        从面试人员数据中提取唯一的岗位名称
        
        Args:
            interview_data: 面试人员数据列表
            
        Returns:
            List[str]: 唯一岗位名称列表
        """
        positions = set()
        for interview in interview_data:
            position_name = interview.get('position_name', '').strip()
            if position_name:
                positions.add(position_name)
        
        unique_positions = list(positions)
        self.logger.debug(f"从面试名单中提取到 {len(unique_positions)} 个唯一岗位")
        return unique_positions
    
    def _find_best_match(self, position: Dict, interview_positions: List[str], interview_data: List[Dict]) -> MatchResult:
        """
        为指定职位寻找最佳匹配
        
        Args:
            position: 职位信息
            interview_positions: 面试岗位列表
            interview_data: 面试人员数据（用于代码匹配）
            
        Returns:
            MatchResult: 匹配结果
        """
        position_name = position['position_name']
        position_code = position.get('position_code', '')
        
        # 1. 尝试精确匹配（岗位名称）
        exact_match = self._exact_match(position_name, interview_positions)
        if exact_match:
            return MatchResult(
                position_name=position_name,
                interview_position=exact_match,
                match_type="exact",
                confidence=1.0,
                matched=True
            )
        
        # 2. 尝试代码匹配（如果有岗位代码）
        if position_code:
            code_match = self._code_match(position_code, interview_positions)
            if code_match:
                return MatchResult(
                    position_name=position_name,
                    interview_position=code_match,
                    match_type="code",
                    confidence=0.95,
                    matched=True
                )
        
        # 3. 尝试模糊匹配
        fuzzy_match, confidence = self._fuzzy_match(position_name, interview_positions)
        if fuzzy_match and confidence >= self.fuzzy_threshold:
            return MatchResult(
                position_name=position_name,
                interview_position=fuzzy_match,
                match_type="fuzzy",
                confidence=confidence,
                matched=True
            )
        
        # 4. 无法匹配
        return MatchResult(
            position_name=position_name,
            interview_position="",
            match_type="none",
            confidence=0.0,
            matched=False
        )
    
    def _exact_match(self, position_name: str, interview_positions: List[str]) -> Optional[str]:
        """
        精确匹配岗位名称
        
        Args:
            position_name: 职位名称
            interview_positions: 面试岗位列表
            
        Returns:
            Optional[str]: 匹配的面试岗位名称，如果没有匹配则返回None
        """
        # 直接匹配
        if position_name in interview_positions:
            return position_name
        
        # 忽略大小写匹配
        position_lower = position_name.lower()
        for interview_pos in interview_positions:
            if interview_pos.lower() == position_lower:
                return interview_pos
        
        # 去除空格后匹配
        position_no_space = re.sub(r'\s+', '', position_name)
        for interview_pos in interview_positions:
            interview_no_space = re.sub(r'\s+', '', interview_pos)
            if position_no_space == interview_no_space:
                return interview_pos
        
        return None
    
    def _code_match(self, position_code: str, interview_positions: List[str]) -> Optional[str]:
        """
        基于岗位代码进行匹配
        
        Args:
            position_code: 岗位代码
            interview_positions: 面试岗位列表
            
        Returns:
            Optional[str]: 匹配的面试岗位名称，如果没有匹配则返回None
        """
        if not position_code:
            return None
        
        # 在面试岗位名称中查找包含岗位代码的项
        for interview_pos in interview_positions:
            if position_code in interview_pos:
                return interview_pos
        
        return None
    
    def _fuzzy_match(self, position_name: str, interview_positions: List[str]) -> Tuple[Optional[str], float]:
        """
        模糊匹配岗位名称
        
        Args:
            position_name: 职位名称
            interview_positions: 面试岗位列表
            
        Returns:
            Tuple[Optional[str], float]: 匹配的面试岗位名称和置信度
        """
        best_match = None
        best_ratio = 0.0
        
        # 预处理职位名称（去除常见的无关词汇）
        cleaned_position = self._clean_position_name(position_name)
        
        for interview_pos in interview_positions:
            cleaned_interview = self._clean_position_name(interview_pos)
            
            # 计算相似度
            ratio = SequenceMatcher(None, cleaned_position, cleaned_interview).ratio()
            
            # 如果包含关键词，提高相似度
            if self._contains_keywords(cleaned_position, cleaned_interview):
                ratio = min(1.0, ratio + 0.1)
            
            if ratio > best_ratio:
                best_ratio = ratio
                best_match = interview_pos
        
        return best_match, best_ratio
    
    def _clean_position_name(self, position_name: str) -> str:
        """
        清理岗位名称，去除无关词汇
        
        Args:
            position_name: 原始岗位名称
            
        Returns:
            str: 清理后的岗位名称
        """
        # 去除常见的无关词汇
        stop_words = ['岗位', '职位', '工作', '人员', '专员', '助理', '主管', '经理', '总监', '师']
        
        cleaned = position_name
        for word in stop_words:
            cleaned = cleaned.replace(word, '')
        
        # 去除多余空格和标点
        cleaned = re.sub(r'[^\w\u4e00-\u9fff]', '', cleaned)
        cleaned = re.sub(r'\s+', '', cleaned)
        
        return cleaned.strip()
    
    def _contains_keywords(self, position1: str, position2: str) -> bool:
        """
        检查两个岗位名称是否包含相同的关键词
        
        Args:
            position1: 岗位名称1
            position2: 岗位名称2
            
        Returns:
            bool: 是否包含相同关键词
        """
        # 定义常见的职位关键词
        common_keywords = ['软件', '工程', '开发', '产品', '经理', '设计', '前端', '后端', '数据', '分析', '测试', '运维']
        
        # 检查两个岗位名称中是否包含相同的关键词
        keywords1 = set()
        keywords2 = set()
        
        for keyword in common_keywords:
            if keyword in position1:
                keywords1.add(keyword)
            if keyword in position2:
                keywords2.add(keyword)
        
        # 如果有交集，说明包含相同关键词
        return len(keywords1.intersection(keywords2)) > 0
    
    def _count_candidates(self, interview_position: str, interview_data: List[Dict]) -> int:
        """
        统计指定岗位的面试人员数量
        
        Args:
            interview_position: 面试岗位名称
            interview_data: 面试人员数据列表
            
        Returns:
            int: 面试人员数量
        """
        count = 0
        for interview in interview_data:
            if interview.get('position_name') == interview_position:
                count += 1
        return count
    
    def _generate_match_result(self) -> Dict:
        """
        生成匹配结果统计
        
        Returns:
            Dict: 匹配结果统计信息
        """
        return {
            'mappings': self._position_mappings,
            'unmatched_positions': self._unmatched_positions,
            'unmatched_interviews': self._unmatched_interviews,
            'statistics': {
                'total_positions': len(self._position_mappings) + len(self._unmatched_positions),
                'matched_positions': len(self._position_mappings),
                'unmatched_positions': len(self._unmatched_positions),
                'unmatched_interviews': len(self._unmatched_interviews),
                'match_rate': len(self._position_mappings) / (len(self._position_mappings) + len(self._unmatched_positions)) if (len(self._position_mappings) + len(self._unmatched_positions)) > 0 else 0.0
            }
        }
    
    def create_position_mapping(self, matches: Dict) -> Dict[str, str]:
        """
        创建岗位映射关系字典
        
        Args:
            matches: 匹配结果
            
        Returns:
            Dict[str, str]: 职位名称到面试岗位名称的映射
        """
        mapping = {}
        
        for position_mapping in matches.get('mappings', []):
            mapping[position_mapping.position_name] = position_mapping.interview_position
        
        self.logger.info(f"创建岗位映射关系完成，包含 {len(mapping)} 个映射")
        return mapping
    
    def get_match_statistics(self) -> Dict:
        """
        获取匹配统计信息
        
        Returns:
            Dict: 匹配统计信息
        """
        total_positions = len(self._position_mappings) + len(self._unmatched_positions)
        
        return {
            'total_positions': total_positions,
            'matched_positions': len(self._position_mappings),
            'unmatched_positions': len(self._unmatched_positions),
            'unmatched_interviews': len(self._unmatched_interviews),
            'match_rate': len(self._position_mappings) / total_positions if total_positions > 0 else 0.0,
            'exact_matches': len([m for m in self._position_mappings if m.match_confidence == 1.0]),
            'fuzzy_matches': len([m for m in self._position_mappings if m.match_confidence < 1.0])
        }
    
    def get_unmatched_positions(self) -> List[Dict]:
        """
        获取未匹配的职位列表
        
        Returns:
            List[Dict]: 未匹配的职位信息列表
        """
        return self._unmatched_positions.copy()
    
    def get_unmatched_interviews(self) -> List[str]:
        """
        获取未匹配的面试岗位列表
        
        Returns:
            List[str]: 未匹配的面试岗位名称列表
        """
        return self._unmatched_interviews.copy()
    
    def get_position_mappings(self) -> List[PositionMapping]:
        """
        获取所有岗位映射关系
        
        Returns:
            List[PositionMapping]: 岗位映射关系列表
        """
        return self._position_mappings.copy()
    
    def validate_match_quality(self, min_confidence: float = 0.7) -> Dict:
        """
        验证匹配质量
        
        Args:
            min_confidence: 最低置信度阈值
            
        Returns:
            Dict: 匹配质量报告
        """
        high_quality_matches = []
        low_quality_matches = []
        
        for mapping in self._position_mappings:
            if mapping.match_confidence >= min_confidence:
                high_quality_matches.append(mapping)
            else:
                low_quality_matches.append(mapping)
        
        return {
            'high_quality_matches': high_quality_matches,
            'low_quality_matches': low_quality_matches,
            'high_quality_count': len(high_quality_matches),
            'low_quality_count': len(low_quality_matches),
            'quality_rate': len(high_quality_matches) / len(self._position_mappings) if self._position_mappings else 0.0
        }