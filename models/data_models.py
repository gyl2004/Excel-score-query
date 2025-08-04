"""
数据模型定义
定义了系统中使用的核心数据结构
"""
from dataclasses import dataclass
from typing import Optional, List


@dataclass
class Position:
    """职位数据模型"""
    position_code: str
    position_name: str
    department: str
    sheet_name: str
    
    def __post_init__(self):
        """数据后处理，清理空白字符"""
        self.position_name = self.position_name.strip() if self.position_name else ""
        self.position_code = self.position_code.strip() if self.position_code else ""
        self.department = self.department.strip() if self.department else ""
        self.sheet_name = self.sheet_name.strip() if self.sheet_name else ""


@dataclass
class InterviewCandidate:
    """面试人员数据模型"""
    name: str
    position_name: str
    score: float
    is_qualified: bool
    
    def __post_init__(self):
        """数据后处理，验证分数范围"""
        self.name = self.name.strip() if self.name else ""
        self.position_name = self.position_name.strip() if self.position_name else ""
        
        if self.score < 0:
            raise ValueError(f"分数不能为负数: {self.score}")


@dataclass
class PositionScoreResult:
    """岗位分数结果数据模型"""
    position_code: str
    position_name: str
    department: str  # 用人司局
    department_name: str  # 部门名称
    recruit_count: int  # 招考人数
    candidate_count: int  # 面试人数
    min_score: Optional[float]
    status: str  # "正常", "无面试人员", "数据异常"
    notes: str
    all_scores: Optional[List[float]] = None  # 存储所有面试分数
    
    def __post_init__(self):
        """数据后处理，清理空白字符"""
        # 安全地处理position_code（可能是字符串或数字）
        if self.position_code:
            if isinstance(self.position_code, str):
                self.position_code = self.position_code.strip()
            else:
                self.position_code = str(self.position_code)
        else:
            self.position_code = ""
            
        self.position_name = self.position_name.strip() if self.position_name else ""
        self.department = self.department.strip() if self.department else ""
        self.department_name = self.department_name.strip() if self.department_name else ""
        self.status = self.status.strip() if self.status else ""
        self.notes = self.notes.strip() if self.notes else ""