"""
检查器基类
"""
from abc import ABC, abstractmethod
from typing import List, Dict, Any
from models.rule import Rule
from models.check_result import CheckResult, CheckIssue, CheckStatus


class BaseChecker(ABC):
    """检查器基类"""
    
    def __init__(self, rule: Rule):
        self.rule = rule
    
    @abstractmethod
    def check(self, document_processor) -> CheckResult:
        """
        执行检查
        
        Args:
            document_processor: WordProcessor实例
            
        Returns:
            CheckResult: 检查结果
        """
        pass
    
    def create_result(self, passed: bool, issues: List[CheckIssue] = None, 
                     status: CheckStatus = CheckStatus.FAILED, message: str = "") -> CheckResult:
        """
        创建检查结果
        
        Args:
            passed: 是否通过
            issues: 问题列表
            status: 检查状态
            message: 消息
            
        Returns:
            CheckResult: 检查结果对象
        """
        return CheckResult(
            rule_id=self.rule.id,
            rule_name=self.rule.name,
            category=self.rule.category,
            status=CheckStatus.PASSED if passed else status,
            passed=passed,
            issues=issues or [],
            message=message
        )
    
    def create_issue(self, position: str, description: str, suggestion: str = None) -> CheckIssue:
        """
        创建检查问题
        
        Args:
            position: 位置描述
            description: 问题描述
            suggestion: 修改建议
            
        Returns:
            CheckIssue: 检查问题对象
        """
        return CheckIssue(
            position=position,
            description=description,
            suggestion=suggestion
        )
    
    def get_param(self, key: str, default: Any = None) -> Any:
        """获取规则参数"""
        return self.rule.params.get(key, default)
