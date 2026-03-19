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
                     status: CheckStatus = CheckStatus.FAILED, message: str = "",
                     expected_value: str = None, actual_value: str = None) -> CheckResult:
        """
        创建检查结果
        
        Args:
            passed: 是否通过
            issues: 问题列表
            status: 检查状态
            message: 消息
            expected_value: 期望值
            actual_value: 实际值
            
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
            message=message,
            expected_value=expected_value,
            actual_value=actual_value
        )
    
    def create_not_applicable_result(self, reason: str = "") -> CheckResult:
        """
        创建不适用结果（当文档中无相关内容时）
        
        Args:
            reason: 不适用的原因
            
        Returns:
            CheckResult: 不适用的检查结果
        """
        return CheckResult(
            rule_id=self.rule.id,
            rule_name=self.rule.name,
            category=self.rule.category,
            status=CheckStatus.NOT_APPLICABLE,
            passed=False,
            issues=[],
            message=reason or "文档中无相关内容"
        )
    
    def create_issue(self, position: str, description: str, suggestion: str = None,
                    expected_value: str = None, actual_value: str = None) -> CheckIssue:
        """
        创建检查问题
        
        Args:
            position: 位置描述
            description: 问题描述
            suggestion: 修改建议
            expected_value: 期望值
            actual_value: 实际值
            
        Returns:
            CheckIssue: 检查问题对象
        """
        return CheckIssue(
            position=position,
            description=description,
            suggestion=suggestion,
            expected_value=expected_value,
            actual_value=actual_value
        )
    
    def get_param(self, key: str, default: Any = None) -> Any:
        """获取规则参数"""
        return self.rule.params.get(key, default)
