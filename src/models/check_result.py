"""
检查结果模型定义
"""
from dataclasses import dataclass, field
from typing import List, Optional
from enum import Enum


class CheckStatus(Enum):
    """检查状态"""
    PASSED = "通过"
    FAILED = "未通过"
    WARNING = "警告"
    ERROR = "错误"


@dataclass
class CheckIssue:
    """检查问题"""
    position: str  # 位置描述
    description: str  # 问题描述
    suggestion: Optional[str] = None  # 修改建议
    
    def __str__(self) -> str:
        result = f"位置: {self.position}\n问题: {self.description}"
        if self.suggestion:
            result += f"\n建议: {self.suggestion}"
        return result


@dataclass
class CheckResult:
    """单个规则的检查结果"""
    rule_id: str
    rule_name: str
    category: str  # 规则类别
    status: CheckStatus
    passed: bool
    issues: List[CheckIssue] = field(default_factory=list)
    message: str = ""

    def add_issue(self, issue: CheckIssue):
        """添加问题"""
        self.issues.append(issue)

    def get_issue_count(self) -> int:
        """获取问题数量"""
        return len(self.issues)

    def __str__(self) -> str:
        status_str = "✓" if self.passed else "✗"
        result = f"{status_str} {self.rule_name}: "
        if self.passed:
            result += "通过"
        else:
            result += f"发现 {self.get_issue_count()} 个问题"
        return result


@dataclass
class DocumentCheckResult:
    """文档检查结果"""
    document_path: str
    total_rules: int = 0
    passed_rules: int = 0
    failed_rules: int = 0
    warning_rules: int = 0
    results: List[CheckResult] = field(default_factory=list)
    
    def add_result(self, result: CheckResult):
        """添加检查结果"""
        self.results.append(result)
        self.total_rules += 1
        
        if result.passed:
            self.passed_rules += 1
        elif result.status == CheckStatus.WARNING:
            self.warning_rules += 1
        else:
            self.failed_rules += 1
    
    def get_pass_rate(self) -> float:
        """获取通过率"""
        if self.total_rules == 0:
            return 0.0
        return (self.passed_rules / self.total_rules) * 100
    
    def get_total_issues(self) -> int:
        """获取总问题数"""
        return sum(result.get_issue_count() for result in self.results)
    
    def get_failed_results(self) -> List[CheckResult]:
        """获取未通过的结果"""
        return [r for r in self.results if not r.passed]
    
    def __str__(self) -> str:
        return (
            f"检查结果: {self.passed_rules}/{self.total_rules} 通过 "
            f"({self.get_pass_rate():.1f}%)\n"
            f"发现问题: {self.get_total_issues()} 个"
        )
