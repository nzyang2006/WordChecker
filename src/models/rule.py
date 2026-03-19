"""
规则模型定义
"""
from dataclasses import dataclass, field
from typing import Dict, Any, List, Optional
from enum import Enum


class RuleCategory(Enum):
    """规则类别"""
    FORMAT = "格式检查"
    CONTENT = "内容检查"
    STRUCTURE = "结构检查"


@dataclass
class Rule:
    """检查规则"""
    id: str
    name: str
    description: str
    category: str
    enabled: bool = True
    checker: str = ""
    params: Dict[str, Any] = field(default_factory=dict)
    
    @classmethod
    def from_dict(cls, data: dict) -> 'Rule':
        """从字典创建规则对象"""
        return cls(
            id=data.get('id', ''),
            name=data.get('name', ''),
            description=data.get('description', ''),
            category=data.get('category', ''),
            enabled=data.get('enabled', True),
            checker=data.get('checker', ''),
            params=data.get('params', {})
        )
    
    def to_dict(self) -> dict:
        """转换为字典"""
        return {
            'id': self.id,
            'name': self.name,
            'description': self.description,
            'category': self.category,
            'enabled': self.enabled,
            'checker': self.checker,
            'params': self.params
        }


@dataclass
class RuleGroup:
    """规则分组"""
    name: str
    category: RuleCategory
    rules: List[Rule] = field(default_factory=list)
    
    def add_rule(self, rule: Rule):
        """添加规则"""
        self.rules.append(rule)
    
    def get_enabled_rules(self) -> List[Rule]:
        """获取启用的规则"""
        return [rule for rule in self.rules if rule.enabled]
