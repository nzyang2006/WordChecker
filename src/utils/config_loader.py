"""
配置加载器
"""
import json
import sys
import os
from pathlib import Path
from typing import List, Dict, Any
from models.rule import Rule


def get_base_path():
    """
    获取基础路径（兼容开发环境和打包后的exe）
    
    开发环境：返回项目根目录
    打包后：返回exe所在目录
    """
    if getattr(sys, 'frozen', False):
        # 打包后的exe
        # sys.executable 是exe文件的完整路径
        return Path(sys.executable).parent
    else:
        # 开发环境
        # 从当前文件位置向上查找项目根目录
        current_dir = Path(__file__).parent  # src/utils
        project_root = current_dir.parent.parent  # 项目根目录
        
        # 验证是否为正确的项目根目录
        if (project_root / "config" / "rules.json").exists():
            return project_root
        else:
            # 如果找不到，尝试其他可能的路径
            # 当前工作目录
            cwd = Path.cwd()
            if (cwd / "config" / "rules.json").exists():
                return cwd
            # 返回默认路径
            return project_root


class ConfigLoader:
    """配置加载器"""

    def __init__(self, config_path: str = None):
        if config_path:
            self.config_path = config_path
        else:
            # 使用智能路径查找
            base_path = get_base_path()
            self.config_path = str(base_path / "config" / "rules.json")
        self._config = None

    def load_config(self) -> Dict[str, Any]:
        """加载配置文件"""
        if self._config is None:
            path = Path(self.config_path)
            if not path.exists():
                raise FileNotFoundError(f"配置文件不存在: {self.config_path}")

            with open(path, 'r', encoding='utf-8') as f:
                self._config = json.load(f)

        return self._config
    
    def load_rules(self) -> List[Rule]:
        """加载所有规则"""
        config = self.load_config()
        rules_data = config.get('rules', [])
        return [Rule.from_dict(data) for data in rules_data]
    
    def get_rule_by_id(self, rule_id: str) -> Rule:
        """根据ID获取规则"""
        rules = self.load_rules()
        for rule in rules:
            if rule.id == rule_id:
                return rule
        return None
    
    def get_rules_by_category(self, category: str) -> List[Rule]:
        """根据类别获取规则"""
        rules = self.load_rules()
        return [rule for rule in rules if rule.category == category]
    
    def get_enabled_rules(self) -> List[Rule]:
        """获取所有启用的规则"""
        rules = self.load_rules()
        return [rule for rule in rules if rule.enabled]
    
    def save_rules(self, rules: List[Rule]):
        """保存规则到配置文件"""
        config = self.load_config()
        config['rules'] = [rule.to_dict() for rule in rules]
        
        with open(self.config_path, 'w', encoding='utf-8') as f:
            json.dump(config, f, ensure_ascii=False, indent=2)
    
    def update_rule(self, rule: Rule):
        """更新单个规则"""
        rules = self.load_rules()
        for i, r in enumerate(rules):
            if r.id == rule.id:
                rules[i] = rule
                break
        self.save_rules(rules)
    
    def get_config_version(self) -> str:
        """获取配置版本"""
        config = self.load_config()
        return config.get('version', '1.0.0')
