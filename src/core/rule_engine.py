"""
规则引擎
"""
from typing import List
from pathlib import Path
from models.rule import Rule
from models.check_result import DocumentCheckResult, CheckResult, CheckStatus
from utils.config_loader import ConfigLoader
from core.word_processor import WordProcessor
from core.checkers import CHECKER_MAP


class RuleEngine:
    """规则引擎"""
    
    def __init__(self, config_path: str = None):
        self.config_loader = ConfigLoader(config_path)
        self.rules: List[Rule] = []
        self._load_rules()
    
    def _load_rules(self):
        """加载规则"""
        self.rules = self.config_loader.get_enabled_rules()
    
    def reload_rules(self):
        """重新加载规则"""
        self._load_rules()
    
    def get_rules(self) -> List[Rule]:
        """获取所有规则"""
        return self.rules
    
    def get_rule_by_id(self, rule_id: str) -> Rule:
        """根据ID获取规则"""
        for rule in self.rules:
            if rule.id == rule_id:
                return rule
        return None
    
    def check_document(self, document_path: str, rule_ids: List[str] = None, 
                      progress_callback=None) -> DocumentCheckResult:
        """
        检查文档
        
        Args:
            document_path: 文档路径
            rule_ids: 要检查的规则ID列表，如果为None则检查所有规则
            progress_callback: 进度回调函数，格式：callback(current, total, rule_name)
            
        Returns:
            DocumentCheckResult: 文档检查结果
        """
        # 创建文档检查结果
        doc_result = DocumentCheckResult(document_path=document_path)
        
        # 创建Word处理器
        with WordProcessor(document_path) as processor:
            # 确定要检查的规则
            rules_to_check = self.rules
            if rule_ids:
                rules_to_check = [r for r in self.rules if r.id in rule_ids]
            
            # 过滤掉未启用的规则
            enabled_rules = [r for r in rules_to_check if r.enabled]
            total_rules = len(enabled_rules)
            
            # 执行检查
            for index, rule in enumerate(enabled_rules, start=1):
                # 调用进度回调
                if progress_callback:
                    progress_callback(index, total_rules, rule.name)
                
                try:
                    # 获取检查器类
                    checker_class = CHECKER_MAP.get(rule.checker)
                    if not checker_class:
                        result = CheckResult(
                            rule_id=rule.id,
                            rule_name=rule.name,
                            category=rule.category,
                            status=CheckStatus.ERROR,
                            passed=False,
                            message=f"未找到检查器: {rule.checker}"
                        )
                        doc_result.add_result(result)
                        continue

                    # 创建检查器实例并执行检查
                    checker = checker_class(rule)
                    result = checker.check(processor)
                    doc_result.add_result(result)

                except Exception as e:
                    # 检查过程中出错
                    from models.check_result import CheckIssue
                    error_issue = CheckIssue(
                        position="检查过程",
                        description=f"检查出错: {str(e)}",
                        suggestion="请联系技术支持或查看日志"
                    )
                    result = CheckResult(
                        rule_id=rule.id,
                        rule_name=rule.name,
                        category=rule.category,
                        status=CheckStatus.ERROR,
                        passed=False,
                        issues=[error_issue],
                        message=f"检查出错: {str(e)}"
                    )
                    doc_result.add_result(result)
        
        return doc_result
    
    def add_comments_to_document(self, document_path: str, doc_result: DocumentCheckResult, 
                                 save_path: str = None):
        """
        在文档中添加批注
        
        Args:
            document_path: 文档路径
            doc_result: 检查结果
            save_path: 保存路径，如果为None则覆盖原文件
        """
        with WordProcessor(document_path) as processor:
            # 为每个未通过的检查结果添加批注
            for result in doc_result.get_failed_results():
                for issue in result.issues:
                    try:
                        # 尝试解析位置信息，提取段落索引
                        # 假设位置格式为"第N段"
                        import re
                        match = re.search(r'第(\d+)段', issue.position)
                        if match:
                            paragraph_index = int(match.group(1)) - 1
                            
                            # 构建批注内容
                            comment_text = f"[{result.rule_name}]\n问题: {issue.description}"
                            if issue.suggestion:
                                comment_text += f"\n建议: {issue.suggestion}"
                            
                            # 添加批注
                            processor.add_comment_win32(paragraph_index, comment_text)
                    except Exception as e:
                        # 添加批注失败，记录日志
                        print(f"添加批注失败: {str(e)}")
            
            # 保存文档
            processor.save_document(save_path)
    
    def get_rule_statistics(self) -> dict:
        """获取规则统计信息"""
        total = len(self.rules)
        enabled = sum(1 for r in self.rules if r.enabled)
        
        categories = {}
        for rule in self.rules:
            cat = rule.category
            categories[cat] = categories.get(cat, 0) + 1
        
        return {
            'total': total,
            'enabled': enabled,
            'disabled': total - enabled,
            'categories': categories
        }
