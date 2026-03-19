"""
结构检查器
"""
from typing import List
from core.base_checker import BaseChecker
from models.check_result import CheckIssue


class HeadingSequenceChecker(BaseChecker):
    """标题层级检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查标题层级是否连续"""
        issues = []
        
        # 获取所有标题
        headings = document_processor.get_headings()
        
        if not headings:
            return self.create_result(True, message="文档中没有标题")
        
        # 检查标题层级是否连续
        previous_level = 0
        heading_count = {}
        
        for i, heading in enumerate(headings):
            level = heading.get('level', 1)
            
            # 统计各级标题数量
            heading_count[level] = heading_count.get(level, 0) + 1
            
            # 检查层级跳跃（例如从H1直接跳到H3）
            if level - previous_level > 1 and previous_level > 0:
                issues.append(self.create_issue(
                    position=f"第{heading.get('estimated_page', 1)}页 第{heading.get('index', i)+1}段",
                    description=f"标题层级不连续：从H{previous_level}跳到H{level}",
                    suggestion=f"建议在H{previous_level}和H{level}之间添加H{previous_level+1}级标题"
                ))
            
            previous_level = level
        
        # 检查是否有顶级标题
        if 1 not in heading_count or heading_count[1] == 0:
            issues.append(self.create_issue(
                position="文档各页 文档结构",
                description="文档缺少一级标题",
                suggestion="建议为文档添加一级标题（H1）"
            ))

        # 检查一级标题是否过多
        if heading_count.get(1, 0) > 1:
            issues.append(self.create_issue(
                position="文档各页 文档结构",
                description=f"文档有多个一级标题（{heading_count[1]}个）",
                suggestion="建议一个文档只使用一个一级标题"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)
