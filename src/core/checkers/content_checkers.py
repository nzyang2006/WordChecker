"""
内容检查器
"""
import re
from typing import List
from core.base_checker import BaseChecker
from models.check_result import CheckIssue


class TableCaptionChecker(BaseChecker):
    """表格标题检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查表格是否有标题"""
        issues = []
        
        # 获取标题模式
        caption_pattern = self.get_param('caption_pattern', r'表\d+')
        
        # 获取所有表格
        tables = document_processor.get_tables()
        
        for table in tables:
            index = table.get('index', 0)
            
            # 检查表格是否有标题（通常在表格前一段）
            paragraphs = document_processor.get_paragraphs()
            table_found = False
            
            # 由于python-docx的API限制，我们通过段落数量估算表格位置
            # 这里使用简化逻辑：检查所有段落中是否有表格标题
            
        # 简化检查：查找文档中的表格标题
        has_table_caption = False
        paragraphs = document_processor.get_paragraphs()
        
        for para in paragraphs:
            text = para.get('text', '')
            if re.search(caption_pattern, text):
                has_table_caption = True
                break
        
        if tables and not has_table_caption:
            issues.append(self.create_issue(
                position="表格所在页",
                description=f"文档包含{len(tables)}个表格，但未找到符合规范的表格标题",
                suggestion=f"建议为每个表格添加标题，格式如: {caption_pattern}"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class ImageCaptionChecker(BaseChecker):
    """图片标题检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查图片是否有标题"""
        issues = []
        
        # 获取标题模式
        caption_pattern = self.get_param('caption_pattern', r'图\d+')
        
        # python-docx无法直接获取图片信息
        # 我们通过检查段落数据来查找图片标题
        
        paragraphs = document_processor.get_paragraphs()
        has_image_caption = False
        
        for para in paragraphs:
            text = para.get('text', '')
            if re.search(caption_pattern, text):
                has_image_caption = True
                break
        
        # 注意：这是一个简化检查，实际应用中可能需要更复杂的逻辑
        # 来确定文档是否真的包含图片
        
        # 暂时返回通过，实际应用中需要实现图片检测
        passed = True
        return self.create_result(passed, issues)


class ForbiddenWordsChecker(BaseChecker):
    """敏感词检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查敏感词"""
        issues = []
        
        # 获取敏感词列表
        forbidden_words = self.get_param('forbidden_words', [])
        case_sensitive = self.get_param('case_sensitive', False)
        
        if not forbidden_words:
            # 如果没有配置敏感词，直接通过
            return self.create_result(True)
        
        # 获取所有段落
        paragraphs = document_processor.get_paragraphs()
        
        found_words = []
        
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            
            if not case_sensitive:
                text_lower = text.lower()
            
            for word in forbidden_words:
                search_word = word if case_sensitive else word.lower()
                search_text = text if case_sensitive else text_lower
                
                if search_word in search_text:
                    # 找到敏感词的位置
                    start = search_text.find(search_word)
                    context = text[max(0, start-10):start+len(word)+10]
                    
                    found_words.append({
                        'position': f"第{i+1}段",
                        'word': word,
                        'context': context
                    })
        
        if found_words:
            # 按敏感词分组统计
            word_stats = {}
            for item in found_words:
                word = item['word']
                if word not in word_stats:
                    word_stats[word] = []
                word_stats[word].append(item['position'])
            
            for word, positions in word_stats.items():
                issues.append(self.create_issue(
                    position=f"第{positions[0] if positions else 'N/A'}页等多处",
                    description=f"发现敏感词: '{word}'（共{len(positions)}处）",
                    suggestion="建议删除或替换该词汇"
                ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)
