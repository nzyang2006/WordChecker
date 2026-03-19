"""
格式检查器
"""
import re
from typing import List
from core.base_checker import BaseChecker
from models.check_result import CheckIssue


class TitleFormatChecker(BaseChecker):
    """标题格式检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查文档标题格式"""
        issues = []
        
        # 获取文档信息
        doc_info = document_processor.get_document_info()
        title = doc_info.get('title', '')
        
        # 获取第一个标题段落
        paragraphs = document_processor.get_paragraphs()
        first_heading = None
        for para in paragraphs:
            if para.get('style') and 'Heading' in para.get('style', ''):
                first_heading = para
                break
        
        # 检查标题长度
        max_length = self.get_param('max_length', 50)
        if title and len(title) > max_length:
            issues.append(self.create_issue(
                position="第1页 文档标题",
                description=f"标题长度超过限制（{len(title)} > {max_length}）",
                suggestion="建议缩短标题长度"
            ))
        
        # 检查必需关键词
        required_keywords = self.get_param('required_keywords', [])
        if required_keywords and title:
            missing_keywords = [kw for kw in required_keywords if kw not in title]
            if missing_keywords:
                issues.append(self.create_issue(
                    position="第1页 文档标题",
                    description=f"标题缺少关键词: {', '.join(missing_keywords)}",
                    suggestion="建议在标题中包含必需的关键词"
                ))
        
        # 检查是否有标题
        if not title and not first_heading:
            issues.append(self.create_issue(
                position="第1页 文档",
                description="文档没有设置标题",
                suggestion="建议设置文档标题"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class FontConsistencyChecker(BaseChecker):
    """字体一致性检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查字体是否一致"""
        issues = []
        
        # 获取允许的字体列表
        allowed_fonts = self.get_param('allowed_fonts', ['宋体', 'Times New Roman'])
        check_title = self.get_param('check_title', False)
        
        # 获取所有段落
        paragraphs = document_processor.get_paragraphs()
        
        # 统计字体使用情况
        font_usage = {}
        for para in paragraphs:
            # 跳过标题段落（如果配置要求）
            if not check_title and para.get('style') and 'Heading' in para.get('style', ''):
                continue
            
            for run in para.get('runs', []):
                font = run.get('font')
                if font:
                    font_usage[font] = font_usage.get(font, 0) + 1
        
        # 检查不允许的字体
        forbidden_fonts = set(font_usage.keys()) - set(allowed_fonts)
        if forbidden_fonts:
            for font in forbidden_fonts:
                issues.append(self.create_issue(
                    position="全文各页",
                    description=f"使用了不允许的字体: {font}（出现 {font_usage[font]} 次）",
                    suggestion=f"建议使用以下字体之一: {', '.join(allowed_fonts)}"
                ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class ParagraphSpacingChecker(BaseChecker):
    """段落间距检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查段落间距"""
        issues = []
        
        # 获取配置的间距值
        line_spacing = self.get_param('line_spacing', 1.5)
        before_spacing = self.get_param('before_spacing', 0)
        after_spacing = self.get_param('after_spacing', 0)
        
        # 获取所有段落
        paragraphs = document_processor.get_paragraphs()
        
        # 注意：python-docx 无法直接获取段落间距
        # 这里只做占位检查，实际应用中可能需要使用Win32 COM接口
        
        # 简单检查：段落是否为空
        empty_count = 0
        for i, para in enumerate(paragraphs):
            if not para.get('text', '').strip():
                empty_count += 1
        
        if empty_count > len(paragraphs) * 0.3:  # 空段落超过30%
            issues.append(self.create_issue(
                position="全文各页",
                description=f"文档中空段落较多（{empty_count}个）",
                suggestion="建议删除多余空行，调整段落间距"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class PageNumberChecker(BaseChecker):
    """页码检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查页码"""
        issues = []
        
        # python-docx 无法直接检查页码
        # 这里使用简单的启发式方法
        
        # 获取文档的段落数和表格数
        doc_info = document_processor.get_document_info()
        paragraph_count = doc_info.get('paragraph_count', 0)
        
        # 如果文档超过一定页数（估算），应该有页码
        # 假设每页约20个段落
        estimated_pages = paragraph_count / 20
        
        if estimated_pages > 3:
            # 检查文档中是否包含页码相关的文本
            paragraphs = document_processor.get_paragraphs()
            has_page_number = False
            
            for para in paragraphs:
                text = para.get('text', '').lower()
                if '第' in text and '页' in text:
                    has_page_number = True
                    break
            
            if not has_page_number:
                issues.append(self.create_issue(
                    position="页脚",
                    description="文档可能缺少页码",
                    suggestion="建议为文档添加页码"
                ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class DateFormatChecker(BaseChecker):
    """日期格式检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查日期格式"""
        issues = []
        
        # 获取允许的日期格式
        allowed_formats = self.get_param('allowed_formats', ['YYYY年MM月DD日', 'YYYY-MM-DD', 'YYYY/MM/DD'])
        
        # 定义日期正则表达式
        date_patterns = [
            r'\d{4}年\d{1,2}月\d{1,2}日',  # YYYY年MM月DD日
            r'\d{4}-\d{1,2}-\d{1,2}',      # YYYY-MM-DD
            r'\d{4}/\d{1,2}/\d{1,2}',      # YYYY/MM/DD
            r'\d{1,2}/\d{1,2}/\d{4}',      # MM/DD/YYYY
        ]
        
        # 获取所有段落
        paragraphs = document_processor.get_paragraphs()
        
        date_count = 0
        invalid_dates = []
        
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            
            for pattern in date_patterns:
                matches = re.findall(pattern, text)
                for match in matches:
                    date_count += 1
                    # 检查是否符合允许的格式
                    is_valid = False
                    for allowed_format in allowed_formats:
                        if '年' in allowed_format and '年' in match:
                            is_valid = True
                        elif '-' in allowed_format and '-' in match and '年' not in match:
                            is_valid = True
                        elif '/' in allowed_format and '/' in match and '年' not in match:
                            is_valid = True
                    
                    if not is_valid:
                        invalid_dates.append({
                            'position': f"第{i+1}段",
                            'date': match
                        })
        
        if invalid_dates:
            issues.append(self.create_issue(
                position="全文各页",
                description=f"发现{len(invalid_dates)}处日期格式不符合规范",
                suggestion=f"建议统一使用以下格式之一: {', '.join(allowed_formats)}"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class NumberFormatChecker(BaseChecker):
    """数字格式检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查数字格式"""
        issues = []
        
        # 获取配置
        thousands_separator = self.get_param('thousands_separator', True)
        decimal_places = self.get_param('decimal_places', 2)
        
        # 获取所有段落
        paragraphs = document_processor.get_paragraphs()
        
        # 查找大数字（超过1000）
        large_number_pattern = r'\b\d{4,}\b'
        number_count = 0
        improper_numbers = []
        
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            matches = re.findall(large_number_pattern, text)
            
            for match in matches:
                number_count += 1
                # 检查是否有千位分隔符
                if thousands_separator and ',' not in match:
                    improper_numbers.append({
                        'position': f"第{i+1}段",
                        'number': match
                    })
        
        if improper_numbers:
            issues.append(self.create_issue(
                position="全文各页",
                description=f"发现{len(improper_numbers)}个大数字缺少千位分隔符",
                suggestion="建议为超过1000的数字添加千位分隔符（如：1,000）"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)
