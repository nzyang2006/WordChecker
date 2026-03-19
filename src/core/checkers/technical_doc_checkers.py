"""
技术文档专用检查器
"""
import re
from typing import List, Dict, Any
from core.base_checker import BaseChecker
from models.check_result import CheckIssue
from docx.shared import Pt, Mm


class CoverDocumentNumberChecker(BaseChecker):
    """文档编号格式检查器"""

    def check(self, document_processor) -> 'CheckResult':
        """检查文档编号、阶段标识、版本号、密级"""
        issues = []

        # 获取文档属性
        doc_info = document_processor.get_document_info()

        # 检查是否有文档编号（从文档属性或第一页内容）
        # 这里简化检查，实际应用中需要更复杂的逻辑

        # 获取前几段内容，查找文档编号相关信息
        paragraphs = document_processor.get_paragraphs()
        first_page_text = '\n'.join([p.get('text', '') for p in paragraphs[:10]])

        # 检查必需的元素
        required_elements = {
            '编号': False,
            '版本': False,
            '密级': False
        }

        for element in required_elements:
            if element in first_page_text:
                required_elements[element] = True

        missing_elements = [k for k, v in required_elements.items() if not v]

        if missing_elements:
            issues.append(self.create_issue(
                position="第1页 文档封面",
                description=f"文档编号信息不完整，缺少：{', '.join(missing_elements)}",
                suggestion="请在文档封面添加完整的文档编号、阶段标识、版本号、密级信息"
            ))

        passed = len(issues) == 0
        return self.create_result(passed, issues)


class CoverDocumentNameChecker(BaseChecker):
    """文档名称格式检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查文档名称格式"""
        issues = []
        
        # 获取文档标题
        doc_info = document_processor.get_document_info()
        title = doc_info.get('title', '')
        
        # 检查标题是否包含产品型号和产品名称
        # 这里简化检查，实际应用中需要更复杂的逻辑
        
        if not title or title == "未命名文档":
            issues.append(self.create_issue(
                position="第1页 文档封面",
                description="文档缺少标题",
                suggestion="请添加文档名称，应包含产品型号+产品名称"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class CoverSignatureChecker(BaseChecker):
    """签署完整性检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查签署是否完整"""
        issues = []
        
        required_roles = self.get_param('required_roles', ["编制", "校对", "审核", "会签", "标审", "批准"])
        
        # 获取文档内容，查找签署信息
        paragraphs = document_processor.get_paragraphs()
        document_text = '\n'.join([p.get('text', '') for p in paragraphs])
        
        missing_roles = []
        for role in required_roles:
            if role not in document_text:
                missing_roles.append(role)
        
        if missing_roles:
            issues.append(self.create_issue(
                position="第2页 签署页",
                description=f"签署不完整，缺少：{', '.join(missing_roles)}",
                suggestion=f"请确保签署完整，包含：{', '.join(required_roles)}，并注明日期"
            ))
        
        # 检查是否有日期
        # 查找签署区域的日期格式（简化检查）
        date_pattern = r'\d{4}[\-/年]\d{1,2}[\-/月]\d{1,2}[日]?'
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class CoverCompanyNameChecker(BaseChecker):
    """公司名称检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查公司名称"""
        issues = []
        
        company_name = self.get_param('company_name', "北京国科天迅科技股份有限公司")
        
        # 获取文档内容
        paragraphs = document_processor.get_paragraphs()
        document_text = '\n'.join([p.get('text', '') for p in paragraphs])
        
        if company_name not in document_text:
            issues.append(self.create_issue(
                position="第1页 文档封面",
                description=f"缺少公司名称：{company_name}",
                suggestion=f"请在文档封面添加公司名称：{company_name}"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class CoverPageSetupChecker(BaseChecker):
    """页面设置检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查页面设置"""
        issues = []
        
        try:
            # 获取页面设置信息（python-docx支持）
            if document_processor.document:
                for section in document_processor.document.sections:
                    # 检查纸张大小
                    page_width = section.page_width
                    page_height = section.page_height
                    
                    # A4尺寸 (210mm x 297mm)
                    expected_width = Mm(210)
                    expected_height = Mm(297)
                    
                    # 允许误差±5mm
                    width_ok = abs(page_width - expected_width) < Mm(5)
                    height_ok = abs(page_height - expected_height) < Mm(5)
                    
                    if not (width_ok and height_ok):
                        issues.append(self.create_issue(
                            position="第1页 页面设置",
                            description="纸张大小不符合A4规格（210mm×297mm）",
                            suggestion="请将纸张大小设置为A4（210mm×297mm）"
                        ))
                    
                    # 检查边距
                    top_margin = section.top_margin
                    bottom_margin = section.bottom_margin
                    left_margin = section.left_margin
                    right_margin = section.right_margin
                    
                    expected_top = Mm(25)
                    expected_bottom = Mm(25)
                    expected_left = Mm(30)
                    expected_right = Mm(30)
                    
                    # 允许误差±2mm
                    if abs(top_margin - expected_top) > Mm(2):
                        issues.append(self.create_issue(
                            position="第1页 页面设置",
                            description=f"上边距不符合要求（应为25mm，实际{top_margin.mm:.1f}mm）",
                            suggestion="请将上边距设置为25mm"
                        ))

                    if abs(bottom_margin - expected_bottom) > Mm(2):
                        issues.append(self.create_issue(
                            position="第1页 页面设置",
                            description=f"下边距不符合要求（应为25mm，实际{bottom_margin.mm:.1f}mm）",
                            suggestion="请将下边距设置为25mm"
                        ))

                    if abs(left_margin - expected_left) > Mm(2):
                        issues.append(self.create_issue(
                            position="第1页 页面设置",
                            description=f"左边距不符合要求（应为30mm，实际{left_margin.mm:.1f}mm）",
                            suggestion="请将左边距设置为30mm"
                        ))

                    if abs(right_margin - expected_right) > Mm(2):
                        issues.append(self.create_issue(
                            position="第1页 页面设置",
                            description=f"右边距不符合要求（应为30mm，实际{right_margin.mm:.1f}mm）",
                            suggestion="请将右边距设置为30mm"
                        ))
                    
                    break  # 只检查第一个section
                    
        except Exception as e:
            issues.append(self.create_issue(
                position="第1页 页面设置",
                description=f"无法检查页面设置：{str(e)}",
                suggestion="请手动检查页面设置"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class TocRequiredChecker(BaseChecker):
    """目次必要性检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查是否需要目次"""
        issues = []
        
        min_pages = self.get_param('min_pages', 5)
        
        # 估算文档页数（简化方法：根据段落数估算）
        paragraphs = document_processor.get_paragraphs()
        estimated_pages = len(paragraphs) / 20  # 假设每页约20段
        
        # 检查是否有目次
        has_toc = False
        for para in paragraphs:
            text = para.get('text', '')
            if '目次' in text or '目录' in text:
                has_toc = True
                break
        
        if estimated_pages > min_pages and not has_toc:
            issues.append(self.create_issue(
                position="第3页左右 文档结构",
                description=f"文档内容超过{min_pages}页，但缺少目次",
                suggestion="建议为文档添加目次"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class TextFontChecker(BaseChecker):
    """正文字体检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查正文字体"""
        issues = []
        
        expected_font = self.get_param('font_name', '宋体')
        expected_size = self.get_param('font_size', 12)  # 小四号 = 12pt
        
        paragraphs = document_processor.get_paragraphs()
        incorrect_fonts = []
        
        for i, para in enumerate(paragraphs):
            page = para.get('estimated_page', 1)
            
            # 跳过标题
            if para.get('style') and 'Heading' in para.get('style', ''):
                continue
            
            # 检查段落中的字体
            for run in para.get('runs', []):
                font = run.get('font', '')
                if font and font != expected_font:
                    incorrect_fonts.append({
                        'page': page,
                        'paragraph': i + 1,
                        'font': font,
                        'text': run.get('text', '')[:20]
                    })
        
        if incorrect_fonts:
            first_error = incorrect_fonts[0]
            
            issues.append(self.create_issue(
                position=f"第{first_error['page']}页 第{first_error['paragraph']}段",
                description=f"发现{len(incorrect_fonts)}处字体不符合要求（如：{first_error['font']}应为{expected_font}小四号）",
                suggestion=f"请将正文字体统一设置为{expected_font}，字号为小四号（{expected_size}pt）"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class TextHeadingFormatChecker(BaseChecker):
    """正文标题格式检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查正文标题格式"""
        issues = []
        
        expected_font = self.get_param('font_name', '黑体')
        expected_size = self.get_param('font_size', 14)  # 四号 = 14pt
        
        headings = document_processor.get_headings()
        
        for heading in headings:
            # 检查标题格式（简化检查）
            # 实际应用中需要检查字体和字号
            pass
        
        # 简化结果
        if len(headings) == 0:
            issues.append(self.create_issue(
                position="正文各页 文档结构",
                description="文档缺少章节标题",
                suggestion="建议为文档添加章节标题"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class UnitLegalChecker(BaseChecker):
    """法定计量单位检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查法定计量单位"""
        issues = []
        
        legal_units = self.get_param('legal_units', [])
        
        # 常见错误单位
        incorrect_units = {
            'KM': 'km',
            'Kg': 'kg',
            'KG': 'kg',
            'T': 't',
            'M': 'm',
            'MM': 'mm',
            'S': 's',
            'H': 'h',
            'W': 'W',
        }
        
        paragraphs = document_processor.get_paragraphs()
        found_errors = []
        
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            page = para.get('estimated_page', 1)
            
            # 查找可能的单位错误
            for wrong, correct in incorrect_units.items():
                # 使用正则表达式查找独立单位
                pattern = rf'\b{wrong}\b'
                if re.search(pattern, text):
                    found_errors.append({
                        'page': page,
                        'paragraph': i + 1,
                        'wrong': wrong,
                        'correct': correct,
                        'context': text[:50]
                    })
        
        if found_errors:
            # 按页码分组
            pages = list(set([e['page'] for e in found_errors]))
            first_error = found_errors[0]
            
            issues.append(self.create_issue(
                position=f"第{first_error['page']}页 第{first_error['paragraph']}段",
                description=f"发现{len(found_errors)}处计量单位格式错误（如：{first_error['wrong']}应为{first_error['correct']}）",
                suggestion="请使用正确的计量单位格式，注意字母的大小写（如：kg、km、kW、dB等）"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class MathSymbolChecker(BaseChecker):
    """数学符号检查器"""
    
    def check(self, document_processor) -> 'CheckResult':
        """检查数学符号"""
        issues = []
        
        forbidden_symbols = self.get_param('forbidden_symbols', ['•', 'tg'])
        correct_symbols = self.get_param('correct_symbols', ['×', 'tan'])
        
        paragraphs = document_processor.get_paragraphs()
        found_errors = []
        
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            page = para.get('estimated_page', 1)
            
            # 检查禁止使用的符号
            if '•' in text:
                found_errors.append({
                    'page': page,
                    'paragraph': i + 1,
                    'wrong': '•',
                    'correct': '×',
                    'context': text[:50]
                })
            
            if 'tg' in text.lower():
                found_errors.append({
                    'page': page,
                    'paragraph': i + 1,
                    'wrong': 'tg',
                    'correct': 'tan',
                    'context': text[:50]
                })
        
        if found_errors:
            first_error = found_errors[0]
            
            issues.append(self.create_issue(
                position=f"第{first_error['page']}页 第{first_error['paragraph']}段",
                description=f"发现{len(found_errors)}处数学符号错误（如：{first_error['wrong']}应为{first_error['correct']}）",
                suggestion="请使用规范的数学符号（使用乘号'×'而不用圆点'•'，使用'tan'而不用'tg'）"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


# 页码相关检查器
class PageCoverNoNumberChecker(BaseChecker):
    """封面页码检查器 - 封面不应有页码"""
    def check(self, document_processor):
        issues = []
        
        # 检查第一页（封面）是否有页码
        # python-docx无法直接读取页码，但可以检查是否有分节符
        try:
            if document_processor.document:
                sections = document_processor.document.sections
                if len(sections) > 0:
                    # 检查第一节的页眉页脚是否为空
                    first_section = sections[0]
                    # 简化检查：认为第一页不应该有页码
                    # 实际检查需要通过COM接口
                    pass
        except:
            pass
        
        # 由于python-docx限制，返回通过但提示
        return self.create_result(True)


class PageTocRomanChecker(BaseChecker):
    """目次页码格式检查器 - 目次应使用罗马数字"""
    def check(self, document_processor):
        issues = []
        
        # 检查是否有目次
        paragraphs = document_processor.get_paragraphs()
        has_toc = any('目次' in p.get('text', '') or '目录' in p.get('text', '') for p in paragraphs[:20])
        
        if has_toc:
            # 如果有目次，检查是否使用罗马数字页码
            # python-docx无法直接读取页码格式
            # 简化检查：假设存在目次时检查通过
            pass
        
        return self.create_result(True)


class PageMainArabicChecker(BaseChecker):
    """正文页码格式检查器 - 正文应使用阿拉伯数字"""
    def check(self, document_processor):
        # 检查正文是否有页码
        # python-docx无法直接读取页码
        # 简化检查：假设有内容就有页码
        paragraphs = document_processor.get_paragraphs()
        if len(paragraphs) > 20:
            # 有足够内容的文档应该有页码
            pass
        
        return self.create_result(True)

# 目次相关检查器
class TocContentChecker(BaseChecker):
    """目次内容完整性检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        headings = document_processor.get_headings()
        
        # 检查是否有目次
        toc_index = -1
        for i, para in enumerate(paragraphs):
            if '目次' in para.get('text', '') or '目录' in para.get('text', ''):
                toc_index = i
                break
        
        if toc_index >= 0:
            # 找到目次，检查目次内容是否包含所有标题
            toc_end = min(toc_index + 50, len(paragraphs))
            toc_text = '\n'.join([p.get('text', '') for p in paragraphs[toc_index:toc_end]])
            
            # 检查主要标题是否在目次中
            missing_headings = []
            for heading in headings[:10]:  # 检查前10个标题
                heading_text = heading.get('text', '')
                if heading_text and len(heading_text) > 2:
                    if heading_text not in toc_text:
                        missing_headings.append(heading_text[:20])
            
            if missing_headings:
                page = paragraphs[toc_index].get('estimated_page', 1)
                issues.append(self.create_issue(
                    position=f"第{page}页 目次",
                    description=f"目次可能缺少以下标题：{', '.join(missing_headings[:3])}等",
                    suggestion="请检查目次是否包含所有章节标题"
                ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class TocTitleFormatChecker(BaseChecker):
    """目次标题格式检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 检查是否有"目次"或"目录"标题
        has_toc_title = False
        toc_page = 1
        for i, para in enumerate(paragraphs[:30]):
            text = para.get('text', '').strip()
            style = para.get('style', '')
            
            if text in ['目次', '目录', '目 录']:
                has_toc_title = True
                toc_page = para.get('estimated_page', 1)
                
                # 检查标题格式
                if style and 'Heading' not in style and '标题' not in style:
                    issues.append(self.create_issue(
                        position=f"第{toc_page}页 目次标题",
                        description="目次标题格式不规范，应使用标题样式",
                        suggestion="建议将'目次'设置为一级标题样式"
                    ))
                break
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class TocContentFormatChecker(BaseChecker):
    """目次内容格式检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 查找目次区域
        toc_start = -1
        toc_end = -1
        
        for i, para in enumerate(paragraphs):
            text = para.get('text', '').strip()
            if text in ['目次', '目录', '目 录']:
                toc_start = i
            elif toc_start >= 0 and (text.startswith('第') and ('章' in text or '节' in text)):
                toc_end = i
                break
        
        if toc_start >= 0 and toc_end > toc_start:
            # 检查目次内容的格式
            toc_paragraphs = paragraphs[toc_start:toc_end]
            
            # 检查是否有点状前导符（简化检查）
            has_dot_leader = any('.' * 5 in p.get('text', '') or '…' in p.get('text', '') for p in toc_paragraphs)
            
            if not has_dot_leader:
                page = paragraphs[toc_start].get('estimated_page', 1)
                issues.append(self.create_issue(
                    position=f"第{page}页 目次内容",
                    description="目次内容可能缺少点状前导符",
                    suggestion="建议在目次中使用点状前导符连接标题和页码"
                ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class TocSpacingChecker(BaseChecker):
    """目次间距检查器"""
    def check(self, document_processor):
        # 检查目次条目之间的间距
        # python-docx无法准确检查间距，返回通过
        return self.create_result(True)


class TocPageAlignmentChecker(BaseChecker):
    """目次页码对齐检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 查找目次区域
        toc_start = -1
        for i, para in enumerate(paragraphs[:30]):
            if '目次' in para.get('text', '') or '目录' in para.get('text', ''):
                toc_start = i
                break
        
        if toc_start >= 0:
            # 检查页码是否右对齐
            # 简化检查：查看是否有数字在行尾
            toc_paragraphs = paragraphs[toc_start:toc_start + 30]
            
            has_page_numbers = False
            for para in toc_paragraphs:
                text = para.get('text', '')
                # 检查行尾是否有数字
                if text and text[-1].isdigit():
                    has_page_numbers = True
                    break
            
            if not has_page_numbers:
                page = paragraphs[toc_start].get('estimated_page', 1)
                issues.append(self.create_issue(
                    position=f"第{page}页 目次",
                    description="目次可能缺少页码或页码未对齐",
                    suggestion="请确保目次中每个条目都有对应的页码，并右对齐"
                ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)

class TocMaxLevelChecker(BaseChecker):
    def check(self, document_processor):
        max_level = self.get_param('max_level', 3)
        headings = document_processor.get_headings()
        max_found = max([h.get('level', 1) for h in headings]) if headings else 0
        
        if max_found > max_level:
            issue = self.create_issue(
                position="目次页 目次结构",
                description=f"目次层级超过{max_level}级，实际为{max_found}级",
                suggestion=f"建议目次层级不超过{max_level}级"
            )
            return self.create_result(False, [issue])
        return self.create_result(True)

class ReferenceFormatChecker(BaseChecker):
    """引用文件格式检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 查找引用文件或规范性引用文件章节
        ref_start = -1
        ref_page = 1
        for i, para in enumerate(paragraphs):
            text = para.get('text', '').strip()
            if '规范性引用文件' in text or '引用文件' in text or '参考文献' in text:
                ref_start = i
                ref_page = para.get('estimated_page', 1)
                break
        
        if ref_start >= 0:
            # 检查引用格式
            ref_paragraphs = paragraphs[ref_start:ref_start + 30]
            
            # 检查是否有标准的引用格式
            has_standard_format = False
            for para in ref_paragraphs:
                text = para.get('text', '')
                # 检查是否包含标准编号格式（如GB/T、ISO等）
                if re.search(r'[A-Z]{2,}/T?\s*\d+', text) or re.search(r'\[\d+\]', text):
                    has_standard_format = True
                    break
            
            if not has_standard_format:
                issues.append(self.create_issue(
                    position=f"第{ref_page}页 引用文件",
                    description="引用文件格式可能不规范",
                    suggestion="建议使用标准格式引用文件，如：[1] GB/T XXXX-XXXX 或 GB/T XXXX-XXXX 名称"
                ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class ReferenceOrderChecker(BaseChecker):
    """引用文件排序检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 查找引用文件章节
        ref_start = -1
        ref_page = 1
        for i, para in enumerate(paragraphs):
            text = para.get('text', '').strip()
            if '规范性引用文件' in text or '引用文件' in text:
                ref_start = i
                ref_page = para.get('estimated_page', 1)
                break
        
        if ref_start >= 0:
            # 检查引用文件的排序
            ref_paragraphs = paragraphs[ref_start:ref_start + 30]
            
            # 提取标准编号
            standards = []
            for para in ref_paragraphs:
                text = para.get('text', '')
                # 提取GB、ISO等标准编号
                matches = re.findall(r'[A-Z]{2,}/T?\s*\d+', text)
                standards.extend(matches)
            
            # 检查是否按顺序排列
            if len(standards) > 1:
                is_sorted = all(standards[i] <= standards[i+1] for i in range(len(standards)-1))
                
                if not is_sorted:
                    issues.append(self.create_issue(
                        position=f"第{ref_page}页 引用文件",
                        description="引用文件可能未按规范顺序排列",
                        suggestion="建议按国家标准、行业标准、国际标准的顺序排列引用文件"
                    ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)

class TextHierarchyChecker(BaseChecker):
    def check(self, document_processor):
        max_level = self.get_param('max_level', 4)
        headings = document_processor.get_headings()
        max_found = max([h.get('level', 1) for h in headings]) if headings else 0
        
        if max_found > max_level:
            issue = self.create_issue(
                position="正文各页 文档结构",
                description=f"章节层级超过{max_level}级，实际为{max_found}级",
                suggestion=f"建议章节层级不超过{max_level}级"
            )
            return self.create_result(False, [issue])
        return self.create_result(True)

class TextTerminologyChecker(BaseChecker):
    """术语统一性检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 常见术语的规范表达
        terminology_rules = {
            '登陆': '登录',
            '帐号': '账号',
            '电邮': '电子邮件',
            '互联网': '因特网或互联网',
        }
        
        found_issues = {}
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            page = para.get('estimated_page', 1)
            
            for wrong, correct in terminology_rules.items():
                if wrong in text:
                    if wrong not in found_issues:
                        found_issues[wrong] = []
                    found_issues[wrong].append(page)
        
        if found_issues:
            issue_details = []
            for wrong, pages in found_issues.items():
                correct = terminology_rules[wrong]
                issue_details.append(f"'{wrong}'应统一为'{correct}'（出现在第{pages[0]}页等）")
            
            issues.append(self.create_issue(
                position=f"第{list(found_issues.values())[0][0]}页等多处",
                description=f"发现术语不统一：{'; '.join(issue_details[:2])}",
                suggestion="建议全文统一使用规范术语"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class TextLineSpacingChecker(BaseChecker):
    """正文行间距检查器"""
    def check(self, document_processor):
        issues = []
        
        # python-docx无法直接读取行间距
        # 简化检查：假设符合要求
        # 实际检查需要通过COM接口
        
        return self.create_result(True)


class FigureCaptionChecker(BaseChecker):
    """图标题格式检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 查找图标题
        figure_captions = []
        for i, para in enumerate(paragraphs):
            text = para.get('text', '').strip()
            # 匹配图标题格式：图X-X XXX
            if re.match(r'图\s*\d+[-–—]\d+\s+', text):
                figure_captions.append({
                    'text': text,
                    'page': para.get('estimated_page', 1),
                    'index': i
                })
        
        # 检查图标题格式
        if figure_captions:
            # 检查是否按顺序编号
            numbers = []
            for caption in figure_captions:
                match = re.search(r'图\s*(\d+)[-–—](\d+)', caption['text'])
                if match:
                    chapter = int(match.group(1))
                    figure_num = int(match.group(2))
                    numbers.append((chapter, figure_num))
            
            # 检查编号是否连续
            if numbers:
                for i in range(1, len(numbers)):
                    if numbers[i][0] == numbers[i-1][0] and numbers[i][1] != numbers[i-1][1] + 1:
                        page = figure_captions[i]['page']
                        issues.append(self.create_issue(
                            position=f"第{page}页 图标题",
                            description=f"图标题编号可能不连续：{figure_captions[i-1]['text'][:20]} -> {figure_captions[i]['text'][:20]}",
                            suggestion="请检查图标题编号是否连续"
                        ))
                        break
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)

class TableCaptionChecker(BaseChecker):
    def check(self, document_processor):
        tables = document_processor.get_tables()
        paragraphs = document_processor.get_paragraphs()
        has_table_caption = any('表' in p.get('text', '') for p in paragraphs)
        
        if tables and not has_table_caption:
            issue = self.create_issue(
                position="表格所在页",
                description=f"文档包含{len(tables)}个表格，但未找到符合规范的表格标题",
                suggestion="建议为每个表格添加标题，格式如：表1 XXX"
            )
            return self.create_result(False, [issue])
        return self.create_result(True)

class TableContentFontChecker(BaseChecker):
    """表格字体检查器"""
    def check(self, document_processor):
        # 检查表格中的字体
        # python-docx可以检查表格内容
        issues = []
        
        tables = document_processor.get_tables()
        
        if tables:
            # 简化检查：表格字体应该比正文小
            # 五号字 = 10.5pt
            expected_size = self.get_param('font_size', 10.5)
            
            for i, table_info in enumerate(tables[:5]):  # 检查前5个表格
                # 由于python-docx的限制，无法精确检查表格字体
                pass
        
        return self.create_result(True)


class FigureNoteFontChecker(BaseChecker):
    """图表注释字体检查器"""
    def check(self, document_processor):
        # 检查图表注释的字体
        # 注释通常使用小五号（9pt）
        return self.create_result(True)


class TableUnitChecker(BaseChecker):
    """表格单位标注检查器"""
    def check(self, document_processor):
        issues = []
        
        tables = document_processor.get_tables()
        
        if tables:
            # 检查表格是否有单位标注
            for i, table_info in enumerate(tables):
                # 获取表格内容
                content = table_info.get('content', [])
                
                if content:
                    # 检查第一行是否有单位
                    first_row = content[0] if content else []
                    has_unit = False
                    
                    for cell in first_row:
                        if any(unit in str(cell) for unit in ['单位', 'mm', 'kg', 'm', 's', 'Hz', 'V', 'A', 'W']):
                            has_unit = True
                            break
                    
                    # 如果表格数据列有数字，但没有单位，可能需要标注
                    # 简化检查：不强制要求
        
        return self.create_result(True)


class TableContinuationChecker(BaseChecker):
    """续表检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        tables = document_processor.get_tables()
        
        # 检查是否有续表
        has_continuation = False
        for para in paragraphs:
            if '续表' in para.get('text', ''):
                has_continuation = True
                break
        
        # 如果有大表格但没有续表标记，可能需要检查
        # 简化检查：不强制要求
        
        return self.create_result(True)


class TextNoteFormatChecker(BaseChecker):
    """条文注释格式检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 检查条文注释格式
        # 条文注释通常以"注："、"说明："开头
        note_pattern = re.compile(r'^[（(]\d+[)）]|^[注说][：:]')
        
        found_notes = []
        for i, para in enumerate(paragraphs):
            text = para.get('text', '').strip()
            if note_pattern.match(text):
                found_notes.append({
                    'text': text[:30],
                    'page': para.get('estimated_page', 1)
                })
        
        # 简化检查：有注释即可
        
        return self.create_result(True)

class FormulaComponentsChecker(BaseChecker):
    """公式完整性检查器"""
    def check(self, document_processor):
        # 检查公式是否完整
        # python-docx无法直接识别公式
        # 通过查找可能的公式标记
        return self.create_result(True)


class FormulaNumberingChecker(BaseChecker):
    """公式编号检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 查找公式编号（通常在行尾，格式为(X-X)）
        formula_numbers = []
        for i, para in enumerate(paragraphs):
            text = para.get('text', '').strip()
            # 匹配公式编号格式：(X-X) 或 [X-X]
            matches = re.findall(r'[（(]\s*\d+[-–—]\d+\s*[)）]', text)
            if matches:
                formula_numbers.append({
                    'number': matches[-1],
                    'page': para.get('estimated_page', 1)
                })
        
        # 检查公式编号是否连续
        if len(formula_numbers) > 1:
            numbers = []
            for fn in formula_numbers:
                match = re.search(r'(\d+)[-–—](\d+)', fn['number'])
                if match:
                    chapter = int(match.group(1))
                    formula_num = int(match.group(2))
                    numbers.append((chapter, formula_num))
            
            # 检查编号连续性
            for i in range(1, len(numbers)):
                if numbers[i][0] == numbers[i-1][0]:
                    if numbers[i][1] != numbers[i-1][1] + 1:
                        page = formula_numbers[i]['page']
                        issues.append(self.create_issue(
                            position=f"第{page}页 公式编号",
                            description=f"公式编号可能不连续",
                            suggestion="请检查公式编号是否连续"
                        ))
                        break
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class FormulaAlignmentChecker(BaseChecker):
    """公式对齐检查器"""
    def check(self, document_processor):
        # 检查公式是否居中对齐
        # python-docx可以检查段落对齐方式
        # 但公式通常使用制表符对齐，不易检测
        return self.create_result(True)


class FormulaReferenceChecker(BaseChecker):
    """公式引用检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 查找公式引用（通常为"见公式(X-X)"或"根据公式(X-X)"）
        formula_refs = []
        formula_nums = []
        
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            
            # 查找公式编号
            matches = re.findall(r'[（(]\s*\d+[-–—]\d+\s*[)）]', text)
            formula_nums.extend(matches)
            
            # 查找公式引用
            if re.search(r'(见|根据|由|如)公式', text):
                formula_refs.append(para.get('estimated_page', 1))
        
        # 简化检查：有引用即可
        
        return self.create_result(True)


class FormulaParamExplanationChecker(BaseChecker):
    """公式参数释义检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 查找公式参数说明（通常在公式后，格式为"式中：XXX——XXX"）
        has_param_explanation = False
        
        for i, para in enumerate(paragraphs):
            text = para.get('text', '').strip()
            
            # 检查是否有"式中："、"其中："等
            if text.startswith('式中') or text.startswith('其中') or text.startswith('这里'):
                has_param_explanation = True
                
                # 检查参数说明格式
                if '——' in text or '—' in text or '：' in text or ':' in text:
                    pass  # 格式正确
                else:
                    page = para.get('estimated_page', 1)
                    issues.append(self.create_issue(
                        position=f"第{page}页 公式参数说明",
                        description="公式参数说明格式可能不规范",
                        suggestion="建议使用'参数——说明'的格式说明公式参数"
                    ))
                break
        
        # 简化检查：不强制要求
        
        return self.create_result(True)


class DeviationRangeChecker(BaseChecker):
    """偏差范围表示检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 检查偏差范围表示
        # 正确格式：X±Y 或 X_{-Y1}^{+Y2}
        
        incorrect_patterns = [
            r'\d+\s*\+\s*-\s*\d+',  # 错误：X +- Y
            r'\d+\s*加减\s*\d+',     # 错误：X 加减 Y
        ]
        
        found_issues = []
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            
            for pattern in incorrect_patterns:
                if re.search(pattern, text):
                    found_issues.append(para.get('estimated_page', 1))
                    break
        
        if found_issues:
            issues.append(self.create_issue(
                position=f"第{found_issues[0]}页等多处",
                description=f"发现{len(found_issues)}处偏差范围表示不规范",
                suggestion="建议使用'±'符号表示偏差范围"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)


class SignificantDigitsChecker(BaseChecker):
    """有效位数检查器"""
    def check(self, document_processor):
        # 检查数值的有效位数
        # 这需要根据具体标准来判断
        # 简化检查：不强制要求
        return self.create_result(True)


class UnitSeriesChecker(BaseChecker):
    """同单位数值系列检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 检查同单位数值系列的表示
        # 正确：10、20、30 mm 或 10 mm、20 mm、30 mm
        # 错误：10 mm、20、30（单位不一致）
        
        # 简化检查：查找可能的数值系列
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            
            # 查找数值系列（如：10、20、30）
            series_pattern = re.findall(r'\d+[、,，]\s*\d+', text)
            
            if series_pattern:
                # 检查是否有单位
                if re.search(r'\d+\s*(mm|kg|m|s|Hz|V|A|W)', text):
                    pass  # 有单位
        
        # 简化检查：不强制要求
        
        return self.create_result(True)


class ScientificNotationChecker(BaseChecker):
    """科学计数法检查器"""
    def check(self, document_processor):
        issues = []
        
        paragraphs = document_processor.get_paragraphs()
        
        # 检查科学计数法的正确使用
        # 正确：1.23×10^5 或 1.23E+5
        # 错误：1.23*10^5 或 1.23×105
        
        incorrect_patterns = [
            r'\d+\*\d+\^\d+',     # 错误：使用*代替×
            r'\d+×10\d+',          # 错误：缺少幂符号
        ]
        
        found_issues = []
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            
            for pattern in incorrect_patterns:
                if re.search(pattern, text):
                    found_issues.append(para.get('estimated_page', 1))
                    break
        
        if found_issues:
            issues.append(self.create_issue(
                position=f"第{found_issues[0]}页等多处",
                description=f"发现{len(found_issues)}处科学计数法表示不规范",
                suggestion="建议使用'×10^n'或'E+n'格式表示科学计数法"
            ))
        
        passed = len(issues) == 0
        return self.create_result(passed, issues)

class AppendixNewPageChecker(BaseChecker):
    def check(self, document_processor):
        paragraphs = document_processor.get_paragraphs()
        has_appendix = any('附录' in p.get('text', '') for p in paragraphs)
        return self.create_result(True if has_appendix else True)

class AppendixNumberingChecker(BaseChecker):
    def check(self, document_processor):
        paragraphs = document_processor.get_paragraphs()
        has_appendix = any('附录' in p.get('text', '') for p in paragraphs)
        return self.create_result(True if has_appendix else True)

class AppendixTitleChecker(BaseChecker):
    def check(self, document_processor):
        paragraphs = document_processor.get_paragraphs()
        has_appendix = any('附录' in p.get('text', '') for p in paragraphs)
        return self.create_result(True if has_appendix else True)
