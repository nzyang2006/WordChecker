"""
检查器模块
"""
from .format_checkers import (
    TitleFormatChecker,
    FontConsistencyChecker,
    ParagraphSpacingChecker,
    PageNumberChecker,
    DateFormatChecker,
    NumberFormatChecker
)
from .content_checkers import (
    TableCaptionChecker as TableCaptionCheckerOld,
    ImageCaptionChecker,
    ForbiddenWordsChecker
)
from .structure_checkers import HeadingSequenceChecker
from .technical_doc_checkers import (
    # 文档封面和签署页
    CoverDocumentNumberChecker,
    CoverDocumentNameChecker,
    CoverSignatureChecker,
    CoverCompanyNameChecker,
    CoverPageSetupChecker,
    # 页码
    PageCoverNoNumberChecker,
    PageTocRomanChecker,
    PageMainArabicChecker,
    # 目次
    TocRequiredChecker,
    TocContentChecker,
    TocTitleFormatChecker,
    TocContentFormatChecker,
    TocSpacingChecker,
    TocPageAlignmentChecker,
    TocMaxLevelChecker,
    # 引用文件
    ReferenceFormatChecker,
    ReferenceOrderChecker,
    # 正文
    TextHierarchyChecker,
    TextTerminologyChecker,
    TextFontChecker,
    TextLineSpacingChecker,
    TextHeadingFormatChecker,
    # 图表
    FigureCaptionChecker,
    TableCaptionChecker,
    TableContentFontChecker,
    FigureNoteFontChecker,
    TableUnitChecker,
    TableContinuationChecker,
    TextNoteFormatChecker,
    # 公式
    FormulaComponentsChecker,
    FormulaNumberingChecker,
    FormulaAlignmentChecker,
    FormulaReferenceChecker,
    FormulaParamExplanationChecker,
    # 量、单位及其符号
    UnitLegalChecker,
    MathSymbolChecker,
    DeviationRangeChecker,
    SignificantDigitsChecker,
    UnitSeriesChecker,
    ScientificNotationChecker,
    # 附录
    AppendixNewPageChecker,
    AppendixNumberingChecker,
    AppendixTitleChecker,
)

# 检查器映射表
CHECKER_MAP = {
    # 旧版检查器
    'TitleFormatChecker': TitleFormatChecker,
    'FontConsistencyChecker': FontConsistencyChecker,
    'ParagraphSpacingChecker': ParagraphSpacingChecker,
    'PageNumberChecker': PageNumberChecker,
    'DateFormatChecker': DateFormatChecker,
    'NumberFormatChecker': NumberFormatChecker,
    'ImageCaptionChecker': ImageCaptionChecker,
    'ForbiddenWordsChecker': ForbiddenWordsChecker,
    'HeadingSequenceChecker': HeadingSequenceChecker,
    
    # 技术文档检查器 - 文档封面和签署页
    'CoverDocumentNumberChecker': CoverDocumentNumberChecker,
    'CoverDocumentNameChecker': CoverDocumentNameChecker,
    'CoverSignatureChecker': CoverSignatureChecker,
    'CoverCompanyNameChecker': CoverCompanyNameChecker,
    'CoverPageSetupChecker': CoverPageSetupChecker,
    
    # 技术文档检查器 - 页码
    'PageCoverNoNumberChecker': PageCoverNoNumberChecker,
    'PageTocRomanChecker': PageTocRomanChecker,
    'PageMainArabicChecker': PageMainArabicChecker,
    
    # 技术文档检查器 - 目次
    'TocRequiredChecker': TocRequiredChecker,
    'TocContentChecker': TocContentChecker,
    'TocTitleFormatChecker': TocTitleFormatChecker,
    'TocContentFormatChecker': TocContentFormatChecker,
    'TocSpacingChecker': TocSpacingChecker,
    'TocPageAlignmentChecker': TocPageAlignmentChecker,
    'TocMaxLevelChecker': TocMaxLevelChecker,
    
    # 技术文档检查器 - 引用文件
    'ReferenceFormatChecker': ReferenceFormatChecker,
    'ReferenceOrderChecker': ReferenceOrderChecker,
    
    # 技术文档检查器 - 正文
    'TextHierarchyChecker': TextHierarchyChecker,
    'TextTerminologyChecker': TextTerminologyChecker,
    'TextFontChecker': TextFontChecker,
    'TextLineSpacingChecker': TextLineSpacingChecker,
    'TextHeadingFormatChecker': TextHeadingFormatChecker,
    
    # 技术文档检查器 - 图表
    'FigureCaptionChecker': FigureCaptionChecker,
    'TableCaptionChecker': TableCaptionChecker,
    'TableContentFontChecker': TableContentFontChecker,
    'FigureNoteFontChecker': FigureNoteFontChecker,
    'TableUnitChecker': TableUnitChecker,
    'TableContinuationChecker': TableContinuationChecker,
    'TextNoteFormatChecker': TextNoteFormatChecker,
    
    # 技术文档检查器 - 公式
    'FormulaComponentsChecker': FormulaComponentsChecker,
    'FormulaNumberingChecker': FormulaNumberingChecker,
    'FormulaAlignmentChecker': FormulaAlignmentChecker,
    'FormulaReferenceChecker': FormulaReferenceChecker,
    'FormulaParamExplanationChecker': FormulaParamExplanationChecker,
    
    # 技术文档检查器 - 量、单位及其符号
    'UnitLegalChecker': UnitLegalChecker,
    'MathSymbolChecker': MathSymbolChecker,
    'DeviationRangeChecker': DeviationRangeChecker,
    'SignificantDigitsChecker': SignificantDigitsChecker,
    'UnitSeriesChecker': UnitSeriesChecker,
    'ScientificNotationChecker': ScientificNotationChecker,
    
    # 技术文档检查器 - 附录
    'AppendixNewPageChecker': AppendixNewPageChecker,
    'AppendixNumberingChecker': AppendixNumberingChecker,
    'AppendixTitleChecker': AppendixTitleChecker,
}

__all__ = [
    'CHECKER_MAP',
    # 旧版检查器
    'TitleFormatChecker',
    'FontConsistencyChecker',
    'ParagraphSpacingChecker',
    'PageNumberChecker',
    'DateFormatChecker',
    'NumberFormatChecker',
    'TableCaptionCheckerOld',
    'ImageCaptionChecker',
    'ForbiddenWordsChecker',
    'HeadingSequenceChecker',
    # 技术文档检查器
    'CoverDocumentNumberChecker',
    'CoverDocumentNameChecker',
    'CoverSignatureChecker',
    'CoverCompanyNameChecker',
    'CoverPageSetupChecker',
    'PageCoverNoNumberChecker',
    'PageTocRomanChecker',
    'PageMainArabicChecker',
    'TocRequiredChecker',
    'TocContentChecker',
    'TocTitleFormatChecker',
    'TocContentFormatChecker',
    'TocSpacingChecker',
    'TocPageAlignmentChecker',
    'TocMaxLevelChecker',
    'ReferenceFormatChecker',
    'ReferenceOrderChecker',
    'TextHierarchyChecker',
    'TextTerminologyChecker',
    'TextFontChecker',
    'TextLineSpacingChecker',
    'TextHeadingFormatChecker',
    'FigureCaptionChecker',
    'TableCaptionChecker',
    'TableContentFontChecker',
    'FigureNoteFontChecker',
    'TableUnitChecker',
    'TableContinuationChecker',
    'TextNoteFormatChecker',
    'FormulaComponentsChecker',
    'FormulaNumberingChecker',
    'FormulaAlignmentChecker',
    'FormulaReferenceChecker',
    'FormulaParamExplanationChecker',
    'UnitLegalChecker',
    'MathSymbolChecker',
    'DeviationRangeChecker',
    'SignificantDigitsChecker',
    'UnitSeriesChecker',
    'ScientificNotationChecker',
    'AppendixNewPageChecker',
    'AppendixNumberingChecker',
    'AppendixTitleChecker',
]
