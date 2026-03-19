# Word文档检查器

一个功能完善的Word文档检查桌面应用程序，支持通过配置文件加载检查规则，对Word文档进行全面的质量检查。

## 功能特性

### 核心功能
- ✅ **配置化规则管理** - 通过JSON配置文件灵活管理检查规则
- ✅ **多维度检查** - 支持格式、内容、结构等多个维度的检查
- ✅ **智能批注** - 自动在Word文档中添加批注标记问题
- ✅ **直观界面** - 现代化的桌面应用界面，操作简单直观
- ✅ **详细报告** - 生成详细的检查报告，包含问题位置和修改建议

### 检查规则

#### 格式检查
- 标题格式检查 - 检查文档标题是否符合格式要求
- 字体一致性检查 - 检查文档中字体是否保持一致
- 段落间距检查 - 检查段落间距是否符合规范
- 页码检查 - 检查文档是否包含页码
- 日期格式检查 - 检查日期格式是否统一
- 数字格式检查 - 检查数字格式是否规范

#### 内容检查
- 表格标题检查 - 检查表格是否有标题
- 图片标题检查 - 检查图片是否有标题
- 敏感词检查 - 检查文档中是否包含敏感词

#### 结构检查
- 标题层级检查 - 检查标题层级是否连续

## 安装

### 系统要求
- Python 3.8+
- Windows操作系统（批注功能需要）

### 安装依赖

```bash
pip install -r requirements.txt
```

## 使用方法

### 启动应用

#### Windows
```bash
# 方式1：双击 run.bat
run.bat

# 方式2：命令行启动
cd src
python main.py
```

#### Linux/Mac
```bash
# 方式1：运行脚本
chmod +x run.sh
./run.sh

# 方式2：命令行启动
cd src
python3 main.py
```

### 使用步骤

1. **选择文件** - 点击"选择Word文件"按钮，选择要检查的.docx文档
2. **选择规则** - 在规则列表中勾选需要执行的检查规则
3. **开始检查** - 点击"开始检查"按钮，等待检查完成
4. **查看结果** - 在右侧面板查看检查结果概览和详细信息
5. **添加批注** - 点击"添加批注到文档"按钮，将问题标注到Word文档中

### 配置规则

规则配置文件位于 `config/rules.json`，可以在应用的"规则配置"标签页中直接编辑。

#### 规则配置格式

```json
{
  "id": "rule_id",
  "name": "规则名称",
  "description": "规则描述",
  "category": "格式检查",
  "enabled": true,
  "checker": "CheckerClass",
  "params": {
    "param1": "value1"
  }
}
```

#### 添加自定义规则

1. 在配置文件中添加新的规则定义
2. 在 `src/core/checkers/` 中实现对应的检查器类
3. 在 `src/core/checkers/__init__.py` 中注册检查器

## 项目结构

```
WordChecker/
├── config/                   # 配置文件目录
│   └── rules.json           # 检查规则配置
├── src/                      # 源代码目录
│   ├── main.py              # 主程序入口
│   ├── core/                # 核心功能模块
│   │   ├── rule_engine.py   # 规则引擎
│   │   ├── word_processor.py # Word处理器
│   │   ├── base_checker.py  # 检查器基类
│   │   └── checkers/        # 检查器实现
│   │       ├── format_checkers.py    # 格式检查器
│   │       ├── content_checkers.py   # 内容检查器
│   │       └── structure_checkers.py # 结构检查器
│   ├── models/              # 数据模型
│   │   ├── rule.py          # 规则模型
│   │   └── check_result.py  # 检查结果模型
│   ├── ui/                  # 用户界面
│   │   └── main_window.py   # 主窗口
│   └── utils/               # 工具函数
│       └── config_loader.py # 配置加载器
├── requirements.txt          # 项目依赖
├── run.bat                  # Windows启动脚本
├── run.sh                   # Linux/Mac启动脚本
└── README.md                # 项目说明
```

## 开发自定义检查器

### 检查器基类

所有检查器都继承自 `BaseChecker`：

```python
from src.core.base_checker import BaseChecker
from src.models.check_result import CheckIssue

class MyCustomChecker(BaseChecker):
    def check(self, document_processor) -> CheckResult:
        """执行检查"""
        issues = []
        
        # 获取文档内容
        paragraphs = document_processor.get_paragraphs()
        
        # 执行检查逻辑
        for i, para in enumerate(paragraphs):
            text = para.get('text', '')
            
            # 检查条件
            if "某些条件":
                issues.append(self.create_issue(
                    position=f"第{i+1}段",
                    description="问题描述",
                    suggestion="修改建议"
                ))
        
        # 返回结果
        passed = len(issues) == 0
        return self.create_result(passed, issues)
```

### 注册检查器

在 `src/core/checkers/__init__.py` 中添加：

```python
from .my_checker import MyCustomChecker

CHECKER_MAP = {
    # ...
    'MyCustomChecker': MyCustomChecker,
}
```

### 在配置文件中使用

在 `config/rules.json` 中添加：

```json
{
  "id": "my_custom_rule",
  "name": "我的自定义检查",
  "description": "检查描述",
  "category": "自定义检查",
  "enabled": true,
  "checker": "MyCustomChecker",
  "params": {
    "param1": "value1"
  }
}
```

## 技术栈

- **UI框架** - PyQt5
- **Word处理** - python-docx, pywin32
- **配置管理** - JSON

## 注意事项

1. **批注功能** - 需要在Windows系统上运行，依赖Microsoft Word COM接口
2. **文件格式** - 仅支持.docx格式的Word文档（不支持.doc）
3. **文件备份** - 添加批注前建议备份原始文档
4. **性能优化** - 大型文档检查可能需要较长时间，请耐心等待

## 常见问题

### Q: 启动时提示缺少依赖？
A: 运行 `pip install -r requirements.txt` 安装所有依赖

### Q: 批注功能无法使用？
A: 批注功能仅支持Windows系统，需要安装Microsoft Word

### Q: 检查速度很慢？
A: 大型文档检查需要较长时间，可以通过减少检查规则数量来提高速度

### Q: 如何添加自定义规则？
A: 参考"开发自定义检查器"章节，实现检查器类并在配置文件中注册

## 许可证

本项目仅供学习和内部使用。

## 联系方式

如有问题或建议，请联系开发团队。
