"""
Word文档检查器 - EXE打包脚本
使用PyInstaller将项目打包成独立运行的exe文件
"""
import os
import sys
import shutil
import subprocess
from pathlib import Path

# 项目根目录
PROJECT_ROOT = Path(__file__).parent.absolute()

# 需要包含的数据文件
DATA_FILES = [
    ('config/rules.json', 'config'),
]

# 需要包含的隐藏导入
HIDDEN_IMPORTS = [
    # 项目模块
    'ui',
    'ui.main_window',
    'core',
    'core.rule_engine',
    'core.word_processor',
    'core.base_checker',
    'core.checkers',
    'core.checkers.technical_doc_checkers',
    'models',
    'models.check_result',
    'models.rule',
    'utils',
    # python-docx
    'pyparsing',
    'docx',
    'docx.opc',
    'docx.opc.constants',
    'docx.oxml',
    'docx.oxml.ns',
    'docx.parts',
    'docx.text',
    'docx.enum',
    'docx.enum.text',
    'docx.enum.style',
    'docx.shared',
    # win32com
    'win32com',
    'win32com.client',
    'pythoncom',
    'pywintypes',
    # openpyxl
    'openpyxl',
    'openpyxl.cell',
    'openpyxl.styles',
    # PIL
    'PIL',
    'PIL._imaging',
]


def clean_build():
    """清理构建目录"""
    print("[INFO] Cleaning build directories...")
    
    dirs_to_clean = ['build', 'dist', '__pycache__']
    for dir_name in dirs_to_clean:
        dir_path = PROJECT_ROOT / dir_name
        if dir_path.exists():
            shutil.rmtree(dir_path)
            print(f"  - Removed: {dir_name}")
    
    # 删除spec文件
    spec_file = PROJECT_ROOT / 'WordChecker.spec'
    if spec_file.exists():
        spec_file.unlink()
        print(f"  - Removed: WordChecker.spec")


def install_pyinstaller():
    """安装PyInstaller"""
    print("\n[INFO] Checking PyInstaller installation...")
    try:
        import PyInstaller
        print("  - PyInstaller already installed")
    except ImportError:
        print("  - Installing PyInstaller...")
        subprocess.check_call([sys.executable, '-m', 'pip', 'install', 'pyinstaller'])
        print("  - PyInstaller installed successfully")


def build_exe():
    """构建exe文件"""
    print("\n[INFO] Building executable...")
    
    # 构建pyinstaller命令
    cmd = [
        sys.executable, '-m', 'PyInstaller',
        '--name=WordChecker',
        '--onefile',
        '--windowed',
        '--clean',
        '--noconfirm',
        '--paths=src',  # 添加src目录到Python路径
    ]
    
    # 添加数据文件
    for src, dst in DATA_FILES:
        src_path = PROJECT_ROOT / src
        if src_path.exists():
            cmd.append(f'--add-data={src};{dst}')
            print(f"  - Adding data: {src} -> {dst}")
    
    # 添加隐藏导入
    for module in HIDDEN_IMPORTS:
        cmd.append(f'--hidden-import={module}')
    
    # 添加主入口
    cmd.append('src/main.py')
    
    # 执行打包
    print(f"\n[INFO] Running PyInstaller...")
    print(f"  Command: {' '.join(cmd)}\n")
    
    try:
        subprocess.check_call(cmd, cwd=PROJECT_ROOT)
        print("\n[SUCCESS] Build completed!")
        return True
    except subprocess.CalledProcessError as e:
        print(f"\n[ERROR] Build failed: {e}")
        return False


def create_readme():
    """创建使用说明"""
    readme_content = """# Word文档检查器 - 使用说明

## 运行方法

1. 双击 WordChecker.exe 启动程序
2. 点击"选择文件"按钮，选择要检查的Word文档（.docx格式）
3. 点击"开始检查"按钮，等待检查完成
4. 查看检查结果，可以导出为HTML、JSON、Excel格式

## 系统要求

- Windows 7/8/10/11
- 不需要安装Python或其他依赖

## 功能说明

### 检查项目

程序会对Word文档进行以下检查：

1. **文档结构检查**
   - 文档编号格式
   - 文档名称格式
   - 签署完整性
   - 公司名称

2. **页面设置检查**
   - 页边距
   - 页码格式
   - 页面方向

3. **目次检查**
   - 目次完整性
   - 目次格式
   - 目次页码对齐

4. **内容检查**
   - 标题格式
   - 正文字体
   - 术语统一性
   - 图表标题

5. **表格检查**
   - 表格字体
   - 表格单位
   - 续表处理

6. **公式检查**
   - 公式编号
   - 公式对齐
   - 参数说明

7. **数值检查**
   - 计量单位
   - 数学符号
   - 有效位数

### 导出功能

检查完成后，可以导出以下格式的报告：
- HTML格式：适合浏览器查看
- JSON格式：适合程序处理
- Excel格式：适合数据分析和存档

## 注意事项

1. 仅支持 .docx 格式（Word 2007及以上版本）
2. 文档检查过程中请勿关闭程序
3. 大型文档检查时间较长，请耐心等待
4. 首次运行可能需要较长时间加载

## 问题反馈

如遇到问题，请提供以下信息：
- 操作系统版本
- Word文档示例
- 错误提示截图

---
版本：1.0.0
更新时间：2024
"""
    
    readme_path = PROJECT_ROOT / 'dist' / '使用说明.txt'
    with open(readme_path, 'w', encoding='utf-8') as f:
        f.write(readme_content)
    print(f"\n[INFO] Created usage guide: {readme_path}")


def copy_config():
    """复制配置文件到dist目录"""
    print("\n[INFO] Copying configuration files...")
    
    dist_config = PROJECT_ROOT / 'dist' / 'config'
    dist_config.mkdir(exist_ok=True)
    
    for src, dst in DATA_FILES:
        src_path = PROJECT_ROOT / src
        dst_path = dist_config / Path(src).name
        if src_path.exists():
            shutil.copy2(src_path, dst_path)
            print(f"  - Copied: {src} -> dist/config/{Path(src).name}")


def main():
    """主函数"""
    print("="*60)
    print("Word文档检查器 - 打包工具")
    print("="*60)
    
    # 清理旧的构建文件
    clean_build()
    
    # 安装PyInstaller
    install_pyinstaller()
    
    # 构建exe
    if build_exe():
        # 创建使用说明
        create_readme()
        
        # 复制配置文件
        copy_config()
        
        print("\n" + "="*60)
        print("[SUCCESS] 打包完成！")
        print("="*60)
        print(f"\n输出目录: {PROJECT_ROOT / 'dist'}")
        print(f"可执行文件: {PROJECT_ROOT / 'dist' / 'WordChecker.exe'}")
        print(f"使用说明: {PROJECT_ROOT / 'dist' / '使用说明.txt'}")
        print("\n使用方法:")
        print("  1. 双击 WordChecker.exe 运行程序")
        print("  2. 选择要检查的Word文档")
        print("  3. 点击开始检查")
        print("\n注意事项:")
        print("  - config 文件夹必须与 WordChecker.exe 在同一目录")
        print("  - 首次运行可能需要较长时间加载")
    else:
        print("\n[ERROR] 打包失败，请检查错误信息")
        sys.exit(1)


if __name__ == '__main__':
    main()
