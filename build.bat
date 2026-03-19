@echo off
chcp 65001 >nul
echo ============================================================
echo Word文档检查器 - 打包工具
echo ============================================================
echo.

:: 检查Python是否安装
python --version >nul 2>&1
if errorlevel 1 (
    echo [ERROR] 未检测到Python，请先安装Python 3.8或更高版本
    pause
    exit /b 1
)

:: 运行打包脚本
echo [INFO] 开始打包...
python build_exe.py

if errorlevel 1 (
    echo.
    echo [ERROR] 打包失败！
    pause
    exit /b 1
)

echo.
echo ============================================================
echo 打包完成！
echo ============================================================
echo.
echo 输出目录: dist\
echo 可执行文件: dist\WordChecker.exe
echo.
echo 使用方法:
echo   1. 进入 dist 目录
echo   2. 双击 WordChecker.exe 运行程序
echo   3. 确保 config 文件夹与 WordChecker.exe 在同一目录
echo.

pause
