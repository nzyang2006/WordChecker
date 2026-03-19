@echo off
chcp 65001 >nul
echo ============================================================
echo PyInstaller 安装工具
echo ============================================================
echo.
echo 正在尝试多个镜像源安装 PyInstaller...
echo.

:: 尝试清华镜像
echo [1/4] 尝试清华镜像源...
python -m pip install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple
if errorlevel 1 (
    echo [WARN] 清华镜像源失败
    echo.
    
    :: 尝试阿里云镜像
    echo [2/4] 尝试阿里云镜像源...
    python -m pip install pyinstaller -i https://mirrors.aliyun.com/pypi/simple/
    if errorlevel 1 (
        echo [WARN] 阿里云镜像源失败
        echo.
        
        :: 尝试豆瓣镜像
        echo [3/4] 尝试豆瓣镜像源...
        python -m pip install pyinstaller -i https://pypi.douban.com/simple
        if errorlevel 1 (
            echo [WARN] 豆瓣镜像源失败
            echo.
            
            :: 尝试官方源
            echo [4/4] 尝试PyPI官方源...
            python -m pip install pyinstaller
            if errorlevel 1 (
                echo.
                echo ============================================================
                echo [ERROR] 所有镜像源都失败！
                echo ============================================================
                echo.
                echo 可能的原因：
                echo   1. 网络连接问题
                echo   2. 防火墙阻止
                echo   3. 需要配置代理
                echo.
                echo 解决方案：
                echo   1. 检查网络连接
                echo   2. 关闭防火墙或添加例外
                echo   3. 配置代理后重试
                echo   4. 使用离线安装（参考"打包指南.md"）
                echo.
                pause
                exit /b 1
            )
        )
    )
)

echo.
echo ============================================================
echo [SUCCESS] PyInstaller 安装成功！
echo ============================================================
echo.
echo 现在可以运行 build.bat 进行打包
echo.

pause
