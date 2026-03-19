#!/bin/bash

echo "正在启动Word文档检查器..."
echo ""

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "错误: 未检测到Python，请先安装Python 3.8+"
    exit 1
fi

# 检查依赖是否安装
if ! python3 -c "import PyQt5" &> /dev/null; then
    echo "正在安装依赖..."
    pip3 install -r requirements.txt
    if [ $? -ne 0 ]; then
        echo "错误: 依赖安装失败"
        exit 1
    fi
fi

# 启动应用
cd src
python3 main.py
