#!/bin/bash
echo "========================================"
echo "Excel筛选工具打包脚本"
echo "========================================"
echo ""

# 检查Python是否安装
if ! command -v python3 &> /dev/null; then
    echo "错误: 未找到Python，请先安装Python 3.8或更高版本"
    exit 1
fi

echo "正在安装依赖..."
pip3 install -r requirements.txt -i https://pypi.tuna.tsinghua.edu.cn/simple
if [ $? -ne 0 ]; then
    echo "依赖安装失败"
    exit 1
fi

echo ""
echo "正在打包程序..."
pyinstaller --onefile --windowed --name="Excel数据筛选工具" --clean excel_filter_tool.py
if [ $? -ne 0 ]; then
    echo "打包失败"
    exit 1
fi

echo ""
echo "========================================"
echo "打包完成！"
echo "可执行文件位置: dist/Excel数据筛选工具"
echo "========================================"
echo ""
