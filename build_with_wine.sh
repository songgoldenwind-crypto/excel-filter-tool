#!/bin/bash
# 在Mac上使用Wine打包Windows exe的自动化脚本

set -e

echo "========================================"
echo "Excel工具 - Mac打包Windows exe脚本"
echo "========================================"
echo ""

# 检查Wine是否安装
if ! command -v wine &> /dev/null; then
    echo "❌ 错误: 未找到Wine"
    echo ""
    echo "请先安装Wine："
    echo "  brew install --cask wine-stable"
    echo ""
    exit 1
fi

echo "✓ Wine版本: $(wine --version)"
echo ""

# 设置Wine环境
export WINEPREFIX="$HOME/.wine_pyinstaller"

# 检查Wine环境是否存在
if [ ! -d "$WINEPREFIX" ]; then
    echo "初始化Wine环境..."
    wineboot -i
    echo "✓ Wine环境初始化完成"
    echo ""
fi

# 检查Python是否已安装到Wine
PYTHON_PATH="$WINEPREFIX/drive_c/python/python.exe"

if [ ! -f "$PYTHON_PATH" ]; then
    echo "❌ 错误: Wine环境中未找到Python"
    echo ""
    echo "请按照以下步骤安装："
    echo "1. 访问 https://www.python.org/downloads/windows/"
    echo "2. 下载 'Windows embeddable package' (32-bit或64-bit)"
    echo "3. 解压到 ~/Downloads/python_windows"
    echo "4. 运行以下命令："
    echo ""
    echo "   mkdir -p $WINEPREFIX/drive_c/python"
    echo "   cp -r ~/Downloads/python_windows/* $WINEPREFIX/drive_c/python/"
    echo ""
    exit 1
fi

echo "✓ 找到Python: $PYTHON_PATH"

# 检查pip
PIP_PATH="$WINEPREFIX/drive_c/python/Scripts/pip.exe"

if [ ! -f "$PIP_PATH" ]; then
    echo ""
    echo "未找到pip，正在安装..."

    # 下载get-pip.py
    if [ ! -f "/tmp/get-pip.py" ]; then
        curl -s https://bootstrap.pypa.io/get-pip.py -o /tmp/get-pip.py
    fi

    wine "$PYTHON_PATH" /tmp/get-pip.py
fi

echo "✓ 找到pip"
echo ""

# 检查依赖
echo "检查Python依赖..."

wine "$PIP_PATH" show openpyxl > /dev/null 2>&1 || {
    echo "正在安装 openpyxl..."
    wine "$PIP_PATH" install openpyxl
}

wine "$PIP_PATH" show pyinstaller > /dev/null 2>&1 || {
    echo "正在安装 pyinstaller..."
    wine "$PIP_PATH" install pyinstaller
}

echo "✓ 依赖已就绪"
echo ""

# 开始打包
echo "开始打包..."
echo "这可能需要几分钟，请耐心等待..."
echo ""

# 清理之前的构建
rm -rf build dist spec *.spec 2>/dev/null || true

# 使用Wine中的PyInstaller打包
wine "$WINEPREFIX/drive_c/python/Scripts/pyinstaller.exe" \
    --onefile \
    --windowed \
    --name="Excel数据筛选工具" \
    --clean \
    excel_filter_tool.py

echo ""
echo "========================================"
if [ -f "dist/Excel数据筛选工具.exe" ]; then
    echo "✓ 打包成功！"
    echo ""
    echo "可执行文件位置:"
    echo "  $(pwd)/dist/Excel数据筛选工具.exe"
    echo ""
    echo "文件大小:"
    ls -lh "dist/Excel数据筛选工具.exe" | awk '{print "  " $5}'
    echo ""
    echo "现在可以将此exe文件复制到Windows电脑使用！"
else
    echo "❌ 打包失败，请检查错误信息"
fi
echo "========================================"
