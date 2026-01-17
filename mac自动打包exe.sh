#!/bin/bash
# Mac上自动打包Windows exe - 完整自动化脚本

set -e

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
RED='\033[0;31m'
NC='\033[0m'

SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
WINE_PREFIX="$SCRIPT_DIR/wine_env"
PYTHON_VERSION="3.10.11"
PYTHON_DIR="$WINE_PREFIX/drive_c/python310"

clear
echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}  Excel工具 - Mac打包Windows exe${NC}"
echo -e "${BLUE}========================================${NC}"
echo ""

# ============ 步骤1: 检查Wine ============
echo -e "${YELLOW}[1/6]${NC} 检查Wine..."
if ! command -v wine &> /dev/null; then
    echo -e "${RED}✗${NC} Wine未安装"
    echo ""
    echo "正在安装Wine，这可能需要10-20分钟..."
    echo "请输入您的Mac密码以授权安装："
    echo ""

    brew install --cask wine-stable

    if [ $? -ne 0 ]; then
        echo -e "${RED}✗${NC} Wine安装失败"
        echo "请手动运行: brew install --cask wine-stable"
        exit 1
    fi
fi

WINE_VER=$(wine --version)
echo -e "${GREEN}✓${NC} Wine已安装: $WINE_VER"
echo ""

# ============ 步骤2: 初始化Wine环境 ============
echo -e "${YELLOW}[2/6]${NC} 初始化Wine环境..."
export WINEPREFIX="$WINE_PREFIX"
export WINEARCH=win64

if [ ! -d "$WINE_PREFIX" ]; then
    echo "创建Wine环境（首次运行会显示配置窗口，请点击'安装'）"
    wineboot -i
    sleep 5
fi

echo -e "${GREEN}✓${NC} Wine环境就绪"
echo ""

# ============ 步骤3: 下载Windows版Python ============
echo -e "${YELLOW}[3/6]${NC} 准备Windows版Python..."

PYTHON_ZIP="$SCRIPT_DIR/python_embed.zip"
PYTHON_URL="https://www.python.org/ftp/python/${PYTHON_VERSION}/python-${PYTHON_VERSION}-embed-amd64.zip"

if [ ! -f "$PYTHON_ZIP" ]; then
    echo "正在下载Python ${PYTHON_VERSION}（约25MB）..."

    if command -v curl &> /dev/null; then
        curl -L -o "$PYTHON_ZIP" "$PYTHON_URL"
    elif command -v wget &> /dev/null; then
        wget -O "$PYTHON_ZIP" "$PYTHON_URL"
    else
        echo -e "${RED}✗${NC} 未找到curl或wget"
        exit 1
    fi
fi

echo -e "${GREEN}✓${NC} Python已下载"
echo ""

# ============ 步骤4: 解压并配置Python ============
echo -e "${YELLOW}[4/6]${NC} 配置Python环境..."

# 创建Python目录
mkdir -p "$PYTHON_DIR"

# 解压Python（如果还未解压）
if [ ! -f "$PYTHON_DIR/python.exe" ]; then
    echo "正在解压Python..."
    unzip -q -o "$PYTHON_ZIP" -d "$PYTHON_DIR"

    # 修改python310._pth以支持site-packages
    cat > "$PYTHON_DIR/python310._pth" << 'EOF'
python310.zip
.
../../Lib/site-packages
EOF
fi

echo -e "${GREEN}✓${NC} Python配置完成"
echo ""

# ============ 步骤5: 安装依赖 ============
echo -e "${YELLOW}[5/6]${NC} 安装依赖包（openpyxl, pyinstaller）..."

# 下载get-pip
if [ ! -f "$SCRIPT_DIR/get-pip.py" ]; then
    curl -sS https://bootstrap.pypa.io/get-pip.py -o "$SCRIPT_DIR/get-pip.py"
fi

# 安装pip
wine "$PYTHON_DIR/python.exe" "$SCRIPT_DIR/get-pip.py" 2>/dev/null || true

# 安装依赖
wine "$PYTHON_DIR/Scripts/pip.exe" install --upgrade pip -i https://pypi.tuna.tsinghua.edu.cn/simple 2>/dev/null || true
wine "$PYTHON_DIR/Scripts/pip.exe" install openpyxl -i https://pypi.tuna.tsinghua.edu.cn/simple 2>/dev/null || true
wine "$PYTHON_DIR/Scripts/pip.exe" install pyinstaller -i https://pypi.tuna.tsinghua.edu.cn/simple 2>/dev/null || true

echo -e "${GREEN}✓${NC} 依赖安装完成"
echo ""

# ============ 步骤6: 打包exe ============
echo -e "${YELLOW}[6/6]${NC} 开始打包exe..."
echo "这可能需要1-2分钟，请耐心等待..."
echo ""

cd "$SCRIPT_DIR"

# 清理之前的构建
rm -rf build dist *.spec 2>/dev/null || true

# 使用PyInstaller打包
wine "$PYTHON_DIR/Scripts/pyinstaller.exe" \
    --onefile \
    --windowed \
    --name="Excel数据筛选工具" \
    --clean \
    --noconfirm \
    excel_filter_tool.py

echo ""
if [ -f "dist/Excel数据筛选工具.exe" ]; then
    echo -e "${GREEN}========================================${NC}"
    echo -e "${GREEN}  ✓ 打包成功！${NC}"
    echo -e "${GREEN}========================================${NC}"
    echo ""
    echo "文件位置: $SCRIPT_DIR/dist/Excel数据筛选工具.exe"
    ls -lh "dist/Excel数据筛选工具.exe" | awk '{printf "文件大小: %s\n", $5}'
    echo ""
    echo -e "${GREEN}现在可以将exe文件拷贝到Windows电脑使用了！${NC}"
    echo ""

    # 打开Finder
    open -R "dist/Excel数据筛选工具.exe"
else
    echo -e "${RED}✗${NC} 打包失败，请检查错误信息"
    exit 1
fi
