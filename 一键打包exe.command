#!/bin/bash
# 一键打包脚本 - 自动上传到GitHub并生成Windows exe

set -e

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m' # No Color

# 切换到脚本所在目录
SCRIPT_DIR="$(cd "$(dirname "$0")" && pwd)"
cd "$SCRIPT_DIR"

clear
echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}   Excel数据筛选工具 - 一键打包${NC}"
echo -e "${BLUE}========================================${NC}"
echo ""

# 检查是否登录GitHub
echo -e "${YELLOW}[1/6]${NC} 检查GitHub登录状态..."
if ! gh auth status &>/dev/null; then
    echo "未登录GitHub，正在打开浏览器进行登录..."
    gh auth login
    echo ""
fi

echo -e "${GREEN}✓${NC} 已登录GitHub"
echo ""

# 获取仓库名称
echo -e "${YELLOW}[2/6]${NC} 设置仓库信息"
read -p "请输入仓库名称 (直接回车使用默认: excel-filter-tool): " REPO_NAME
REPO_NAME=${REPO_NAME:-excel-filter-tool}

read -p "请输入GitHub用户名: " GITHUB_USER
if [ -z "$GITHUB_USER" ]; then
    GITHUB_USER=$(gh api user --jq '.login')
fi

echo "仓库: $GITHUB_USER/$REPO_NAME"
echo ""

# 初始化git仓库
echo -e "${YELLOW}[3/6]${NC} 初始化本地仓库..."
git init 2>/dev/null || true
git branch -M main 2>/dev/null || true
echo -e "${GREEN}✓${NC} Git仓库初始化完成"
echo ""

# 创建GitHub仓库
echo -e "${YELLOW}[4/6]${NC} 创建GitHub仓库..."
if gh repo view "$GITHUB_USER/$REPO_NAME" &>/dev/null; then
    echo -e "${GREEN}✓${NC} 仓库已存在"
else
    echo "正在创建新仓库..."
    gh repo create "$REPO_NAME" --public
    echo -e "${GREEN}✓${NC} 仓库创建成功"
fi
echo ""

# 准备文件
echo -e "${YELLOW}[5/6]${NC} 准备文件..."
mkdir -p .github/workflows
cp .github/workflows/build.yml .github/workflows/build.yml 2>/dev/null || true
git add excel_filter_tool.py .github/workflows/build.yml 2>/dev/null || git add excel_filter_tool.py
git add requirements.txt README.md 2>/dev/null || true

if git diff --staged --quiet; then
    echo "没有需要提交的更改"
else
    git commit -m "Add Excel filter tool" || true
fi

git branch -M main 2>/dev/null || true

# 推送到GitHub
echo -e "${YELLOW}[6/6]${NC} 上传到GitHub..."
git remote remove origin 2>/dev/null || true
git remote add origin "https://github.com/$GITHUB_USER/$REPO_NAME.git"
git push -u origin main 2>/dev/null || git push -u origin main --force
echo -e "${GREEN}✓${NC} 文件上传完成"
echo ""

# 触发构建
echo -e "${YELLOW}正在触发自动构建..."
gh workflow run build.yml --repo "$GITHUB_USER/$REPO_NAME"
echo -e "${GREEN}✓${NC} 构建已触发"
echo ""

# 等待构建
echo "等待构建完成（约1-2分钟）..."
echo "你可以按 Ctrl+C 退出，稍后手动查看"
echo ""

sleep 5

# 监控构建状态
for i in {1..30}; do
    clear
    echo -e "${BLUE}========================================${NC}"
    echo -e "${BLUE}   正在构建中... ($i/30)${NC}"
    echo -e "${BLUE}========================================${NC}"
    echo ""

    RUN_ID=$(gh run list --repo "$GITHUB_USER/$REPO_NAME" --limit 1 --json databaseId --jq '.[0].databaseId')

    if [ -n "$RUN_ID" ]; then
        STATUS=$(gh run view "$RUN_ID" --repo "$GITHUB_USER/$REPO_NAME" --json status --jq '.status')
        CONCLUSION=$(gh run view "$RUN_ID" --repo "$GITHUB_USER/$REPO_NAME" --json conclusion --jq '.conclusion')

        echo "构建状态: $STATUS"

        if [ "$CONCLUSION" = "success" ]; then
            echo ""
            echo -e "${GREEN}========================================${NC}"
            echo -e "${GREEN}   ✓ 构建成功！${NC}"
            echo -e "${GREEN}========================================${NC}"
            echo ""
            echo "正在下载exe文件..."

            # 下载构建产物
            gh run download "$RUN_ID" --repo "$GITHUB_USER/$REPO_NAME" --name "Excel数据筛选工具-Windows" --dir .

            if [ -f "Excel数据筛选工具.exe" ]; then
                echo ""
                echo -e "${GREEN}✓${NC} 下载完成！"
                echo ""
                echo "文件位置: $(pwd)/Excel数据筛选工具.exe"
                ls -lh "Excel数据筛选工具.exe" | awk '{print "文件大小: " $5}'
                echo ""
                echo -e "${GREEN}========================================${NC}"
                echo -e "${GREEN}   大功告成！现在可以将exe文件${NC}"
                echo -e "${GREEN}   拷贝到Windows电脑使用了！${NC}"
                echo -e "${GREEN}========================================${NC}"

                # 打开finder
                open .
                break
            fi
        elif [ "$CONCLUSION" = "failure" ]; then
            echo ""
            echo -e "${YELLOW}构建失败${NC}"
            echo "请查看: https://github.com/$GITHUB_USER/$REPO_NAME/actions"
            break
        fi
    fi

    sleep 4
done

echo ""
echo "如果构建还在进行中，可以稍后访问以下链接下载："
echo "https://github.com/$GITHUB_USER/$REPO_NAME/actions"
echo ""
