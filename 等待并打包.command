#!/bin/bash
# 自动等待Wine安装完成后开始打包

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
NC='\033[0m'

clear
echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}  等待Wine安装完成...${NC}"
echo -e "${BLUE}========================================${NC}"
echo ""

cd /Users/songjinfeng/Desktop/shuju

# 等待Wine安装完成
while true; do
    if command -v wine &> /dev/null; then
        echo -e "${GREEN}✓${NC} Wine安装完成！"
        sleep 2
        echo ""
        echo "开始自动打包流程..."
        echo ""
        sleep 2
        exec bash mac自动打包exe.sh
    fi

    echo -ne "${YELLOW}等待中...${NC} $(date '+%H:%M:%S')\\r"
    sleep 5
done
