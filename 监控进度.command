#!/bin/bash
# Wine安装进度监控

GREEN='\033[0;32m'
BLUE='\033[0;34m'
YELLOW='\033[1;33m'
CYAN='\033[0;36m'
NC='\033[0m'

clear
echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}  Wine安装进度监控${NC}"
echo -e "${BLUE}========================================${NC}"
echo ""

DOWNLOAD_FILE="$HOME/Library/Caches/Homebrew/downloads/*wine*.tar.xz.incomplete"

while true; do
    clear
    echo -e "${BLUE}========================================${NC}"
    echo -e "${BLUE}  Wine安装进度监控${NC}"
    echo -e "${BLUE}========================================${NC}"
    echo ""

    # 检查Wine是否已安装
    if command -v wine &> /dev/null 2>&1; then
        WINE_VER=$(wine --version)
        echo -e "${GREEN}✓ Wine已安装完成！${NC}"
        echo "版本: $WINE_VER"
        echo ""
        echo -e "${GREEN}========================================${NC}"
        echo -e "${GREEN}  现在可以开始打包exe了！${NC}"
        echo -e "${GREEN}========================================${NC}"
        echo ""
        echo "请关闭此窗口，然后双击运行："
        echo -e "${CYAN}mac自动打包exe.sh${NC}"
        echo ""
        break
    fi

    # 检查下载文件
    if ls $DOWNLOAD_FILE 2> /dev/null | head -1 | grep -q .; then
        FILE=$(ls $DOWNLOAD_FILE 2>/dev/null | head -1)
        SIZE=$(ls -lh "$FILE" 2>/dev/null | awk '{print $5}')
        SIZE_BYTES=$(ls -l "$FILE" 2>/dev/null | awk '{print $5}')
        TIMESTAMP=$(stat -f "%Sm" -t "%H:%M:%S" "$FILE" 2>/dev/null || stat -c "%y" "$FILE" 2>/dev/null | cut -d'.' -f1)

        echo -e "${YELLOW}⏳ Wine正在下载中...${NC}"
        echo ""
        echo "当前文件大小: ${CYAN}$SIZE${NC}"
        echo "目标大小: 约 200MB"
        echo "最后更新: $TIMESTAMP"
        echo ""

        # 估算进度
        if [ -n "$SIZE_BYTES" ]; then
            # 简单计算进度
            PERCENT=$((SIZE_BYTES * 100 / 200000000))
            if [ $PERCENT -lt 100 ]; then
                echo -e "${BLUE}进度: [$(printf '%*s' $((PERCENT/2)) | tr ' ' '=')$(printf '%*s' $((50-PERCENT/2)) | tr ' ' '-')] ${PERCENT}%${NC}"
            fi
        fi
        echo ""

        # 检查进程
        if ps aux | grep -E "curl.*wine|brew.*wine" | grep -v grep | grep -q .; then
            echo -e "${GREEN}✓${NC} 下载进程运行中"
        else
            echo -e "${YELLOW}⚠${NC} 下载暂停，等待恢复..."
        fi

        echo ""
        echo -e "${CYAN}预计剩余时间: 5-10分钟${NC}"

    else
        echo -e "${YELLOW}正在准备下载...${NC}"
        echo ""
        echo "检查Homebrew状态..."
    fi

    echo ""
    echo -e "${CYAN}当前时间: $(date '+%H:%M:%S')${NC}"
    echo ""
    echo -e "${YELLOW}提示: Ctrl+C 退出监控${NC}"
    echo ""

    sleep 3
done
