#!/bin/bash
# MCP SQL Server 資料庫切換腳本
# 用途: 快速切換 MCP Server 連接的資料庫

SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 顏色定義
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}MCP SQL Server 資料庫切換工具${NC}"
echo -e "${BLUE}========================================${NC}"
echo

# 檢查可用的環境配置檔案
ENV_FILES=()
for file in .env.*; do
    if [ -f "$file" ]; then
        ENV_FILES+=("$file")
    fi
done

if [ ${#ENV_FILES[@]} -eq 0 ]; then
    echo -e "${RED}錯誤: 未找到任何 .env.* 配置檔案${NC}"
    exit 1
fi

# 顯示當前連接
if [ -f .env ]; then
    CURRENT_DB=$(grep "^DB_NAME=" .env | cut -d'=' -f2)
    CURRENT_SERVER=$(grep "^DB_SERVER=" .env | cut -d'=' -f2)
    echo -e "${GREEN}目前連接:${NC} $CURRENT_SERVER - $CURRENT_DB"
    echo
fi

# 顯示可用的資料庫配置
echo -e "${YELLOW}可用的資料庫配置:${NC}"
echo
for i in "${!ENV_FILES[@]}"; do
    file="${ENV_FILES[$i]}"
    db_name=$(grep "^DB_NAME=" "$file" | cut -d'=' -f2)
    db_server=$(grep "^DB_SERVER=" "$file" | cut -d'=' -f2)
    db_port=$(grep "^DB_PORT=" "$file" | cut -d'=' -f2)

    config_name="${file#.env.}"
    echo -e "  ${GREEN}[$((i+1))]${NC} $config_name"
    echo -e "      伺服器: $db_server:$db_port"
    echo -e "      資料庫: $db_name"
    echo
done

# 如果提供了參數，直接切換
if [ -n "$1" ]; then
    CHOICE=$1
else
    # 讀取用戶選擇
    echo -n "請選擇要切換的資料庫 [1-${#ENV_FILES[@]}]: "
    read CHOICE
fi

# 驗證選擇
if ! [[ "$CHOICE" =~ ^[0-9]+$ ]] || [ "$CHOICE" -lt 1 ] || [ "$CHOICE" -gt ${#ENV_FILES[@]} ]; then
    echo -e "${RED}錯誤: 無效的選擇${NC}"
    exit 1
fi

# 執行切換
SELECTED_FILE="${ENV_FILES[$((CHOICE-1))]}"
cp "$SELECTED_FILE" .env

NEW_DB=$(grep "^DB_NAME=" .env | cut -d'=' -f2)
NEW_SERVER=$(grep "^DB_SERVER=" .env | cut -d'=' -f2)

echo
echo -e "${GREEN}✓ 已成功切換到: $NEW_SERVER - $NEW_DB${NC}"
echo
echo -e "${YELLOW}注意: 請重新啟動 MCP Server 以使更改生效${NC}"
echo -e "${YELLOW}重啟指令: claude mcp restart${NC}"
echo
