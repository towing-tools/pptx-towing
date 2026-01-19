#!/bin/bash
#
# build.sh - PPTX生成スクリプト（1コマンド実行）
#
# 使い方:
#   ./build.sh <project_name>
#
# 例:
#   ./build.sh 2026_business_plan
#
# 前提条件:
#   - Node.js がインストールされていること
#   - Python 3 がインストールされていること
#   - npm install が実行済みであること
#   - output/<project_name>/content.md が存在すること
#

set -e  # エラー時に即座に終了

# カラー出力
RED='\033[0;31m'
GREEN='\033[0;32m'
YELLOW='\033[1;33m'
BLUE='\033[0;34m'
NC='\033[0m' # No Color

# スクリプトのディレクトリを取得
SCRIPT_DIR="$(cd "$(dirname "${BASH_SOURCE[0]}")" && pwd)"
cd "$SCRIPT_DIR"

# 引数チェック
if [ -z "$1" ]; then
    echo -e "${RED}エラー: プロジェクト名を指定してください${NC}"
    echo ""
    echo "使い方: ./build.sh <project_name>"
    echo ""
    echo "例:"
    echo "  ./build.sh 2026_business_plan"
    echo "  ./build.sh investor_pitch"
    echo ""
    echo "利用可能なプロジェクト:"
    if [ -d "output" ]; then
        ls -1 output/ 2>/dev/null | grep -v "^$" | sed 's/^/  - /'
    else
        echo "  (outputディレクトリが存在しません)"
    fi
    exit 1
fi

PROJECT=$1
PROJECT_DIR="output/$PROJECT"
CONTENT_FILE="$PROJECT_DIR/content.md"
SLIDES_DIR="$PROJECT_DIR/slides"
INTERMEDIATE_FILE="$PROJECT_DIR/intermediate.pptx"
FINAL_FILE="$PROJECT_DIR/final.pptx"
TEMPLATE_FILE="assets/2026TOWINGtemplate.pptx"

echo -e "${BLUE}========================================${NC}"
echo -e "${BLUE}  TOWING PPTX Generator${NC}"
echo -e "${BLUE}========================================${NC}"
echo ""

# ===== 事前チェック =====

echo -e "${YELLOW}[1/5] 事前チェック...${NC}"

# content.mdの存在確認
if [ ! -f "$CONTENT_FILE" ]; then
    echo -e "${RED}エラー: $CONTENT_FILE が見つかりません${NC}"
    echo ""
    echo "content.mdを作成してから再実行してください。"
    echo "参照: docs/procedure.md"
    exit 1
fi

# テンプレートの存在確認
if [ ! -f "$TEMPLATE_FILE" ]; then
    echo -e "${RED}エラー: $TEMPLATE_FILE が見つかりません${NC}"
    exit 1
fi

# node_modulesの存在確認
if [ ! -d "node_modules" ]; then
    echo -e "${YELLOW}node_modulesが見つかりません。npm installを実行します...${NC}"
    npm install
fi

# slidesディレクトリの作成
mkdir -p "$SLIDES_DIR"

echo -e "${GREEN}  ✓ 事前チェック完了${NC}"
echo ""

# ===== Phase 3: Markdown → 中間PPTX =====

echo -e "${YELLOW}[2/5] Markdown → HTML変換...${NC}"
echo "  入力: $CONTENT_FILE"

# 古い中間ファイルを削除
rm -f "$SLIDES_DIR"/*.html 2>/dev/null || true
rm -f "$INTERMEDIATE_FILE" 2>/dev/null || true

echo -e "${YELLOW}[3/5] HTML → PPTX変換（Playwright使用）...${NC}"
echo "  出力: $INTERMEDIATE_FILE"

node scripts/generate_from_md.js "$CONTENT_FILE" "$INTERMEDIATE_FILE"

if [ ! -f "$INTERMEDIATE_FILE" ]; then
    echo -e "${RED}エラー: 中間PPTXの生成に失敗しました${NC}"
    exit 1
fi

echo -e "${GREEN}  ✓ 中間PPTX生成完了${NC}"
echo ""

# ===== Phase 4: テンプレート注入 =====

echo -e "${YELLOW}[4/5] テンプレート注入...${NC}"
echo "  テンプレート: $TEMPLATE_FILE"
echo "  出力: $FINAL_FILE"

python3 scripts/inject_slides.py \
    "$TEMPLATE_FILE" \
    "$INTERMEDIATE_FILE" \
    "$FINAL_FILE"

if [ ! -f "$FINAL_FILE" ]; then
    echo -e "${RED}エラー: 最終PPTXの生成に失敗しました${NC}"
    exit 1
fi

echo -e "${GREEN}  ✓ テンプレート注入完了${NC}"
echo ""

# ===== 完了 =====

echo -e "${YELLOW}[5/5] 完了処理...${NC}"

# ファイルサイズを表示
FINAL_SIZE=$(ls -lh "$FINAL_FILE" | awk '{print $5}')
SLIDE_COUNT=$(ls -1 "$SLIDES_DIR"/*.html 2>/dev/null | wc -l | tr -d ' ')

echo ""
echo -e "${GREEN}========================================${NC}"
echo -e "${GREEN}  生成完了！${NC}"
echo -e "${GREEN}========================================${NC}"
echo ""
echo -e "  プロジェクト: ${BLUE}$PROJECT${NC}"
echo -e "  スライド数:   ${BLUE}${SLIDE_COUNT}枚${NC}"
echo -e "  ファイルサイズ: ${BLUE}$FINAL_SIZE${NC}"
echo ""
echo -e "  ${GREEN}出力ファイル:${NC}"
echo -e "    $FINAL_FILE"
echo ""
echo -e "  ${YELLOW}中間ファイル:${NC}"
echo -e "    $SLIDES_DIR/ (HTML)"
echo -e "    $INTERMEDIATE_FILE"
echo ""

# macOSの場合、ファイルを開くか確認
if [ "$(uname)" = "Darwin" ]; then
    read -p "PowerPointで開きますか？ (y/N): " -n 1 -r
    echo
    if [[ $REPLY =~ ^[Yy]$ ]]; then
        open "$FINAL_FILE"
    fi
fi
