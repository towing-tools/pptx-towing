#
# build.ps1 - PPTX生成スクリプト（Windows PowerShell版）
#
# 使い方:
#   .\build.ps1 <project_name>
#
# 例:
#   .\build.ps1 2026_business_plan
#
# 前提条件:
#   - Node.js がインストールされていること
#   - Python 3 がインストールされていること
#   - npm install が実行済みであること
#   - output/<project_name>/content.md が存在すること
#
# 注意:
#   Windows環境での動作は十分にテストされていません。
#   問題があれば報告してください。
#

param(
    [Parameter(Position=0)]
    [string]$ProjectName
)

$ErrorActionPreference = "Stop"

# スクリプトのディレクトリに移動
$ScriptDir = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $ScriptDir

# 引数チェック
if ([string]::IsNullOrEmpty($ProjectName)) {
    Write-Host "エラー: プロジェクト名を指定してください" -ForegroundColor Red
    Write-Host ""
    Write-Host "使い方: .\build.ps1 <project_name>"
    Write-Host ""
    Write-Host "例:"
    Write-Host "  .\build.ps1 2026_business_plan"
    Write-Host "  .\build.ps1 investor_pitch"
    Write-Host ""
    Write-Host "利用可能なプロジェクト:"
    if (Test-Path "output") {
        Get-ChildItem -Path "output" -Directory | ForEach-Object { Write-Host "  - $($_.Name)" }
    } else {
        Write-Host "  (outputディレクトリが存在しません)"
    }
    exit 1
}

$Project = $ProjectName
$ProjectDir = "output\$Project"
$ContentFile = "$ProjectDir\content.md"
$SlidesDir = "$ProjectDir\slides"
$IntermediateFile = "$ProjectDir\intermediate.pptx"
$FinalFile = "$ProjectDir\final.pptx"
$TemplateFile = "assets\2026TOWINGtemplate.pptx"

Write-Host "========================================" -ForegroundColor Blue
Write-Host "  TOWING PPTX Generator (Windows)" -ForegroundColor Blue
Write-Host "========================================" -ForegroundColor Blue
Write-Host ""

# ===== 事前チェック =====

Write-Host "[1/5] 事前チェック..." -ForegroundColor Yellow

# content.mdの存在確認
if (-not (Test-Path $ContentFile)) {
    Write-Host "エラー: $ContentFile が見つかりません" -ForegroundColor Red
    Write-Host ""
    Write-Host "content.mdを作成してから再実行してください。"
    Write-Host "参照: docs\procedure.md"
    exit 1
}

# テンプレートの存在確認
if (-not (Test-Path $TemplateFile)) {
    Write-Host "エラー: $TemplateFile が見つかりません" -ForegroundColor Red
    exit 1
}

# node_modulesの存在確認
if (-not (Test-Path "node_modules")) {
    Write-Host "node_modulesが見つかりません。npm installを実行します..." -ForegroundColor Yellow
    npm install
}

# slidesディレクトリの作成
if (-not (Test-Path $SlidesDir)) {
    New-Item -ItemType Directory -Path $SlidesDir -Force | Out-Null
}

Write-Host "  ✓ 事前チェック完了" -ForegroundColor Green
Write-Host ""

# ===== Phase 3: Markdown → 中間PPTX =====

Write-Host "[2/5] Markdown → HTML変換..." -ForegroundColor Yellow
Write-Host "  入力: $ContentFile"

# 古い中間ファイルを削除
Get-ChildItem -Path $SlidesDir -Filter "*.html" -ErrorAction SilentlyContinue | Remove-Item -Force
if (Test-Path $IntermediateFile) { Remove-Item $IntermediateFile -Force }

Write-Host "[3/5] HTML → PPTX変換（Playwright使用）..." -ForegroundColor Yellow
Write-Host "  出力: $IntermediateFile"

node scripts/generate_from_md.js $ContentFile $IntermediateFile

if (-not (Test-Path $IntermediateFile)) {
    Write-Host "エラー: 中間PPTXの生成に失敗しました" -ForegroundColor Red
    exit 1
}

Write-Host "  ✓ 中間PPTX生成完了" -ForegroundColor Green
Write-Host ""

# ===== Phase 4: テンプレート注入 =====

Write-Host "[4/5] テンプレート注入..." -ForegroundColor Yellow
Write-Host "  テンプレート: $TemplateFile"
Write-Host "  出力: $FinalFile"

python scripts/inject_slides.py $TemplateFile $IntermediateFile $FinalFile

if (-not (Test-Path $FinalFile)) {
    Write-Host "エラー: 最終PPTXの生成に失敗しました" -ForegroundColor Red
    exit 1
}

Write-Host "  ✓ テンプレート注入完了" -ForegroundColor Green
Write-Host ""

# ===== 完了 =====

Write-Host "[5/5] 完了処理..." -ForegroundColor Yellow

# ファイルサイズを取得
$FinalSize = (Get-Item $FinalFile).Length
$FinalSizeFormatted = "{0:N2} KB" -f ($FinalSize / 1KB)
$SlideCount = (Get-ChildItem -Path $SlidesDir -Filter "*.html" | Measure-Object).Count

Write-Host ""
Write-Host "========================================" -ForegroundColor Green
Write-Host "  生成完了！" -ForegroundColor Green
Write-Host "========================================" -ForegroundColor Green
Write-Host ""
Write-Host "  プロジェクト: $Project" -ForegroundColor Cyan
Write-Host "  スライド数:   ${SlideCount}枚" -ForegroundColor Cyan
Write-Host "  ファイルサイズ: $FinalSizeFormatted" -ForegroundColor Cyan
Write-Host ""
Write-Host "  出力ファイル:" -ForegroundColor Green
Write-Host "    $FinalFile"
Write-Host ""
Write-Host "  中間ファイル:" -ForegroundColor Yellow
Write-Host "    $SlidesDir\ (HTML)"
Write-Host "    $IntermediateFile"
Write-Host ""

# Windowsの場合、ファイルを開くか確認
$response = Read-Host "PowerPointで開きますか？ (y/N)"
if ($response -eq "y" -or $response -eq "Y") {
    Start-Process $FinalFile
}
