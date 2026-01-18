# pptx-towing-v2 CHANGELOG

## [2026-01-16] - Column Layout Fix & PDF Conversion

### Fixed
- **Column Layout Parser**: `::: column` 構文が正しく認識されるようになった
  - 問題: `::: column` がテキストとしてそのまま出力されていた
  - 原因: `generate_from_md.js` のパーサーが `::: column` を認識していなかった
  - 修正: `parseMarkdown()` 関数に `::: column` の検出ロジックを追加
  - `columnIndex` 変数で左右カラムのインデックスを追跡
  - カラム比率（例: 50/50, 60/40）が正しく適用されるように

### Changed
- **CSS Layout**: `justify-content: center` に変更し、コンテンツを垂直中央寄せに
- **Gap調整**: 要素間の余白を `10pt` → `12pt` に拡大

### Documentation
- **MANDATORY_GUIDE_FOR_PPTX_GENERATION.md**: カラムレイアウトの記述ルールを追加
  - 基本構文（必須パターン）
  - 重要ルール（閉じタグの順序、空行の入れ方）
  - カラム内にボックスを入れる場合の構文
  - よく使うレイアウトパターン（テンプレート集）3種類
- **スライドコンテンツ設計ガイドライン（セクション3.5）を追加**
  - レイアウト選択の原則（コンテンツ量に応じた推奨レイアウト表）
  - 最小コンテンツ量の基準（NG例とOK例）
  - シングルカラム/2カラム/ヘッダー+2カラムの充実テンプレート
  - コンテンツが足りない場合の対処法（具体的情報・背景・補足の追加方法）

### Added
- **PDF Conversion Workflow**: Documented the reliable workflow for generating slide thumbnails.

### Notes: PDF Conversion for Thumbnails
- **中間ファイル（temp_*.pptx）はLibreOffice（soffice）でPDF変換可能**
- **inject_slides.py適用後のファイルはPDF変換に失敗する場合がある**
- サムネイル生成には中間ファイルを使用することを推奨

**推奨ワークフロー**:
```bash
# 1. Markdownから中間PPTXを生成
node scripts/generate_from_md.js input.md temp_output.pptx

# 2. 中間ファイルをPDFに変換（サムネイル用）
cp temp_output.pptx /tmp/
cd /tmp && soffice --headless --convert-to pdf temp_output.pptx

# 3. PDFからサムネイル画像を生成
pdftoppm -jpeg -r 150 temp_output.pdf slide

# 4. テンプレート注入（最終成果物）
python3 scripts/inject_slides.py assets/2026TOWINGtemplate.pptx temp_output.pptx final_output.pptx
```

### Known Issues
- iCloud Driveパス上のファイルは直接PDF変換できない場合がある（/tmpにコピーして実行）
- inject_slides.py適用後のファイルはXML構造の変更によりLibreOfficeでの読み込みに失敗することがある

---

## [2026-01-15] - Layout Engine Overhaul & Parser Fixes

### Added
- **Multi-column Layout Support**: Introduced `::: columns [left]/[right]` syntax (e.g., `50/50`, `40/60`).
- **Visual Validation Tools**: Added `scripts/thumbnail.py` using `pdf2image` + `Keynote (AppleScript)` pipeline for reliable macOS thumbnail generation.
- **Direct Thumbnail Extraction**: Added `scripts/extract_thumb.py` for checking embedded thumbnails.

### Changed
- **Markdown Parser Logic**:
  - Rewrote the parser state machine to handle nested containers correctly (e.g., `box` inside `column`).
  - Fixed stack management to prevent premature closing of parent containers.
- **CSS / Layout Engine**:
  - **Flexbox Ratio**: Implemented dynamic `flex-grow` based on column ratio syntax.
  - **Text Wrapping**: Added `min-width: 0`, `flex-basis: 0`, and `overflow-wrap: break-word` to force text wrapping within columns and prevent horizontal overflow.
  - **Spacing**: Adjusted content `gap` (10pt) and margin to optimize vertical space usage.
  - **Reset**: Added CSS reset for `h1`, `h2`, `p` to remove browser default margins causing vertical misalignments.
- **Validation Policy**:
  - Downgraded `html2pptx` overflow errors from **Fatal** to **Warning**. The generation process no longer halts on text overflow, prioritizing output generation over strict constraint enforcement.
- **Typography**:
  - Reduced body font size from `14pt` to `12.5pt` to accommodate more content.
  - Removed deprecated `.header-bar` visual element.

### Known Issues
- **Vertical Overflow**: Despite font resizing, large amounts of text still overflow the slide boundaries. The current solution relies on manual Markdown trimming or ignoring warnings.
- **Layout Robustness**: Complex nesting or extreme aspect ratios may still cause rendering artifacts in PptxGenJS conversion.