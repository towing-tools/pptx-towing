# MANDATORY GUIDE FOR PPTX GENERATION v3.0

**重要**: このドキュメントは再構成されました。

---

## 新しいドキュメント構造

スライド生成のガイドは `docs/` ディレクトリに移動し、段階的に参照できるように分割されました。

### 必須参照（最初に読む）

**📄 `docs/procedure.md`** - メイン手順書
- 作業手順、核心原則、フォルダ構造ルール
- **スライド生成時は必ずこのファイルから開始してください**

### 必要に応じて参照

| ファイル | 内容 | 参照タイミング |
|---------|------|---------------|
| `docs/templates.md` | スライドテンプレート集 | Phase 1-2: Outline/Content作成時 |
| `docs/content-guide.md` | コンテンツ設計ガイド | Phase 2: Content作成時 |
| `docs/config-reference.md` | theme_config.json詳細 | デザイン調整が必要な時 |
| `docs/troubleshooting.md` | エラー対処法 | エラー発生時 |

---

## 【重要】Sandbox環境での制限事項

**Gemini sandbox等の制限された環境では、PPTX生成スクリプトが動作しません。**

### 理由

- PPTX生成にはPlaywright（ブラウザ自動化ツール）が必要
- sandbox環境ではOSレベルのライブラリが不足しており、インストール不可

### 解決策：役割分担

| 環境 | 担当 | 成果物 |
|------|------|--------|
| AIアシスタント（sandbox） | Phase 1-2: content.md作成 | `output/<project>/content.md` |
| ユーザー（ローカル） | Phase 3-4: PPTX生成 | `output/<project>/final.pptx` |

### AIアシスタントへの指示

**content.md作成完了時、必ず以下をユーザーに伝えてください：**

```
content.mdの作成が完了しました。

PPTXを生成するには、ローカル環境のターミナルで以下のコマンドを実行してください：

cd /path/to/pptx-towing-v3
./build.sh <project_name>

【重要】sandbox環境でnpm installを実行した場合、node_modulesに互換性がありません。
以下のコマンドで再インストールしてから実行してください：

rm -rf node_modules package-lock.json
npm install
./build.sh <project_name>
```

---

## クイックスタート

### 1. 専用フォルダを作成（必須）

```bash
mkdir -p pptx-towing-v3/output/<project_name>/slides
```

### 2. procedure.mdを読む

```
docs/procedure.md を開いて手順に従う
```

### 3. content.md を作成

AIアシスタントがPhase 1-2を実行し、`output/<project_name>/content.md`を作成

### 4. スライド生成（1コマンド）

**ローカル環境で実行：**

```bash
cd pptx-towing-v3
./build.sh <project_name>

# 例
./build.sh 2026_business_plan
```

`build.sh` は以下を自動実行：
- 事前チェック（ファイル存在確認、npm install等）
- Phase 3: Markdown → 中間PPTX（Playwright使用）
- Phase 4: テンプレート注入
- 完了メッセージ表示

### 手動実行する場合

```bash
# Phase 3: 中間PPTX生成
node scripts/generate_from_md.js \
  output/<project_name>/content.md \
  output/<project_name>/intermediate.pptx

# Phase 4: テンプレート注入
python3 scripts/inject_slides.py \
  assets/2026TOWINGtemplate.pptx \
  output/<project_name>/intermediate.pptx \
  output/<project_name>/final.pptx
```

---

## 重要な変更点（v3.0）

### 1. フォルダ構造の義務化
- スライド生成時は必ず `output/<project_name>/` フォルダを作成
- 中間ファイル（HTML, MD, PPTX）はすべてそのフォルダ内に保存
- pptx-towing-v3 直下へのファイル散乱は禁止

### 2. 言語指定の追加
- 日本語/英語/併記をユーザーに確認
- メタデータで `language`, `bilingual` を指定

### 3. メッセージ設計の強化
- 各スライドに必ず「伝えたいメッセージ」を含める
- `primary` または `highlight` ボックスで強調

### 4. コンテンツ量の基準
- スライド面積の70-80%を使用
- 箇条書きは最低5項目（MEDIUMの場合）
- 無駄な空白を避ける

---

## 禁止事項

- ❌ CSS/HTMLの直接記述
- ❌ 座標のハードコード
- ❌ theme_config.jsonの無断改変
- ❌ pptx-towing-v3直下への中間ファイル作成

---

## 旧ガイドからの移行

旧 `MANDATORY_GUIDE_FOR_PPTX_GENERATION.md` の内容は以下のように分割されました：

- セクション1-2（原則、Config） → `docs/procedure.md`, `docs/config-reference.md`
- セクション3-4（Markdown、テンプレート） → `docs/templates.md`
- セクション5（コンテンツ設計） → `docs/content-guide.md`
- セクション6（パイプライン） → `docs/procedure.md`
- セクション7-9（チェックリスト、禁止事項、トラブル） → `docs/content-guide.md`, `docs/troubleshooting.md`

---

**次のステップ**: `docs/procedure.md` を開いて作業を開始してください。
