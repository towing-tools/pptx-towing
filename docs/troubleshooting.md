# トラブルシューティング (Troubleshooting)

**このファイルはエラー発生時に参照してください。**

---

## 一般的なエラー

### 「Text box ends too close to bottom edge」

**原因**: コンテンツが多すぎてスライドからはみ出している

**解決策**:
1. Markdownの文章を減らす
2. 箇条書きの項目数を減らす（最大6項目推奨）
3. ボックス内のテキストを短くする（100字以下推奨）
4. スライドを分割する

```markdown
# 変更前（多すぎる）
* 項目1
* 項目2
* 項目3
* 項目4
* 項目5
* 項目6
* 項目7
* 項目8

# 変更後（適切）
* 項目1
* 項目2
* 項目3
* 項目4
* 項目5
```

---

### 「Unwrapped text」

**原因**: テキストがHTMLタグで囲まれていない

**解決策**:
- Markdownの構文を確認
- 空行が適切に入っているか確認
- ボックスの閉じタグ（`:::`）が正しいか確認

```markdown
# 正しい例
::: box
**タイトル**
* 項目1
* 項目2
:::

# 間違い例（閉じタグがない）
::: box
**タイトル**
* 項目1
* 項目2
```

---

### 色がtheme_config.jsonと違う

**原因**: インラインスタイルが指定されている

**解決策**:
- Markdownに色指定を直接書かない
- `theme_config.json`の設定を使用

```markdown
# NG（色を直接指定）
<span style="color: red">テキスト</span>

# OK（ボックススタイルを使用）
::: box warning
テキスト
:::
```

---

## カラムレイアウトのエラー

### カラムが表示されない

**原因**: 閉じタグの順序が間違っている

**正しい構文**:
```markdown
::: columns 50/50

::: column
左カラムの内容
:::

::: column
右カラムの内容
:::

:::
```

**間違いやすいパターン**:
```markdown
# NG: columnの後に改行がない
::: columns 50/50
::: column
内容
:::
:::

# NG: 閉じタグが足りない
::: columns 50/50
::: column
内容
::: column
内容
:::
```

---

### カラム内のボックスが崩れる

**正しい構文**:
```markdown
::: columns 50/50

::: column
::: box
ボックス内容
:::
:::

::: column
::: box highlight
ボックス内容
:::
:::

:::
```

**ポイント**:
- box → column → columns の順で閉じる
- 各閉じタグ（`:::`）は独立した行に

---

## 生成スクリプトのエラー

### node scripts/generate_from_md.js が失敗

**確認事項**:
1. Node.jsがインストールされているか
2. 依存関係がインストールされているか

```bash
cd pptx-towing-v3
npm install
```

3. 入力ファイルパスが正しいか
4. 出力ディレクトリが存在するか

---

### python3 scripts/inject_slides.py が失敗

**確認事項**:
1. Python 3がインストールされているか
2. 入力ファイル（テンプレート、中間PPTX）が存在するか
3. ファイルパスにスペースや特殊文字がないか

```bash
# パスを引用符で囲む
python3 scripts/inject_slides.py \
  "assets/2026TOWINGtemplate.pptx" \
  "output/project/intermediate.pptx" \
  "output/project/final.pptx"
```

---

## デザイン品質の問題

### コントラスト比が低い

**確認方法**:
1. WebAIM Contrast Checker にアクセス
2. 背景色と文字色を入力
3. WCAG AA (4.5:1) 以上か確認

**対処法**:
- 明るい背景には暗い文字
- 暗い背景には明るい文字
- primaryボックスは白文字で自動的にコントラスト確保

---

### スライドの上下が不揃い

**原因**: content_densityが統一されていない

**対処法**:
1. 各スライドのメタデータで`content_density`を明示
2. LOWスライドは意図的に少なくする
3. MEDIUMスライドは5-6項目を目安に

---

### 空白が多すぎる

**対処法**:
1. 箇条書きの項目を増やす（最低5項目）
2. より詳細な説明を追加
3. 画像のプレースホルダーを配置
4. ボックスを追加

```markdown
::: box
**[画像: グラフの説明]**
- 種類: 棒グラフ
- データ: 売上推移
:::
```

---

## ファイル構造のエラー

### 中間ファイルが散らかっている

**対処法**:
必ず専用フォルダを作成してから作業

```bash
mkdir -p pptx-towing-v3/output/<project_name>/slides

# 生成コマンドでは出力先を指定
node scripts/generate_from_md.js \
  output/<project_name>/content.md \
  output/<project_name>/intermediate.pptx
```

---

### 古いファイルが残っている

**対処法**:
プロジェクトフォルダごと削除して再生成

```bash
rm -rf pptx-towing-v3/output/<project_name>
mkdir -p pptx-towing-v3/output/<project_name>/slides
```

---

## よくある質問

### Q: 日本語と英語を併記するには？

**A**: メタデータで指定し、本文で併記

```markdown
***
language: "ja-en"
bilingual: true
***

# 事業概要 (Business Overview)

* 売上増加 / Revenue growth
* 市場拡大 / Market expansion
```

---

### Q: 画像を追加するには？

**A**: 現在のパイプラインでは画像は自動挿入されません。

代替案:
1. プレースホルダーを記載
2. 最終PPTXを開いて手動で画像を挿入

```markdown
::: box
**[画像を挿入: assets/images/logo.png]**
:::
```

---

### Q: 表を追加するには？

**A**: Markdown形式で記述（ただし生成時の対応は限定的）

```markdown
| 項目 | 値 | 備考 |
|------|-----|------|
| 売上 | 100億 | 前年比+10% |
| 利益 | 10億 | 前年比+5% |
```

複雑な表は最終PPTXで手動編集を推奨

---

## サポート

問題が解決しない場合:
1. エラーメッセージを確認
2. 入力Markdownの構文を再確認
3. `procedure.md` の手順に従っているか確認
