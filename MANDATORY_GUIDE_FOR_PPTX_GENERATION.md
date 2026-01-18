# MANDATORY GUIDE FOR PPTX GENERATION v2.0 (鉄の掟・改良版)

このプロジェクトでスライドを生成するすべてのAIおよび開発者は、以下の「デザインと内容の分離」原則に絶対に従わなければなりません。

---

## 1. 核心原則：デザインと内容の完全分離

デザイン（色、位置、フォント、余白）は**プログラムが支配**し、AIは**構造と内容のみ**を考えます。

*   **憲法 (Design)**: `assets/theme_config.json` が唯一の正解です。
*   **自由 (Content)**: Markdownファイル (`.md`) でスライドの構造を記述します。
*   **執行 (Pipeline)**: 指定されたスクリプトを使い、手作業でのXML/CSS編集は一切禁止します。

---

## 2. 憲法：`assets/theme_config.json`（改良版）

### 2.1 概要

このファイルはデザインの「絶対的なルール」です。

*   タイトルの座標、本文のフォントサイズ、利用可能な配色パレットが定義されています。
*   **デザインに不満がある場合は、Markdownをいじるのではなく、このJSONを修正して全体の整合性を保って下さい。**

### 2.2 拡張構成（v2.0 新規）

```json
{
  "version": "2.0",
  "metadata": {
    "project_name": "2026 TOWING",
    "color_theory": "60-30-10-rule",
    "last_updated": "2026-01-16",
    "maintainer": "Design System Team"
  },

  "color_palette": {
    "systematic": {
      "primary": {
        "color": "#0A4BA3",
        "role": "Base background, main text, trust element",
        "semantics": "Trust, corporate, stability",
        "contrast_ratio": "7.5:1",
        "wcag_level": "AAA"
      },
      "secondary": {
        "color": "#F7931E",
        "role": "Emphasis, call-to-action, highlight",
        "semantics": "Energy, growth, attention",
        "contrast_ratio": "6.8:1",
        "wcag_level": "AAA"
      },
      "success": {
        "color": "#22C55E",
        "role": "Positive messages, achievements, approvals",
        "semantics": "Success, approval, positive",
        "usage_context": "Growth metrics, completed tasks"
      },
      "warning": {
        "color": "#EAB308",
        "role": "Caution, important notes, pending items",
        "semantics": "Warning, attention needed",
        "usage_context": "Risks, challenges, to-do items"
      },
      "danger": {
        "color": "#EF4444",
        "role": "Critical issues, risks, errors",
        "semantics": "Error, negative, critical",
        "usage_context": "Errors, critical risks, blockers"
      },
      "neutral": {
        "color": "#64748B",
        "role": "Supporting text, borders, secondary information",
        "semantics": "Neutral, supporting",
        "usage_context": "Captions, metadata, borders"
      }
    },
    "usage_ratios": {
      "dominant": "60%",
      "secondary": "30%",
      "accent": "10%",
      "comment": "Primary color 60%, Secondary 30%, Accent colors 10%"
    }
  },

  "typography": {
    "hierarchy": {
      "h1_title_slide": {
        "font": "Arial, Helvetica, sans-serif",
        "size": 60,
        "weight": 700,
        "line_height": 1.1,
        "role": "Presentation main title",
        "color": "primary"
      },
      "h2_section_title": {
        "font": "Arial, Helvetica, sans-serif",
        "size": 44,
        "weight": 700,
        "line_height": 1.2,
        "role": "Main slide titles",
        "color": "primary"
      },
      "h3_subsection": {
        "font": "Arial, Helvetica, sans-serif",
        "size": 32,
        "weight": 600,
        "line_height": 1.3,
        "role": "Box titles, subsection headings",
        "color": "primary"
      },
      "body_text": {
        "font": "Arial, Helvetica, sans-serif",
        "size": 24,
        "weight": 400,
        "line_height": 1.4,
        "role": "Main content, bullet points",
        "color": "neutral"
      },
      "body_small": {
        "font": "Arial, Helvetica, sans-serif",
        "size": 18,
        "weight": 400,
        "line_height": 1.5,
        "role": "Captions, notes, metadata",
        "color": "neutral"
      }
    },
    "contrast_requirements": {
      "normal_text": "4.5:1",
      "large_text": "3:1",
      "wcag_standard": "WCAG 2.1 Level AA minimum",
      "testing_tool": "WebAIM Contrast Checker"
    }
  },

  "spacing_system": {
    "unit": "8px",
    "scale": {
      "xs": "4px",
      "sm": "8px",
      "md": "16px",
      "lg": "24px",
      "xl": "32px",
      "2xl": "48px",
      "3xl": "64px"
    }
  },

  "component_guidelines": {
    "boxes": {
      "default": {
        "background": "secondary",
        "text_color": "primary",
        "border_width": "0px",
        "padding": "lg",
        "border_radius": "0px",
        "use_case": "Standard information box"
      },
      "highlight": {
        "background": "#FFF8E7",
        "text_color": "primary",
        "border_left": "4px solid secondary",
        "padding": "lg",
        "border_radius": "0px",
        "use_case": "Important takeaway, emphasis"
      },
      "warning": {
        "background": "#FEF3C7",
        "text_color": "#78350F",
        "border_left": "4px solid warning",
        "padding": "lg",
        "border_radius": "0px",
        "use_case": "Risks, challenges, cautions"
      },
      "primary": {
        "background": "primary",
        "text_color": "#FFFFFF",
        "padding": "lg",
        "border_radius": "0px",
        "font_weight": 600,
        "use_case": "Critical messages, main CTA"
      }
    }
  }
}
```

### 2.3 Config使用のルール

- **色には常に「意味」を付与** → 赤は「リスク」、緑は「成功」という一貫したルール  
- **コントラスト比を明示** → WCAG AA (4.5:1) 以上を必須  
- **60-30-10ルール** → theme_config.jsonで比率を定義し、AIに比率遵守を指示  
- **コンポーネントの再利用性** → boxes, columns, cardsはすべてConfigで定義  

---

## 3. 入力：Markdown 記述ルール（改良版）

スライドの内容は以下の形式で記述します。

### 3.0 ファイルヘッダ（メタデータ）- 新規追加

すべてのMarkdownファイルの冒頭に以下を記述してください：

```markdown
***
title: "プレゼンテーション全体タイトル"
subtitle: "サブタイトル"
author: "著者名"
date: 2026-01-16
total_slides: 12
target_audience: "経営層 / 技術チーム / 営業部"
key_message: "3単語で伝えたい中心メッセージ"
tone: "corporate"
visual_strategy:
  primary_chart_type: "comparison"
  use_icons: false
  use_images: false
***
```

**効果**: AIが冒頭で制約条件を認識し、不適切な提案を避ける。

### 3.1 基本構造

```markdown
# スライドタイトル
ここにスライドの導入文を書きます。**強調**も使えます。

* 箇条書き1
* 箇条書き2

***
# 次のスライドタイトル
...
```

### 3.2 スライドメタデータ - 新規追加

各スライドの冒頭に以下を記述（省略可、推奨）：

```markdown
***
slide_number: 3
layout: "2-column-comparison"
key_message: "ビジネスインパクトを明示"
visual_element: "side-by-side table"
content_density: "medium"
color_scheme: "default"
***

# スライドタイトル
...
```

**メタデータフィールド説明**:
- `layout`: スライド構成方式（`single-column`, `2-column-comparison`, `header-main-cta` など）
- `key_message`: スライドの要点（1文）
- `visual_element`: ビジュアル要素の種類
- `content_density`: 情報量レベル（`low` / `medium` / `high`）
- `color_scheme`: ボックスの色スキーム（`default` / `highlight` / `warning` / `primary`）

**効果**: AIが各スライドの役割を理解し、予期しないレイアウト変更を防止。

### 3.3 拡張コンポーネント (::: box)

デザイン性を高めるために「ボックス」を使用できます。スタイル名はConfigのパレットに基づきます。

*   `::: box`: 標準（セカンダリアクセント）
*   `::: box highlight`: 強調（ゴールド背景、左ボーダー）
*   `::: box warning`: 警告・課題（黄色背景、警告色ボーダー）
*   `::: box primary`: 反転（プライマリ背景・白抜き文字）

### 3.4 視覚化要素マークアップ - 新規追加

ビジュアル化が必要な場合、以下のマークアップを使用してください：

```markdown
# 売上推移と市場シェア

`[VIZ: line_chart]`
- X軸: 年度（2020-2026）
- Y軸: 売上（百万円）
- 系列1: 当社売上（青）
- 系列2: 市場平均（グレー）
- 注釈: 2024年から成長加速
`[/VIZ]`

結論: 当社のシェアは年間15%で拡大中
```

**利用可能なマークアップ**:
```
[VIZ: line_chart]        → 折れ線グラフ
[VIZ: bar_chart]         → 棒グラフ
[VIZ: pie_chart]         → 円グラフ
[VIZ: comparison_table]  → 比較表
[VIZ: 2x2_matrix]        → 2×2マトリクス
[VIZ: flow_diagram]      → フロー図
[VIZ: timeline]          → タイムライン
[VIZ: icon_row]          → アイコン列
```

### 3.5 カラムレイアウト (::: columns)

**2カラムレイアウトを使用する際は、以下の構文を厳守してください。**

#### 基本構文（必須パターン）
```markdown
::: columns 50/50

::: column
（左カラムの内容）
:::

::: column
（右カラムの内容）
:::

:::
```

#### 重要ルール
1. **`::: columns XX/YY`** で開始（XX:YY は左右の比率、例: `50/50`, `60/40`, `40/60`）
2. **`::: column`** で各カラムを開始（必ず2つ必要）
3. **`:::`** で各要素を閉じる
4. **閉じタグの順序**: column → column → columns の順で閉じる
5. **空行**: `::: columns` の直後と `:::` の直前に空行を入れると安定

#### カラム内にボックスを入れる場合
```markdown
::: columns 50/50

::: column
::: box
**経済的要因**
* 子育て費用の増大
* 非正規雇用の増加
:::
:::

::: column
::: box warning
**社会的要因**
* 晩婚化・未婚化の進行
* 仕事と育児の両立困難
:::
:::

:::
```

---

## 4. スライド構成テンプレート集（改良版）

### 4.1 スライド別設計ガイド

#### Slide 1: Title Slide
**情報量**: 最小（タイトル + 副題のみ）  
**ビジュアル**: 企業ロゴ、背景画像  
**禁止**: 本文、箇条書き

```markdown
***
slide_number: 1
layout: "title-slide"
content_density: "low"
***

# スライドのメイン題目
サブタイトルまたは日付
```

---

#### Slide 2-3: 背景・課題提示
**情報量**: 中  
**必須ビジュアル**: グラフ or 統計数値  
**構成パターン**: 導入文 + ボックス + 3-5箇条書き

```markdown
***
slide_number: 2
layout: "single-column"
key_message: "なぜこの課題が重要なのか"
visual_element: "statistic_chart"
content_density: "medium"
***

# なぜこの課題が重要なのか

背景文（1-2行）: 日本の少子化は〜

::: box primary
**重要な数値**: 2025年の出生数は68万人（前年比3.2%減）
:::

**主な要因**:
* 経済的困窮層の増加
* 晩婚化の進行（男性平均初婚年齢: 31.7歳）
* 仕事と育児の両立困難
```

---

#### Slide 4-5: 解決策・提案
**情報量**: 中-高  
**必須ビジュアル**: 2-column comparison or flow diagram  
**構成パターン**: 複数の提案を並列提示

```markdown
***
slide_number: 4
layout: "2-column-comparison"
key_message: "3つの提案と最適解"
visual_element: "comparison_table"
content_density: "high"
color_scheme: "highlight"
***

# 提案内容：3つのアプローチ

`[VIZ: comparison_table]`

::: columns 50/50

::: column
::: box
**提案A: 支援策の充実**
* 児童手当の拡充
* 教育費の無償化
* 保育施設の整備
* 期待効果: 出生率1.5%改善
:::
:::

::: column
::: box highlight
**提案B: 働き方改革（推奨）**
* 時短勤務制度の拡充
* リモートワーク推進
* 男性育休の取得支援
* 期待効果: 出生率2.8%改善
:::
:::

:::

`[/VIZ]`
```

---

#### Slide 9-10: 実装計画
**情報量**: 高  
**必須ビジュアル**: Timeline or Gantt chart  
**構成パターン**: 時系列 + 責任者 + 予算

```markdown
***
slide_number: 9
layout: "header-main"
key_message: "2026年の実装ロードマップ"
visual_element: "timeline"
content_density: "high"
***

# 実装ロードマップ：Q1-Q4 2026

`[VIZ: timeline]`
* Q1 2026: 基礎調査・ステークホルダー協議
* Q2 2026: 試験運用・パイロット実施
* Q3 2026: フィードバック収集・改善
* Q4 2026: 本格実施・システム統合
`[/VIZ]`

| フェーズ | 期間 | 予算 | 責任部門 |
|---------|------|------|---------|
| 企画 | Q1 | 50万 | 営業 |
| 実行 | Q2-Q3 | 500万 | 技術 |
| 最適化 | Q4 | 100万 | 営業 |
```

---

#### Slide 12: まとめ・CTA
**情報量**: 最小  
**ビジュアル**: 最小限  
**必須**: Call-to-Action（次のステップ）

```markdown
***
slide_number: 12
layout: "title-cta"
key_message: "次のステップの明示"
content_density: "low"
color_scheme: "primary"
***

# 次のステップ

::: box primary
**アクション**
1. 提案を経営会議で承認（1月31日）
2. 実装チーム編成（2月15日）
3. 詳細計画書作成（3月31日）
:::

**ご質問・ご意見をお聞きします**
```

---

### 4.2 よく使うレイアウトパターン

#### パターン1: シンプル2カラム比較
```markdown
***
slide_number: X
layout: "2-column-comparison"
content_density: "medium"
***

# スライドタイトル

::: columns 50/50

::: column
::: box
**項目A**
* ポイント1
* ポイント2
* ポイント3
:::
:::

::: column
::: box highlight
**項目B**
* ポイント1
* ポイント2
* ポイント3
:::
:::

:::
```

---

#### パターン2: ヘッダー + 2カラム
```markdown
***
slide_number: X
layout: "header-main-cta"
content_density: "medium"
***

# スライドタイトル

::: box primary
**重要なメッセージ**: ここに要約を書く
:::

::: columns 60/40

::: column
**詳細説明**
* 箇条書き1
* 箇条書き2
* 箇条書き3
* 箇条書き4
* 箇条書き5
:::

::: column
::: box highlight
**ポイント**
結論や目標を記載
:::
:::

:::
```

---

#### パターン3: 導入文 + ハイライトボックス + 箇条書き
```markdown
***
slide_number: X
layout: "single-column"
content_density: "medium"
***

# スライドタイトル

導入文をここに書きます。

::: box highlight
**強調したい内容**をここに書きます。
:::

* 箇条書き1
* 箇条書き2
* 箇条書き3
* 箇条書き4
* 箇条書き5
```

---

## 5. スライドコンテンツ設計（充実したスライドを作るために）

**1スライド1メッセージ**の原則を守りつつ、**スライド全体を使い切る**ことが重要です。  
下部が空いたスライドは見栄えが悪く、情報が薄い印象を与えます。

### 5.1 レイアウト選択の原則

| コンテンツ量 | 推奨レイアウト | 最小コンテンツ | 備考 |
|-------------|---------------|---------------|------|
| 比較・対比が必要な2項目 | 2カラム（50/50） | 各カラムに4-5項目 | 高度な比較に使用 |
| メイン + 補足情報 | 2カラム（60/40 or 70/30） | 左5項目+右3項目 | 左にリスト、右にまとめ |
| 単一トピックの深掘り | シングルカラム | 5-7項目+ボックス | 導入文 + ボックス + 箇条書き |
| 重要メッセージ + 詳細 | ヘッダーボックス + 2カラム | primary box + 各カラム3-4項目 | primary box を上部に |

### 5.2 コンテンツ密度レベルの定義

- **LOW** (タイトルスライド、転換スライド)  
  タイトル + 1-2個のボックス のみ

- **MEDIUM** (標準スライド)  
  タイトル + 導入文 + ボックス + 4-6項目

- **HIGH** (データ提示スライド)  
  タイトル + グラフ + テーブル + 6-8項目

### 5.3 最小コンテンツ量の基準

**❌ スカスカになるパターン（避けるべき）**:
- 箇条書き3項目のみ
- ボックス1つと箇条書き2項目
- 2カラムで各1-2項目

**✅ 充実したスライドの基準**:

#### シングルカラムの場合
```markdown
# タイトル
導入文（1-2文、背景や文脈を説明）

::: box highlight
**キーメッセージ**（1-2文の要約または数値データ）
:::

* 箇条書き1（具体的な説明や数値を含む）
* 箇条書き2
* 箇条書き3
* 箇条書き4
* 箇条書き5（最低5項目、できれば6-7項目）
```

#### 2カラムの場合
```markdown
# タイトル

::: columns 50/50

::: column
::: box
**カテゴリA**
* 項目1（具体的に）
* 項目2
* 項目3
* 項目4（各カラム最低4項目）
:::
:::

::: column
::: box highlight
**カテゴリB**
* 項目1
* 項目2
* 項目3
* 項目4
:::
:::

:::
```

#### ヘッダー + 2カラムの場合
```markdown
# タイトル

::: box primary
**重要メッセージ**: 2-3文で要約。具体的な数値や目標を含める。
:::

::: columns 60/40

::: column
**詳細説明**
* 項目1（具体的な施策や内容）
* 項目2
* 項目3
* 項目4
* 項目5（左カラムは5項目以上）
:::

::: column
::: box highlight
**ポイント/目標**
結論や目標を2-3文で記載。
数値目標があれば明記。
:::
:::

:::
```

### 5.4 コンテンツが足りない場合の対処法

1. **より具体的な情報を追加**
   - 数値データ（〇〇%、〇〇万人など）
   - 年次・期限（2023年〜、2030年目標など）
   - 具体例や事例

2. **背景・文脈を追加**
   - なぜこのトピックが重要なのか
   - 歴史的経緯
   - 他との比較

3. **補足情報を追加**
   - 課題・リスク
   - 今後の展望
   - 関連する取り組み

4. **構造を見直す**
   - 1スライドの内容が薄いなら、前後のスライドと統合を検討
   - 逆に情報過多なら、スライドを分割

### 5.5 NG例とOK例

**❌ NG: スカスカスライド**
```markdown
# 少子化の原因

* 経済的要因
* 社会的要因
* 価値観の変化
```

**✅ OK: 充実したスライド**
```markdown
***
slide_number: 5
layout: "2-column-comparison"
content_density: "high"
***

# 少子化の主な原因

::: columns 50/50

::: column
::: box
**経済的要因**
* 子育て費用の増大（大学卒業まで約3,000万円）
* 非正規雇用の増加（若年層の約4割）
* 住宅費・教育費の高騰
* 実質賃金の伸び悩み
* 将来への経済的不安
:::
:::

::: column
::: box warning
**社会的要因**
* 晩婚化の進行（平均初婚年齢：男性31歳、女性30歳）
* 未婚率の上昇（50歳時点：男性28%、女性18%）
* 仕事と育児の両立困難
* 長時間労働文化
* 価値観の多様化
:::
:::

:::
```

---

## 6. 実行パイプライン（この順序を厳守）

### Phase 1: Outline生成（新規）

AIに概要構成を作成させます。

```bash
# AI に指定：
「このプレゼン資料から、12スライドの Outline を作ってください。
以下のテンプレートを参考にしてください。
[Slide 1]: Title Slide
[Slide 2-3]: Background
[Slide 4-5]: Solutions
...
各スライドの information density と layout を明示してください。」

出力: outline.md
```

### Phase 2: Content定義（新規）

Outline をもとに各スライドの詳細テキストを作成させます。

```bash
# AI に指定：
「outline.md に基づいて、各スライドの詳細コンテンツを作成してください。
content_density レベルに合わせて、箇条書きは最大6項目、
ボックス内は100字以下にしてください。」

出力: content.md
```

### Phase 3: Markdownから中間PPTXを生成

AIが作成したMarkdownを、Configに基づいてスライド化します。

```bash
node scripts/generate_from_md.js <入力.md> <一時出力.pptx>
```

*   **バリデーション**: ここでエラー（文字溢れ等）が出た場合、AIはMarkdownの文字数を減らすか、Configのフォントサイズを調整して再実行してください。

### Phase 4: 公式テンプレートへの注入

中間PPTXを正本テンプレート（`assets/2026TOWINGtemplate.pptx`）に合成します。

```bash
python3 scripts/inject_slides.py assets/2026TOWINGtemplate.pptx <一時出力.pptx> <最終成果物.pptx>
```

---

## 7. 品質チェックリスト

### 7.1 生成前チェック（Outline段階）

- [ ] 全体のスライド数は5-15枚か？（推奨: 10-12枚）
- [ ] 各スライドの「情報量レベル」が明示されているか？
- [ ] スライド間の論理的な流れは構成されているか？
- [ ] ビジュアル化が必要なスライドには `visual_element` が付いているか？
- [ ] CTAまたは結論スライドが明示されているか？

### 7.2 生成時チェック（Markdown生成後）

- [ ] 箇条書きは最大6項目を超えていないか？
- [ ] ボックス内のテキストは100字を超えていないか？
- [ ] 色の使い分けが一貫しているか（警告=赤、成功=緑）？
- [ ] フォントサイズの階層が正しいか？
- [ ] 余白が十分確保されているか？
- [ ] すべてのメタデータが正しく記述されているか？

### 7.3 生成後チェック（PPTX化前）

- [ ] テキストが「Text box ends too close to bottom edge」エラーを起こしていないか？
- [ ] すべてのテキストが正しいタグに包まれているか？
- [ ] 色のコントラスト比は WCAG AA 以上か？
- [ ] デザイン統一性に逸脱がないか？
- [ ] すべての `[VIZ: ...]` マークアップが正しくクローズされているか？

---

## 8. 禁止事項 (PROHIBITIONS)

*   ❌ **CSS/HTMLの直接記述**: `html2pptx` は `<style>` タグやインラインCSSを拒否します。
*   ❌ **座標のハードコード**: スライド上の位置計算をAIが行ってはいけません。
*   ❌ **非標準タグの使用**: `<p>`, `<h1>-<h6>`, `<ul>`, `<li>` 以外のテキストタグは使用しないでください。
*   ❌ **Configの改ざん**: 生成時に勝手にカラーパレットやフォントサイズを変更しないでください。

---

## 9. トラブルシューティング

### 9.1 一般的なエラー

*   **「Text box ends too close to bottom edge」エラー**:
    *   内容が多すぎます。Markdownの文章を削るか、`theme_config.json` の `body` フォントサイズを小さくしてください。
    *   あるいは、スライドを分割して情報量を減らしてください。

*   **「Unwrapped text」エラー**:
    *   すべてのテキストはタグ（`<p>`など）の中にある必要があります。Markdownエンジンの出力を確認してください。

*   **色がtheme_config.jsonと違う**:
    *   インラインカラースタイルが指定されていないか確認してください。
    *   すべての色指定は `theme_config.json` から参照してください。

### 9.2 デザイン品質関連

*   **コントラスト比が低い**:
    *   WebAIM Contrast Checker で確認し、WCAG AA (4.5:1) 以上を確保してください。
    *   `theme_config.json` の `contrast_ratio` を参照してください。

*   **スライドの上下が不揃い**:
    *   各スライドの `content_density` を確認し、情報量を調整してください。
    *   ボックスやカラムのパディング設定を確認してください。