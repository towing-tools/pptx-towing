# theme_config.json リファレンス (Config Reference)

**このファイルはデザイン調整が必要な場合に参照してください。**

---

## 基本原則

`assets/theme_config.json` はデザインの「絶対的なルール」です。

- デザインに不満がある場合は、**Markdownではなくこのファイル**を修正
- 修正時は全体の整合性を保つこと
- 変更後は必ずテスト生成で確認

---

## カラーパレット

### システマティックカラー

| 名前 | カラーコード | 用途 | セマンティクス |
|------|-------------|------|---------------|
| primary | `#0A4BA3` | 基調色、メインテキスト、信頼要素 | 信頼、企業、安定 |
| secondary | `#F7931E` | 強調、CTA、ハイライト | エネルギー、成長、注目 |
| success | `#22C55E` | ポジティブメッセージ、達成 | 成功、承認 |
| warning | `#EAB308` | 注意、重要な注記 | 警告、注意が必要 |
| danger | `#EF4444` | 重大な問題、リスク、エラー | エラー、重大 |
| neutral | `#64748B` | 補足テキスト、ボーダー | 中立、補助 |

### 60-30-10ルール

- **60%**: primary（基調色）
- **30%**: secondary（補助色）
- **10%**: accent（アクセント色：success, warning, danger）

---

## タイポグラフィ

### フォント階層

```json
"typography": {
  "hierarchy": {
    "h1_title_slide": {
      "size": 60,
      "weight": 700,
      "role": "プレゼンテーション全体タイトル"
    },
    "h2_section_title": {
      "size": 44,
      "weight": 700,
      "role": "スライドタイトル"
    },
    "h3_subsection": {
      "size": 32,
      "weight": 600,
      "role": "ボックスタイトル、サブセクション"
    },
    "body_text": {
      "size": 24,
      "weight": 400,
      "role": "本文、箇条書き"
    },
    "body_small": {
      "size": 18,
      "weight": 400,
      "role": "キャプション、注釈"
    }
  }
}
```

### コントラスト要件

- 通常テキスト: 4.5:1以上（WCAG AA）
- 大きいテキスト: 3:1以上
- 確認ツール: WebAIM Contrast Checker

---

## スペーシングシステム

### 基本単位: 8px

| 名前 | サイズ | 用途 |
|------|--------|------|
| xs | 4px | 最小マージン |
| sm | 8px | 要素間の狭い余白 |
| md | 16px | 標準余白 |
| lg | 24px | セクション間 |
| xl | 32px | 大きなセクション間 |
| 2xl | 48px | 主要セクション間 |
| 3xl | 64px | ページ全体の余白 |

---

## コンポーネントガイドライン

### ボックススタイル

```json
"boxes": {
  "default": {
    "background": "secondary",
    "text_color": "primary",
    "border_width": "0px",
    "padding": "lg",
    "use_case": "標準的な情報ボックス"
  },
  "highlight": {
    "background": "#FFF8E7",
    "text_color": "primary",
    "border_left": "4px solid secondary",
    "use_case": "重要なポイント、強調"
  },
  "warning": {
    "background": "#FEF3C7",
    "text_color": "#78350F",
    "border_left": "4px solid warning",
    "use_case": "リスク、課題、注意"
  },
  "primary": {
    "background": "primary",
    "text_color": "#FFFFFF",
    "font_weight": 600,
    "use_case": "最重要メッセージ、CTA"
  }
}
```

---

## レイアウト設定

### スライドサイズ

- **幅**: 960pt（16:9標準）
- **高さ**: 540pt

### ヘッダー領域

```json
"header": {
  "title": {
    "x_pt": 40,
    "y_pt": 30
  }
}
```

### コンテンツ領域

```json
"content": {
  "x_pt": 40,
  "y_pt": 90,
  "w_pt": 880,
  "h_pt": 420
}
```

---

## 設定変更の例

### フォントサイズを大きくする

```json
// 変更前
"body_text": { "size": 24 }

// 変更後
"body_text": { "size": 28 }
```

**注意**: サイズを大きくするとテキストがはみ出す可能性あり

---

### ボックスの色を変更する

```json
// highlightの背景色を変更
"highlight": {
  "background": "#E8F4FD",  // 青系に変更
  "border_left": "4px solid #0A4BA3"
}
```

---

### 余白を調整する

```json
// コンテンツ領域の上部余白を増やす
"content": {
  "y_pt": 100,  // 90 → 100
  "h_pt": 410   // 高さを調整
}
```

---

## 変更時の注意事項

### 必ず確認すること

1. **コントラスト比**: 変更後もWCAG AA (4.5:1) を満たすか
2. **整合性**: 他のコンポーネントとの色バランス
3. **テスト生成**: 実際にスライドを生成して確認

### 変更禁止事項

- スライドサイズの変更（16:9を維持）
- フォントファミリーの無断変更
- ボックスの基本構造の変更

---

## config.jsonの完全な構造

```json
{
  "version": "2.0",
  "metadata": {
    "project_name": "2026 TOWING",
    "color_theory": "60-30-10-rule",
    "last_updated": "2026-01-18"
  },
  "color_palette": {
    "systematic": { ... },
    "usage_ratios": { ... }
  },
  "typography": {
    "hierarchy": { ... },
    "contrast_requirements": { ... }
  },
  "spacing_system": {
    "unit": "8px",
    "scale": { ... }
  },
  "component_guidelines": {
    "boxes": { ... }
  }
}
```

詳細は `assets/theme_config.json` を直接参照してください。
