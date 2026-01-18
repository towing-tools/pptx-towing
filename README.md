# pptx-towing

TOWING社内向けPowerPointスライド自動生成ツール

## セットアップ手順（初心者向け）

### 前提条件

- **macOS** を使用していること
- **Node.js** がインストールされていること

Node.jsがインストールされているか確認：
```bash
node --version
```

インストールされていない場合は、[Node.js公式サイト](https://nodejs.org/)からダウンロードしてください。

### 1. リポジトリをクローン

```bash
git clone https://github.com/YOUR_USERNAME/pptx-towing.git
cd pptx-towing
```

### 2. 依存パッケージをインストール

```bash
npm install
```

このコマンドで、`package.json`に記載された必要なパッケージ（pptxgenjs, playwright, sharp）が自動的にインストールされます。

### 3. テンプレートファイルをダウンロード（必須）

**重要**: テンプレートファイルは機密情報のためGitHubには含まれていません。

社内共有時のSlackメッセージに添付されている `2026TOWINGtemplate.pptx` をダウンロードしてください。

ダウンロード後、`assets/` フォルダに配置：
```
pptx-towing/
  └── assets/
      └── 2026TOWINGtemplate.pptx  ← ここに配置
```

### 4. システムツール（オプション）

一部の機能には追加ツールが必要です：

```bash
# Homebrewがない場合は先にインストール
/bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"

# PDF関連機能を使う場合
brew install poppler

# LibreOffice連携を使う場合
brew install --cask libreoffice
```

## 使い方

### HTMLからPowerPoint生成

```bash
node scripts/html2pptx.js [入力ファイル]
```

### Markdownから生成

```bash
node scripts/generate_from_md.js [入力ファイル]
```

## プロジェクト構成

```
pptx-towing/
├── assets/              # テンプレートファイル（.pptxは非公開）
├── scripts/             # メインスクリプト
│   ├── html2pptx.js     # HTML→PPTX変換
│   ├── generate_from_md.js  # Markdown→PPTX変換
│   ├── inject_slides.py # スライド挿入（Python）
│   └── inventory.py     # インベントリ管理（Python）
├── package.json         # Node.js依存関係
├── requirements.txt     # Python依存関係（参考）
└── README.md            # このファイル
```

## トラブルシューティング

### `npm install` でエラーが出る

Node.jsのバージョンが古い可能性があります。v18以上を推奨：
```bash
node --version  # v18.x.x 以上か確認
```

### テンプレートが見つからないエラー

`assets/2026TOWINGtemplate.pptx` が配置されているか確認してください。

## ライセンス

社内利用限定
