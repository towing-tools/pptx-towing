# pptx-towing

AIにパワーポイントを自動作成させるためのツール
A tool for automatic PowerPoint generation using AI agents

---

## これは何？ / What is this?

**Gemini CLI**、**Claude Code**、**Codex** などの「エージェント型AI」に、会社のテンプレートに沿ったパワーポイントを自動生成させるためのツールです。
This tool enables AI agents like **Gemini CLI**, **Claude Code**, and **Codex** to automatically generate PowerPoint presentations following your company template.

### エージェント型AIとは？ / What are AI Agents?

通常のブラウザ版AIとの違い:
Key differences from browser-based AI:

| | ブラウザ版AI / Browser AI | エージェント型AI / AI Agent |
|---|---|---|
| 動作場所 / Location | クラウド上 / Cloud | あなたのPC上 / Your PC |
| ファイル操作 / File Operations | できない / No | **できる / Yes** |
| プログラム実行 / Run Programs | できない / No | **できる / Yes** |
| フォルダ作成 / Create Folders | できない / No | **できる / Yes** |

エージェント型AIは、あなたのPCでプログラムを実行したり、ファイルを作成・編集したりできます。
AI agents can run programs and create/edit files on your PC.

このツールは、その能力を使ってパワーポイントを自動生成します。
This tool leverages those capabilities to automatically generate PowerPoint presentations.

### 動作イメージ / How It Works

```
1. Gemini CLI（またはClaude Code）を起動
   Launch Gemini CLI (or Claude Code)

2. AIに指示を出す:
   Give instructions to the AI:
   「MANDATORY_GUIDE_FOR_PPTX_GENERATION.md を読んで、
    2026年度事業計画のスライドを作成してください」
   "MANDATORY_GUIDE_FOR_PPTX_GENERATION.md and
    create slides for the 2026 business plan"

3. AIが自動で:
   AI automatically:
   - ガイドを読む / Reads the guide
   - content.md を作成 / Creates content.md
   - build.sh を実行 / Runs build.sh
   - final.pptx を生成 / Generates final.pptx

4. 完成したパワーポイントを開いて確認
   Open and review the completed PowerPoint
```

---

## 対応環境 / Supported Environments

| OS | 対応状況 / Status | 備考 / Notes |
|---|---|---|
| macOS | ✅ 動作確認済み / Tested | 推奨環境 / Recommended |
| Windows | ⚠️ 動作未確認 / Untested | 問題があれば報告してください / Report issues if any |
| Linux | ✅ 動作確認済み / Tested | |

---

## セットアップ手順 / Setup Instructions

### Step 1: エージェント型AIをインストール / Install an AI Agent

以下のいずれかをインストールしてください:
Install one of the following:

- **Gemini CLI** (推奨 / Recommended) - Googleのエージェント型AI（TOWINGメンバーなら無料で使えます）/ Google's AI agent (Free for TOWING members)
- **Claude Code** - Anthropicのエージェント型AI（有料だが最もおすすめ）/ Anthropic's AI agent (Paid, but highly recommended)
- **Codex** - OpenAIのエージェント型AI（有料）/ OpenAI's AI agent (Paid)

インストール方法は、各ツールの公式ドキュメントを参照するか、佐久間に聞いてください。
For installation instructions, refer to the official documentation or ask Sakuma.

> **ポイント / Key Point**: エージェント型AIは、ローカル環境でプログラムの実行権限、ファイル・フォルダの作成権限を持っています。これがブラウザ版との大きな違いです。
> AI agents have permission to run programs and create files/folders in your local environment. This is the major difference from browser-based AI.

### Step 2: このフォルダでAIを起動 / Launch AI in This Folder

1. このフォルダ全体をダウンロード / Download this entire folder
2. ターミナル（Windowsの場合はPowerShell）を開く / Open Terminal (PowerShell on Windows)
3. このフォルダに移動して、AIを起動 / Navigate to this folder and launch AI

```bash
# macOS / Linux
cd /path/to/pptx-towing-v3
gemini   # または claude / or claude

# Windows (PowerShell)
cd C:\path\to\pptx-towing-v3
gemini   # または claude / or claude
```

または、このフォルダが確認できる親フォルダでAIを起動してもOKです。
Alternatively, you can launch AI in a parent folder where this folder is visible.

### Step 3: テンプレートファイルを配置 / Place Template File

**重要 / Important**: テンプレートファイルは社外秘のためGitHubには含まれていません。
The template file is confidential and not included in GitHub.

Slackの社内共有メッセージに添付されている `2026TOWINGtemplate.pptx` をダウンロードし、`assets/` フォルダに配置してください:
Download `2026TOWINGtemplate.pptx` from the Slack internal message and place it in the `assets/` folder:

```
pptx-towing-v3/
  └── assets/
      └── 2026TOWINGtemplate.pptx  ← ここに配置 / Place here
```

### Step 4: 必要なツールの確認 / Check Required Tools

以下のツールがインストールされている必要があります:
The following tools must be installed:

```bash
# Node.js (v18以上 / v18 or higher)
node --version

# Python 3
python3 --version   # macOS/Linux
python --version    # Windows
```

インストールされていない場合:
If not installed:
- Node.js: https://nodejs.org/ からダウンロード / Download from https://nodejs.org/
- Python: https://www.python.org/ からダウンロード / Download from https://www.python.org/

### Step 5: 依存パッケージのインストール / Install Dependencies

```bash
cd pptx-towing-v3
npm install
```

---

## 使い方 / How to Use

### 1. AIにスライド作成を依頼 / Request Slide Creation from AI

Gemini CLI や Claude Code を起動した状態で、以下のように指示:
With Gemini CLI or Claude Code running, give instructions like:

```
pptx-towing-v3/MANDATORY_GUIDE_FOR_PPTX_GENERATION.md を読んで、
〇〇についてのスライドを作成してください

Read pptx-towing-v3/MANDATORY_GUIDE_FOR_PPTX_GENERATION.md and
create slides about 〇〇
```

例 / Examples:
- 「2026年度Q4の売上報告スライドを10枚で作成して」/ "Create 10 slides for Q4 2026 sales report"
- 「新製品の紹介プレゼンを作成して」/ "Create a presentation introducing the new product"
- 「投資家向けピッチ資料を作成して」/ "Create an investor pitch deck"

### 2. AIが自動で処理 / AI Processes Automatically

AIは以下を自動で行います:
AI will automatically:
1. ガイドを読んで作業手順を理解 / Read the guide and understand the workflow
2. `output/<プロジェクト名>/content.md` を作成 / Create `output/<project_name>/content.md`
3. `build.sh`（Windowsの場合は`build.ps1`）を実行 / Run `build.sh` (or `build.ps1` on Windows)
4. `output/<プロジェクト名>/final.pptx` を生成 / Generate `output/<project_name>/final.pptx`

### 3. 完成したファイルを確認 / Check the Completed File

```
output/<プロジェクト名>/final.pptx
```

を開いて内容を確認してください。
Open the file and review the content.

---

## 手動でビルドする場合 / Manual Build

AIを使わず、手動でビルドすることも可能です。
You can also build manually without using AI.

**macOS / Linux:**
```bash
./build.sh <プロジェクト名 / project_name>
```

**Windows (PowerShell):**
```powershell
.\build.ps1 <プロジェクト名 / project_name>
```

---

## プロジェクト構成 / Project Structure

```
pptx-towing-v3/
├── assets/                    # テンプレートファイル / Template files
│   └── 2026TOWINGtemplate.pptx
├── docs/                      # ドキュメント / Documentation
│   ├── procedure.md           # メイン手順書 / Main procedure
│   ├── templates.md           # スライドテンプレート集 / Slide templates
│   ├── content-guide.md       # コンテンツ設計ガイド / Content design guide
│   ├── config-reference.md    # 設定リファレンス / Config reference
│   └── troubleshooting.md     # トラブルシューティング / Troubleshooting
├── scripts/                   # スクリプト / Scripts
│   ├── generate_from_md.js    # Markdown → PPTX変換 / Markdown to PPTX conversion
│   └── inject_slides.py       # テンプレート注入 / Template injection
├── output/                    # 出力先（AIが自動作成）/ Output (auto-created by AI)
│   └── <project_name>/
│       ├── content.md         # コンテンツ定義 / Content definition
│       ├── intermediate.pptx  # 中間ファイル / Intermediate file
│       └── final.pptx         # 最終成果物 / Final output
├── build.sh                   # ビルドスクリプト (macOS/Linux) / Build script
├── build.ps1                  # ビルドスクリプト (Windows) / Build script
├── MANDATORY_GUIDE_FOR_PPTX_GENERATION.md  # AIに読ませるガイド / Guide for AI
└── README.md                  # このファイル / This file
```

---

## トラブルシューティング / Troubleshooting

### `npm install` でエラーが出る / Errors with `npm install`

Node.jsのバージョンが古い可能性があります。v18以上を推奨:
Your Node.js version might be outdated. v18 or higher is recommended:
```bash
node --version  # v18.x.x 以上か確認 / Check if v18.x.x or higher
```

### テンプレートが見つからないエラー / Template Not Found Error

`assets/2026TOWINGtemplate.pptx` が配置されているか確認してください。
Verify that `assets/2026TOWINGtemplate.pptx` is placed correctly.

### Windowsで `build.ps1` が実行できない / Cannot Execute `build.ps1` on Windows

PowerShellの実行ポリシーを変更する必要があるかもしれません:
You may need to change the PowerShell execution policy:
```powershell
Set-ExecutionPolicy -ExecutionPolicy RemoteSigned -Scope CurrentUser
```

### Playwrightのエラー / Playwright Errors

初回実行時にブラウザのインストールが必要な場合があります:
Browser installation may be required on first run:
```bash
npx playwright install chromium
```

---

## Windows環境について / About Windows Environment

**注意 / Note**: Windows環境での動作は十分にテストされていません。
Operation on Windows has not been thoroughly tested.

問題が発生した場合は、以下の情報と共に報告してください:
If you encounter issues, please report with the following information:
- エラーメッセージ / Error message
- 実行したコマンド / Command executed
- Node.js / Python のバージョン / Node.js / Python version

報告先: 佐久間まで
Report to: Sakuma

### WSLを使う場合（推奨）/ Using WSL (Recommended)

WSL（Windows Subsystem for Linux）を使うと、macOS/Linuxと同じ手順で使えます:
Using WSL (Windows Subsystem for Linux) allows you to follow the same steps as macOS/Linux:

```powershell
# WSLインストール（管理者権限のPowerShellで）
# WSL installation (in PowerShell with admin privileges)
wsl --install

# WSL内で操作 / Operations within WSL
wsl
cd /mnt/c/path/to/pptx-towing-v3
./build.sh <project_name>
```

---

## ライセンス / License

社内利用限定
Internal use only
