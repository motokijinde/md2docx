# Markdown to Docx / PDF Converter

MarkdownファイルをMicrosoft Word文書 (.docx) または PDFファイル (.pdf) に変換するPythonスクリプトです。
外部設定ファイル (`config.json`) を使用して、フォントの種類やサイズ、色を柔軟にカスタマイズできます。

## 特徴
- **Markdown解析**: 見出し、リスト、太字、コードブロック、引用、画像、表に対応。
- **Mermaid対応**: コードブロック内のMermaid記法を画像として自動変換・埋め込み（Kroki.io APIを使用）。
- **スタイリング**: 日本語フォント（Meiryoなど）や等幅フォントを適切に設定可能。
- **PDF出力**: ReportLabライブラリを使用してPDFへの直接出力に対応。

## 動作環境
- Python 3.x
- Ubuntu 24.04 LTS (動作確認済み環境)

### Python3のインストール
Python3がインストールされていない場合は、以下のコマンドを実行してください。

```bash
sudo apt update
sudo apt install python3 python3-pip python3-venv
```

### 必須ライブラリ
以下のコマンドで必要なライブラリをインストールしてください。

```bash
# Word出力用 & PDF出力用
pip install python-docx reportlab --break-system-packages
```

※ Ubuntu環境でシステムパッケージとしてインストールする場合:
```bash
sudo apt update
sudo apt install python3-docx python3-reportlab
```

## 使い方

基本的な使い方は以下の通りです。

```bash
python3 md2docx.py <入力ファイル> [出力ファイル] [オプション]
```

### 例

**MarkdownをWord (.docx) に変換する**
```bash
python3 md2docx.py input.md
# -> input.docx が生成されます
```

**MarkdownをPDF (.pdf) に変換する**
```bash
python3 md2docx.py input.md --pdf
# -> input.pdf が生成されます
```

**出力ファイル名を指定する**
```bash
python3 md2docx.py input.md output_file.docx
```

### ヘルプの表示
```bash
python3 md2docx.py --help
```

### バージョン確認
```bash
python3 md2docx.py --version
```

## 設定 (config.json)

`md2docx.py` と同じディレクトリにある `config.json` を編集することで、フォントや色をカスタマイズできます。

### 設定ファイルの例

```json
{
    "fonts": {
        "normal": {
            "name": "Segoe UI",
            "eastAsia": "Meiryo",
            "size": 10.5,
            "pdf_name": "JapaneseFont"
        },
        "heading": {
            "name": "Segoe UI",
            "eastAsia": "Meiryo",
            "colors": [0, 0, 128],
            "sizes": [24, 18, 14, 12]
        },
        "code": {
            "docx_name": "Consolas",
            "pdf_name": "Courier",
            "size": 9,
            "colors": [0, 0, 128]
        },
        "bold": {
            "colors": [165, 42, 42]
        }
    },
    "pdf_font_paths": [
        "/usr/share/fonts/truetype/vlgothic/VL-Gothic-Regular.ttf",
        "/usr/share/fonts/opentype/noto/NotoSansCJK-Regular.ttc",
        "/usr/share/fonts/opentype/ipafont-gothic/ipag.ttf",
        "/usr/share/fonts/truetype/takao-gothic/TakaoGothic.ttf",
        "/usr/share/fonts/truetype/fonts-japanese-gothic.ttf"
    ]
}
```

### 設定項目一覧

| セクション | キー | 説明 | 備考 |
| :--- | :--- | :--- | :--- |
| `fonts.normal` | `name` | 通常テキストの欧文フォント名 | Docx用 |
| `fonts.normal` | `eastAsia` | 通常テキストの日本語フォント名 | Docx用 |
| `fonts.normal` | `size` | 通常テキストのサイズ (pt) | |
| `fonts.normal` | `pdf_name` | PDF生成時のフォント参照名 | 通常は "JapaneseFont" |
| `fonts.heading` | `name` | 見出しの欧文フォント名 | |
| `fonts.heading` | `eastAsia` | 見出しの日本語フォント名 | |
| `fonts.heading` | `colors` | 見出しの文字色 `[R, G, B]` | 例: `[0, 0, 128]` (Navy) |
| `fonts.heading` | `sizes` | 見出しレベル(H1-H4)ごとのサイズ配列 | 例: `[24, 18, 14, 12]` |
| `fonts.heading` | `page_break_level` | 改ページを行う見出しレベルの閾値 | デフォルト: 2 (H1, H2で改ページ) |
| `fonts.code` | `docx_name` | コードブロックのフォント名 (Docx) | 例: "Consolas" |
| `fonts.code` | `pdf_name` | コードブロックのフォント名 (PDF) | 例: "Courier" |
| `fonts.code` | `size` | コードブロックの文字サイズ (pt) | |
| `fonts.code` | `colors` | コードブロックの文字色 `[R, G, B]` | |
| `fonts.bold` | `colors` | 太字箇所の強調色 `[R, G, B]` | 例: `[165, 42, 42]` (Brown) |
| `(root)` | `pdf_font_paths` | PDF生成用のフォント検索パスリスト | Linux環境のパスを指定 |
