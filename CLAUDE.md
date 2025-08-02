# Excel Checker Project - Claude Instructions

## 🐍 Python仮想環境

### 仮想環境のアクティベーション
プロジェクト作業前に必ず以下の手順で仮想環境をアクティベートしてください：

```bash
cd C:\develop\claude-code\excel-checker\venv
Scripts\activate
```

### 仮想環境の確認
アクティベーション後、プロンプトに `(venv)` が表示されることを確認してください。

### 仮想環境の非アクティベーション
作業終了時は以下のコマンドで非アクティベートしてください：

```bash
deactivate
```

## 📁 プロジェクト構造

```
excel-checker/
├── venv/                    # Python仮想環境
├── sample_data/            # テスト用Excelファイル
│   ├── 顧客管理システム設計書.xlsx
│   ├── 在庫管理API詳細設計書.xlsx
│   └── テストケース仕様書.xlsx
├── src/                    # ソースコード
├── tests/                  # テストコード
└── CLAUDE.md              # このファイル
```

## 🔧 開発ルール

### 1. 作業開始時
1. 仮想環境をアクティベート
2. 必要な依存関係をインストール
3. テスト用データの確認

### 2. コード作成時
- Python 3.x を使用
- 型ヒント（Type Hints）を適用
- docstringを記述
- テストコードも併記

### 3. Excel処理
- openpyxl または pandas を使用
- セル単位での処理に対応
- シート間・ファイル間の比較機能

