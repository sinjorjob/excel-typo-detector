"""
Excel誤字脱字検出ツール

主要モジュール:
- extract: Excelファイルからセルデータを抽出
- normalize: テキストの正規化処理
- detect: 表記ゆれ・誤字の検出
- llm: LLM連携による判定
- report: 結果出力・レポート生成
- dictionary_io: 辞書ファイルの読み込み
"""

__version__ = "1.0.0"
__author__ = "Excel Checker Team"