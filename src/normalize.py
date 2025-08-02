"""
テキストの正規化処理モジュール
全角/半角、長音、ハイフンなどの統一を行う
"""

import re
import unicodedata
from typing import List, Optional

import neologdn
from loguru import logger

from .extract import CellData


class TextNormalizer:
    """テキスト正規化クラス"""
    
    def __init__(self):
        """初期化"""
        # 正規化パターンの定義
        self.patterns = {
            # ハイフン類の統一（‐–—-）
            'hyphen': re.compile(r'[‐–—−]'),
            # 長音記号の統一
            'long_vowel': re.compile(r'[ー−]'),
            # 波ダッシュの統一
            'wave_dash': re.compile(r'[〜～]'),
            # パーセント記号の統一
            'percent': re.compile(r'[％%]'),
            # 連続空白の統一
            'spaces': re.compile(r'\s+'),
            # 全角括弧の統一
            'brackets': re.compile(r'[（）]'),
            # 中黒・ドットの統一
            'dots': re.compile(r'[・･]'),
        }
        
        # 置換文字の定義
        self.replacements = {
            'hyphen': '-',
            'long_vowel': 'ー',
            'wave_dash': '～',
            'percent': '%',
            'spaces': ' ',
            'brackets_open': '(',
            'brackets_close': ')',
            'dots': '・',
        }
    
    def normalize_text(self, text: str) -> str:
        """
        テキストの正規化処理
        
        Args:
            text: 正規化対象のテキスト
            
        Returns:
            正規化されたテキスト
        """
        if not text or text.strip() == "":
            return text
            
        try:
            # 1. Unicode正規化（NFC）
            normalized = unicodedata.normalize('NFC', text)
            
            # 2. neologdnによる日本語テキスト正規化
            # - 全角/半角の統一
            # - 不要な文字の除去
            # - 文字幅の統一
            normalized = neologdn.normalize(normalized)
            
            # 3. 個別パターンの正規化
            normalized = self._apply_custom_patterns(normalized)
            
            # 4. 前後の空白をトリム
            normalized = normalized.strip()
            
            return normalized
            
        except Exception as e:
            logger.warning(f"Text normalization failed: {text[:50]}... - {e}")
            return text
    
    def _apply_custom_patterns(self, text: str) -> str:
        """
        カスタム正規化パターンを適用
        
        Args:
            text: 対象テキスト
            
        Returns:
            パターン適用後のテキスト
        """
        result = text
        
        # ハイフン類の統一
        result = self.patterns['hyphen'].sub(self.replacements['hyphen'], result)
        
        # 長音記号の統一
        result = self.patterns['long_vowel'].sub(self.replacements['long_vowel'], result)
        
        # 波ダッシュの統一
        result = self.patterns['wave_dash'].sub(self.replacements['wave_dash'], result)
        
        # パーセント記号の統一
        result = self.patterns['percent'].sub(self.replacements['percent'], result)
        
        # 中黒・ドットの統一
        result = self.patterns['dots'].sub(self.replacements['dots'], result)
        
        # 括弧の統一（全角→半角）
        result = result.replace('（', '(').replace('）', ')')
        
        # 連続空白の統一
        result = self.patterns['spaces'].sub(self.replacements['spaces'], result)
        
        return result
    
    def normalize_cells(self, cells: List[CellData]) -> List[CellData]:
        """
        セルデータリストの正規化
        
        Args:
            cells: セルデータのリスト
            
        Returns:
            正規化されたセルデータのリスト
        """
        logger.info(f"Normalizing {len(cells)} cells")
        
        for cell in cells:
            try:
                cell.text_norm = self.normalize_text(cell.text)
            except Exception as e:
                logger.warning(f"Failed to normalize cell {cell.file_name}:{cell.sheet_name}:{cell.cell_address} - {e}")
                cell.text_norm = cell.text
        
        logger.info("Text normalization completed")
        return cells
    
    def compare_normalization(self, original: str, normalized: str) -> bool:
        """
        正規化前後のテキストを比較
        
        Args:
            original: 元のテキスト
            normalized: 正規化後のテキスト
            
        Returns:
            変更があった場合True
        """
        return original != normalized
    
    def get_normalization_stats(self, cells: List[CellData]) -> dict:
        """
        正規化統計情報を取得
        
        Args:
            cells: セルデータのリスト
            
        Returns:
            統計情報の辞書
        """
        stats = {
            'total_cells': len(cells),
            'normalized_cells': 0,
            'unchanged_cells': 0,
            'normalization_types': {
                'hyphen': 0,
                'long_vowel': 0,
                'wave_dash': 0,
                'percent': 0,
                'brackets': 0,
                'dots': 0,
                'spaces': 0,
            }
        }
        
        for cell in cells:
            if self.compare_normalization(cell.text, cell.text_norm):
                stats['normalized_cells'] += 1
                
                # 正規化タイプの分析
                self._analyze_normalization_type(cell.text, cell.text_norm, stats['normalization_types'])
            else:
                stats['unchanged_cells'] += 1
        
        return stats
    
    def _analyze_normalization_type(self, original: str, normalized: str, type_stats: dict):
        """
        正規化タイプを分析
        
        Args:
            original: 元のテキスト
            normalized: 正規化後のテキスト
            type_stats: タイプ別統計辞書
        """
        # 各パターンの検出
        if self.patterns['hyphen'].search(original):
            type_stats['hyphen'] += 1
            
        if self.patterns['long_vowel'].search(original):
            type_stats['long_vowel'] += 1
            
        if self.patterns['wave_dash'].search(original):
            type_stats['wave_dash'] += 1
            
        if self.patterns['percent'].search(original):
            type_stats['percent'] += 1
            
        if '（' in original or '）' in original:
            type_stats['brackets'] += 1
            
        if self.patterns['dots'].search(original):
            type_stats['dots'] += 1
            
        if self.patterns['spaces'].search(original):
            type_stats['spaces'] += 1


def normalize_sample_text(text: str) -> str:
    """
    サンプルテキストの正規化を行う便利関数
    
    Args:
        text: テキスト
        
    Returns:
        正規化されたテキスト
    """
    normalizer = TextNormalizer()
    return normalizer.normalize_text(text)


if __name__ == "__main__":
    # テスト実行
    normalizer = TextNormalizer()
    
    # テストケース
    test_cases = [
        "ユーザー",
        "ユーザ",
        "データ・ベース",
        "データベース",
        "100％完了",
        "100%完了",
        "システム—管理",
        "システム-管理",
        "ログ・イン機能",
        "　　空白　　多い　　",
        "（全角括弧）",
        "(半角括弧)",
    ]
    
    print("=== 正規化テスト ===")
    for text in test_cases:
        normalized = normalizer.normalize_text(text)
        print(f"'{text}' → '{normalized}'" + (" [変更]" if text != normalized else " [変更なし]"))