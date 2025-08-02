"""
クロスファイル用語統一チェックモジュール
全ファイル・全シートから用語を収集し、表記の統一性をチェック
"""

import yaml
from typing import List, Optional, Dict, Any, Set, Tuple
from dataclasses import dataclass
from collections import defaultdict
from pathlib import Path
from loguru import logger

from .extract import CellData


@dataclass
class DetectionResult:
    """検出結果を格納するデータクラス"""
    file_name: str
    sheet_name: str
    cell_address: str
    row: int
    column: int
    original: str
    suggested_fix: str
    issue_type: str
    reason: str
    confidence: float
    canonical: Optional[str] = None
    
    def to_dict(self) -> Dict[str, Any]:
        """辞書形式に変換"""
        return {
            'file_name': self.file_name,
            'sheet_name': self.sheet_name,
            'cell_address': self.cell_address,
            'row': self.row,
            'column': self.column,
            'original': self.original,
            'suggested_fix': self.suggested_fix,
            'issue_type': self.issue_type,
            'reason': self.reason,
            'confidence': self.confidence,
            'canonical': self.canonical
        }


@dataclass
class TermOccurrence:
    """用語の出現情報を格納するデータクラス"""
    file_name: str
    sheet_name: str
    cell_address: str
    row: int
    column: int
    original_text: str


class MechanicalDetector:
    """
    クロスファイル用語統一チェッククラス
    全ファイル・全シートから用語を収集し、表記の統一性をチェック
    """
    
    def __init__(self, canonicals_path: str = "dict/canonicals.yml"):
        """初期化"""
        self.canonicals_path = canonicals_path
        self.canonical_rules = self._load_canonical_rules()
    
    def _load_canonical_rules(self) -> Dict[str, str]:
        """canonical rules を読み込み"""
        canonical_map = {}
        
        try:
            if Path(self.canonicals_path).exists():
                with open(self.canonicals_path, 'r', encoding='utf-8') as f:
                    data = yaml.safe_load(f)
                    
                if data and 'rules' in data:
                    for rule in data['rules']:
                        expected = rule.get('expected', '')
                        patterns = rule.get('patterns', [])
                        
                        # expected自体もマップに追加
                        canonical_map[expected.lower()] = expected
                        
                        # patternsもマップに追加
                        for pattern in patterns:
                            canonical_map[pattern.lower()] = expected
                            
                logger.info(f"Loaded {len(canonical_map)} canonical rules")
            else:
                logger.warning(f"Canonical rules file not found: {self.canonicals_path}")
                
        except Exception as e:
            logger.error(f"Error loading canonical rules: {e}")
            
        return canonical_map
    
    def detect_normalization_variants(self, cells: List[CellData]) -> List[DetectionResult]:
        """
        クロスファイル用語統一チェック
        
        Args:
            cells: セルデータのリスト
            
        Returns:
            検出結果のリスト
        """
        logger.info(f"Starting cross-file term consistency check for {len(cells)} cells")
        
        # 1. 用語の出現パターンを収集
        term_occurrences = self._collect_term_occurrences(cells)
        
        # 2. 表記ゆれを検出
        inconsistencies = self._detect_term_inconsistencies(term_occurrences)
        
        # 3. 検出結果を生成
        results = self._generate_detection_results(inconsistencies)
        
        logger.info(f"Cross-file consistency check completed: {len(results)} inconsistencies detected")
        return results
    
    def _collect_term_occurrences(self, cells: List[CellData]) -> Dict[str, List[TermOccurrence]]:
        """用語の出現パターンを収集（正規化前の元テキストから）"""
        term_occurrences = defaultdict(list)
        
        for cell in cells:
            # 正規化前の元テキストを使用
            original_text = cell.text
            if not original_text or original_text.strip() == "":
                continue
                
            # 元テキストから用語を抽出
            terms = self._extract_terms(original_text)
            
            for term in terms:
                if len(term.strip()) < 2:  # 短すぎる用語は除外
                    continue
                    
                # 抽出した用語をそのままキーとする（正規化なし）
                normalized_term = term
                
                occurrence = TermOccurrence(
                    file_name=cell.file_name,
                    sheet_name=cell.sheet_name,
                    cell_address=cell.cell_address,
                    row=cell.row,
                    column=cell.column,
                    original_text=term  # 元の用語を保存
                )
                
                term_occurrences[term].append(occurrence)
        
        # 単一出現の用語は除外（表記ゆれは複数出現が前提）
        filtered_occurrences = {
            term: occurrences 
            for term, occurrences in term_occurrences.items() 
            if len(occurrences) > 1
        }
        
        # デバッグ: 収集された用語の詳細出力
        for term, occurrences in filtered_occurrences.items():
            original_forms = [occ.original_text for occ in occurrences]
            logger.debug(f"Term '{term}': {original_forms}")
        
        logger.debug(f"Collected {len(filtered_occurrences)} terms with multiple occurrences")
        return filtered_occurrences
    
    def _extract_terms(self, text: str) -> List[str]:
        """テキストから用語を抽出（正規化前の元テキストから）"""
        import re
        
        # 複合語も考慮した用語抽出
        # まず全体を1つの用語として、その後で個別の構成要素も抽出
        terms = []
        
        # 1. 全体を1つの用語として追加（半角カタカナ ｦ-ﾟ を追加）
        if re.match(r'[ぁ-んァ-ヶｦ-ﾟ一-龯a-zA-Z0-9・\-_ー]+$', text.strip()):
            terms.append(text.strip())
        
        # 2. 構成要素を分割して抽出（文字種別ごとに分割）
        # ひらがな・全角カタカナ・半角カタカナ・漢字・英数字を別々に抽出
        patterns = [
            r'[ぁ-ん]+',      # ひらがな
            r'[ァ-ヶー]+',     # 全角カタカナ
            r'[ｦ-ﾟ]+',       # 半角カタカナ
            r'[一-龯]+',      # 漢字
            r'[a-zA-Z0-9]+',  # 英数字
        ]
        
        for pattern in patterns:
            components = re.findall(pattern, text)
            for component in components:
                if len(component) >= 2:  # 2文字以上の構成要素のみ
                    terms.append(component)
        
        # デバッグ出力
        if terms:
            logger.debug(f"Extracted terms from '{text}': {terms}")
        
        return terms
    
    def _detect_term_inconsistencies(self, term_occurrences: Dict[str, List[TermOccurrence]]) -> Dict[str, Dict[str, Any]]:
        """表記ゆれを検出"""
        inconsistencies = {}
        
        for term, occurrences in term_occurrences.items():
            # 異なる表記のバリエーションを収集
            variants = defaultdict(list)
            for occurrence in occurrences:
                variants[occurrence.original_text].append(occurrence)
            
            # 複数の表記がある場合のみ不統一として検出
            if len(variants) > 1:
                inconsistencies[term] = {
                    'variants': dict(variants),
                    'canonical': self._get_canonical_form(term, list(variants.keys()))
                }
        
        logger.debug(f"Detected {len(inconsistencies)} terms with inconsistent notation")
        return inconsistencies
    
    def _get_canonical_form(self, term: str, variant_list: List[str]) -> Optional[str]:
        """正準形を取得"""
        # variant_listの各項目がcanonical_rulesにあるかチェック（大文字小文字区別あり）
        for variant in variant_list:
            if variant in self.canonical_rules:
                return self.canonical_rules[variant]
            # 小文字でもチェック（英単語用）
            if variant.lower() in self.canonical_rules:
                return self.canonical_rules[variant.lower()]
        
        # termでもチェック
        if term in self.canonical_rules:
            return self.canonical_rules[term]
        
        return None
    
    def _generate_detection_results(self, inconsistencies: Dict[str, Dict[str, Any]]) -> List[DetectionResult]:
        """検出結果を生成"""
        results = []
        
        for term, inconsistency_data in inconsistencies.items():
            variants = inconsistency_data['variants']
            canonical = inconsistency_data['canonical']
            
            if canonical:
                # canonicals.ymlに定義がある場合：具体的な修正提案
                for variant_text, occurrences in variants.items():
                    if variant_text != canonical:  # 正準形以外の表記
                        for occurrence in occurrences:
                            result = DetectionResult(
                                file_name=occurrence.file_name,
                                sheet_name=occurrence.sheet_name,
                                cell_address=occurrence.cell_address,
                                row=occurrence.row,
                                column=occurrence.column,
                                original=variant_text,
                                suggested_fix=canonical,
                                issue_type="variant",
                                reason=f"表記統一ルールに従い「{canonical}」に統一",
                                confidence=0.90,
                                canonical=canonical
                            )
                            results.append(result)
            else:
                # canonicals.ymlに定義がない場合：不統一の指摘のみ
                # 最初の出現箇所にまとめて報告
                first_variant = list(variants.keys())[0]
                first_occurrence = variants[first_variant][0]
                
                # 出現箇所一覧を作成
                locations = []
                for variant_text, occurrences in variants.items():
                    for occ in occurrences:
                        locations.append(f"{occ.file_name}:{occ.sheet_name}:{occ.cell_address}({variant_text})")
                
                locations_str = ", ".join(locations[:5])  # 最初の5箇所のみ表示
                if len(locations) > 5:
                    locations_str += f" など{len(locations)}箇所"
                
                result = DetectionResult(
                    file_name=first_occurrence.file_name,
                    sheet_name=first_occurrence.sheet_name,
                    cell_address=first_occurrence.cell_address,
                    row=first_occurrence.row,
                    column=first_occurrence.column,
                    original=first_variant,
                    suggested_fix=None,
                    issue_type="variant",
                    reason=f"表記が統一されていません。出現箇所: {locations_str}",
                    confidence=0.85,
                    canonical=None
                )
                results.append(result)
        
        return results
    
    
    def get_detection_statistics(self, results: List[DetectionResult]) -> Dict[str, Any]:
        """
        検出統計情報を取得
        
        Args:
            results: 検出結果のリスト
            
        Returns:
            統計情報の辞書
        """
        stats = {
            'total_detections': len(results),
            'variant_count': sum(1 for r in results if r.issue_type == 'variant'),
            'avg_confidence': sum(r.confidence for r in results) / len(results) if results else 0,
            'detection_types': {}
        }
        
        # 検出タイプ別の統計
        canonical_count = sum(1 for r in results if r.canonical)
        inconsistency_count = sum(1 for r in results if not r.canonical)
        
        stats['detection_types'] = {
            '統一ルール適用': canonical_count,
            '表記不統一指摘': inconsistency_count
        }
        
        return stats


def test_cross_file_detector():
    """クロスファイル用語統一チェックのテスト関数"""
    detector = MechanicalDetector()
    
    # テストケース（複数ファイル・シートの用語統一チェック）
    test_cases = [
        # ファイル1
        ("file1.xlsx", "Sheet1", "A1", "ユーザー管理"),
        ("file1.xlsx", "Sheet1", "A2", "ユーザ登録"),
        ("file1.xlsx", "Sheet2", "B1", "データベース設計"),
        
        # ファイル2  
        ("file2.xlsx", "Sheet1", "C1", "ユーザ情報"),
        ("file2.xlsx", "Sheet1", "C2", "DB接続"),
        ("file2.xlsx", "Sheet2", "D1", "ログイン機能"),
        ("file2.xlsx", "Sheet2", "D2", "ログ・イン画面"),
        
        # canonicals.ymlでマッピングされる用語
        ("file3.xlsx", "Sheet1", "E1", "ﾕｰｻﾞ設定"),  # ユーザー
        ("file3.xlsx", "Sheet1", "E2", "ユーザ一覧"),
        ("file3.xlsx", "Sheet2", "F1", "データ・ベース操作"),  # データベース
    ]
    
    print("=== クロスファイル用語統一チェック テスト ===")
    
    from .extract import CellData
    
    # テストセルデータを作成
    test_cells = []
    for i, (file_name, sheet_name, cell_address, text) in enumerate(test_cases):
        cell = CellData(
            file_name=file_name,
            sheet_name=sheet_name,
            cell_address=cell_address,
            row=i+1,
            column=1,
            text=text
        )
        test_cells.append(cell)
    
    print(f"Testing with {len(test_cells)} cells")
    
    # 検出実行
    results = detector.detect_normalization_variants(test_cells)
    
    # 結果表示
    print(f"\n検出結果 ({len(results)}件):")
    for result in results:
        if result.suggested_fix:
            print(f"  [{result.file_name}:{result.sheet_name}:{result.cell_address}] '{result.original}' → '{result.suggested_fix}'")
            print(f"    理由: {result.reason}")
        else:
            print(f"  [{result.file_name}:{result.sheet_name}:{result.cell_address}] '{result.original}'")
            print(f"    理由: {result.reason}")
        print()
    
    # 統計表示
    stats = detector.get_detection_statistics(results)
    print(f"統計: {stats['total_detections']}件検出")
    for reason, count in stats['detection_types'].items():
        print(f"  {reason}: {count}件")


if __name__ == "__main__":
    test_cross_file_detector()