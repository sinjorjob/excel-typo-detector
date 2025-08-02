#!/usr/bin/env python3
"""
Excel表記ゆれ・誤字検出ツール
メイン実行ファイル
"""

import sys
import argparse
import os
from pathlib import Path
from typing import Dict, List, Any, Optional
import time

import yaml
from loguru import logger
from dotenv import load_dotenv

# .envファイルを読み込み
load_dotenv()

# プロジェクトのsrcディレクトリをパスに追加
sys.path.insert(0, str(Path(__file__).parent / "src"))

from src.extract import ExcelExtractor
from src.mechanical_detector import MechanicalDetector
from src.llm import LLMReviewer
from src.report import ReportGenerator
# アルゴリズム検出機能は削除済み - LLMレビューのみ使用


class ExcelChecker:
    """Excel校正ツールのメインクラス"""
    
    def __init__(self, config_path: str = "config.yml"):
        """
        初期化
        
        Args:
            config_path: 設定ファイルパス
        """
        self.config = self._load_config(config_path)
        self.start_time = time.time()
        
        # ログ設定
        self._setup_logging()
        
        # 各モジュールの初期化
        self.extractor = None
        self.mechanical_detector = None
        self.llm_reviewer = None
        self.reporter = None
        
        logger.info("Excel Checker initialized")
    
    def _load_config(self, config_path: str) -> Dict[str, Any]:
        """設定ファイルを読み込み"""
        try:
            with open(config_path, "r", encoding="utf-8") as f:
                config = yaml.safe_load(f)
            logger.info(f"Configuration loaded from {config_path}")
            return config
        except Exception as e:
            logger.error(f"Failed to load configuration: {e}")
            sys.exit(1)
    
    def _setup_logging(self):
        """ログ設定"""
        log_config = self.config.get('logging', {})
        log_level = log_config.get('level', 'INFO')
        log_file = log_config.get('file', 'logs/excel_checker.log')
        
        # ログディレクトリ作成
        log_dir = Path(log_file).parent
        log_dir.mkdir(parents=True, exist_ok=True)
        
        # ログ設定をクリア
        logger.remove()
        
        # コンソール出力
        logger.add(
            sys.stdout, 
            level=log_level,
            format="<green>{time:YYYY-MM-DD HH:mm:ss}</green> | <level>{level: <8}</level> | {message}"
        )
        
        # ファイル出力
        logger.add(
            log_file,
            level=log_level,
            format="{time:YYYY-MM-DD HH:mm:ss} | {level: <8} | {name} | {message}",
            rotation="10 MB",
            retention="30 days",
            encoding="utf-8"
        )
    
    def initialize_modules(self):
        """各モジュールを初期化"""
        logger.info("Initializing modules...")
        
        # 1. Extractor
        self.extractor = ExcelExtractor(self.config)
        logger.debug("Excel extractor initialized")
        
        # 2. Mechanical Detector
        self.mechanical_detector = MechanicalDetector()
        logger.debug("Mechanical detector initialized")
        
        # 3. LLM Reviewer (メイン機能)
        llm_config = self.config.get('llm', {})
        if llm_config.get('enabled', False):
            try:
                self.llm_reviewer = LLMReviewer(self.config)
                logger.debug("LLM reviewer initialized")
            except Exception as e:
                logger.error(f"LLM reviewer initialization failed: {e}")
                logger.error("LLM reviewer is required for this tool")
                sys.exit(1)
        else:
            logger.error("LLM is disabled in config, but it's required for this tool")
            sys.exit(1)
        
        # 4. Reporter
        output_dir = self.config.get('output', {}).get('dir', 'data/output')
        self.reporter = ReportGenerator(self.config, output_dir)
        logger.debug("Report generator initialized")
        
        logger.info("All modules initialized successfully")
    
    def process_files(self, input_path: str, enable_llm: bool = True) -> Dict[str, Any]:
        """
        ファイル処理のメイン処理
        
        Args:
            input_path: 入力パス（ファイルまたはディレクトリ）
            enable_llm: LLMレビューを有効にするか
            
        Returns:
            処理結果の辞書
        """
        logger.info(f"Starting file processing: {input_path}")
        
        # 1. ファイル抽出
        logger.info("Step 1/6: Extracting cells from Excel files...")
        if Path(input_path).is_file():
            cells = self.extractor.extract_from_file(input_path)
        else:
            cells = self.extractor.extract_from_directory(input_path)
        
        if not cells:
            logger.error("No cells extracted. Check input files.")
            return {'error': 'No cells extracted'}
        
        logger.info(f"Extracted {len(cells)} cells")
        
        # 2. 機械的検出（全角/半角表記ゆれ）
        logger.info("Step 2/6: Mechanical detection of notation variants...")
        mechanical_results = self.mechanical_detector.detect_normalization_variants(cells)
        mechanical_stats = self.mechanical_detector.get_detection_statistics(mechanical_results)
        logger.info(f"Mechanical detection found {len(mechanical_results)} notation variants")
        
        # 3. LLMレビュー（メイン処理）
        llm_results = []
        llm_stats = {}
        
        if enable_llm and self.llm_reviewer:
            logger.info("Step 3/6: LLM review of all cells...")
            try:
                llm_results = self.llm_reviewer.review_all_cells(cells)
                llm_stats = self.llm_reviewer.get_review_statistics(llm_results)
                logger.info(f"LLM reviewed {len(cells)} cells, found {len(llm_results)} issues")
            except Exception as e:
                logger.error(f"LLM review failed: {e}")
                raise
        else:
            logger.error("LLM review is required but not available")
            raise RuntimeError("LLM reviewer is required")
        
        # 4. 結果の統合（重複除去）
        logger.info("Step 4/6: Merging detection results...")
        all_results = self._merge_detection_results(mechanical_results, llm_results)
        logger.info(f"Total unique issues after merging: {len(all_results)}")
        
        # 5. レポート生成
        logger.info("Step 5/6: Generating reports...")
        
        # 統計情報
        combined_stats = {
            'mechanical_detection': mechanical_stats,
            'llm_review': llm_stats,
            'processing_time': time.time() - self.start_time
        }
        
        generated_files = self.reporter.generate_all_reports(
            [], # アルゴリズム検出結果は空
            all_results,  # 統合された結果
            combined_stats
        )
        
        logger.info(f"Generated {len(generated_files)} report files")
        
        # 6. 完了
        total_time = time.time() - self.start_time
        logger.info(f"Step 6/6: Processing completed in {total_time:.2f} seconds")
        
        # 結果サマリー
        result_summary = {
            'input_path': input_path,
            'total_cells': len(cells),
            'mechanical_detections': len(mechanical_results),
            'llm_reviews': len(llm_results),
            'total_issues': len(all_results),
            'generated_files': generated_files,
            'statistics': combined_stats,
            'processing_time': total_time
        }
        
        return result_summary
    
    def _merge_detection_results(self, mechanical_results: List, llm_results: List) -> List:
        """
        機械的検出とLLM検出の結果を統合し、重複を除去
        
        Args:
            mechanical_results: 機械的検出結果（DetectionResult型）
            llm_results: LLM検出結果（LLMReviewResponse型）
            
        Returns:
            統合された検出結果（重複除去済み）
        """
        # 重複判定用のキーを作成する関数
        def create_mechanical_key(result):
            return f"{result.file_name}:{result.sheet_name}:{result.cell_address}:{result.original}"
        
        def create_llm_key(result):
            # LLMReviewResponseの場合、source_detectionから情報を取得
            if hasattr(result, 'source_detection') and result.source_detection:
                cell = result.source_detection.cell_data
                return f"{cell.file_name}:{cell.sheet_name}:{cell.cell_address}:{result.original}"
            else:
                # source_detectionがない場合はoriginalのみでキー生成
                return f"unknown:unknown:unknown:{result.original}"
        
        # 機械的検出の結果をマップに格納
        merged_map = {}
        for result in mechanical_results:
            key = create_mechanical_key(result)
            merged_map[key] = result
            logger.debug(f"Added mechanical result: {key}")
        
        # LLM検出結果を追加（重複は上書き）
        for result in llm_results:
            key = create_llm_key(result)
            if key in merged_map:
                # 機械的検出と同じセル・テキストの場合
                existing = merged_map[key]
                if existing.issue_type == "variant" and result.issue_type == "typo":
                    # 機械的検出が表記ゆれ、LLMが誤字の場合、LLMを優先
                    merged_map[key] = result
                    logger.debug(f"LLM overrode mechanical result: {key}")
                else:
                    # それ以外は機械的検出を優先（高信頼度のため）
                    logger.debug(f"Kept mechanical result over LLM: {key}")
            else:
                # 新規のLLM検出結果
                merged_map[key] = result
                logger.debug(f"Added LLM result: {key}")
        
        return list(merged_map.values())
    
    def print_summary(self, results: Dict[str, Any]):
        """処理結果のサマリーを表示"""
        print("\n" + "="*60)
        print("           LLM誤字脱字検出 処理結果")
        print("="*60)
        
        if 'error' in results:
            print(f"❌ エラー: {results['error']}")
            return
        
        stats = results.get('statistics', {})
        
        # 基本情報
        print(f"[FILES] 入力パス: {results['input_path']}")
        print(f"[TIME]  処理時間: {results['processing_time']:.2f}秒")
        print(f"[DATA]  セル数: {results['total_cells']:,}件")
        print()
        
        # 機械的検出結果
        mechanical_stats = stats.get('mechanical_detection', {})
        if results['mechanical_detections'] > 0:
            print("[MECHANICAL] 機械的検出結果:")
            print(f"  表記ゆれ検出: {results['mechanical_detections']:,}件")
            print(f"  平均信頼度: {mechanical_stats.get('avg_confidence', 0):.2f}")
            
            # 検出タイプ別
            detection_types = mechanical_stats.get('detection_types', {})
            for dtype, count in detection_types.items():
                print(f"  {dtype}: {count}件")
            print()
        else:
            print("[MECHANICAL] 機械的な表記ゆれは見つかりませんでした")
            print()
        
        # LLMレビュー結果
        llm_stats = stats.get('llm_review', {})
        if results['llm_reviews'] > 0:
            print("[LLM] LLMレビュー結果:")
            print(f"  修正要項目: {results['llm_reviews']:,}件")
            print(f"  平均信頼度: {llm_stats.get('avg_confidence', 0):.2f}")
            print(f"  API リクエスト数: {llm_stats.get('request_count', 0):,}回")
            
            # 問題タイプ別
            issue_counts = llm_stats.get('issue_type_counts', {})
            print(f"  誤字: {issue_counts.get('typo', 0):,}件")
            print(f"  表記ゆれ: {issue_counts.get('variant', 0):,}件")
            print(f"  高信頼度修正: {llm_stats.get('high_confidence', 0):,}件")
            print()
        else:
            print("[LLM] 修正が必要な項目は見つかりませんでした")
            print()
        
        # 統合結果
        print(f"[TOTAL] 重複除去後の修正要項目: {results['total_issues']:,}件")
        print()
        
        # 生成ファイル
        print("[OUTPUT] 生成ファイル:")
        for file_type, file_path in results['generated_files'].items():
            file_name = Path(file_path).name
            print(f"  {file_type}: {file_name}")
        
        print("="*60)
        print("[COMPLETE] 処理が完了しました！")


def parse_arguments():
    """コマンドライン引数をパース"""
    parser = argparse.ArgumentParser(
        description="Excel表記ゆれ・誤字検出ツール",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog="""
使用例:
  python main.py sample_data/                    # ディレクトリ内のExcelファイルを処理
  python main.py sample_data/test.xlsx           # 単一ファイルを処理
  python main.py sample_data/ --no-llm           # LLMレビューなしで処理
  python main.py sample_data/ --config my.yml   # カスタム設定で処理
        """
    )
    
    parser.add_argument(
        'input',
        help='入力ファイルまたはディレクトリパス'
    )
    
    parser.add_argument(
        '--config', '-c',
        default='config.yml',
        help='設定ファイルパス（デフォルト: config.yml）'
    )
    
    parser.add_argument(
        '--no-llm',
        action='store_true',
        help='LLMレビューを無効にする'
    )
    
    parser.add_argument(
        '--output-dir', '-o',
        help='出力ディレクトリ（設定ファイルを上書き）'
    )
    
    parser.add_argument(
        '--log-level',
        choices=['DEBUG', 'INFO', 'WARNING', 'ERROR'],
        default='INFO',
        help='ログレベル（デフォルト: INFO）'
    )
    
    parser.add_argument(
        '--version', '-v',
        action='version',
        version='Excel Checker v1.0.0'
    )
    
    return parser.parse_args()


def validate_input(input_path: str) -> bool:
    """入力パスの妥当性をチェック"""
    path = Path(input_path)
    
    if not path.exists():
        logger.error(f"Input path does not exist: {input_path}")
        return False
    
    if path.is_file():
        if not path.suffix.lower() in ['.xlsx', '.xls']:
            logger.error(f"Input file is not an Excel file: {input_path}")
            return False
    elif path.is_dir():
        excel_files = list(path.glob("*.xlsx")) + list(path.glob("*.xls"))
        if not excel_files:
            logger.error(f"No Excel files found in directory: {input_path}")
            return False
    
    return True


def main():
    """メイン関数"""
    try:
        # コマンドライン引数をパース
        args = parse_arguments()
        
        # 入力検証
        if not validate_input(args.input):
            sys.exit(1)
        
        # 設定ファイル検証
        if not Path(args.config).exists():
            logger.error(f"Configuration file not found: {args.config}")
            sys.exit(1)
        
        # Excel Checkerを初期化
        checker = ExcelChecker(args.config)
        
        # 出力ディレクトリの上書き
        if args.output_dir:
            checker.config['output']['dir'] = args.output_dir
        
        # ログレベルの上書き
        if args.log_level:
            checker.config['logging']['level'] = args.log_level
            checker._setup_logging()
        
        # モジュール初期化
        checker.initialize_modules()
        
        # 処理実行
        results = checker.process_files(
            args.input, 
            enable_llm=not args.no_llm
        )
        
        # 結果表示
        checker.print_summary(results)
        
        # 正常終了
        sys.exit(0)
        
    except KeyboardInterrupt:
        logger.info("Processing interrupted by user")
        sys.exit(1)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")
        sys.exit(1)


if __name__ == "__main__":
    main()