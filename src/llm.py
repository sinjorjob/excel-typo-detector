"""
LLM連携モジュール
迷う候補のみを高精度でLLM判定する
"""

import json
import os
from typing import List, Dict, Any, Optional
from dataclasses import dataclass, asdict, field
from enum import Enum
from datetime import datetime

import anthropic
import openai
from loguru import logger

from .extract import CellData


# 以前detect.pyにあった定義を移行
class IssueType(Enum):
    """問題タイプ"""
    VARIANT = "variant"  # 表記ゆれ
    TYPO = "typo"       # 誤字・誤変換
    NONE = "none"       # 問題なし


@dataclass
class DetectionResult:
    """検出結果（互換性のため残す）"""
    cell_data: CellData
    issue_type: IssueType
    original: str
    suggested_fix: Optional[str] = None
    canonical: Optional[str] = None
    confidence: float = 0.0
    reason: str = ""
    auto_fix: bool = False
    context: str = ""
    related_terms: List[str] = field(default_factory=list)


@dataclass
class LLMReviewRequest:
    """LLMレビューリクエスト"""
    context: str
    original: str
    canonical: Optional[str]
    related: List[str]


@dataclass 
class LLMReviewResponse:
    """LLMレビュー応答"""
    issue_type: str  # "variant", "typo", "none"
    original: str
    suggested_fix: Optional[str]
    canonical: Optional[str]
    reason: str
    confidence: float
    source_detection: Optional['DetectionResult'] = None  # 元の検出結果への参照
    
    def to_dict(self) -> Dict[str, Any]:
        """辞書形式に変換"""
        data = asdict(self)
        # source_detectionは辞書化から除外
        if 'source_detection' in data:
            del data['source_detection']
        return data


class LLMProvider(Enum):
    """LLMプロバイダー"""
    ANTHROPIC = "anthropic"
    OPENAI = "openai"


class LLMReviewer:
    """LLMレビュークラス"""
    
    def __init__(self, config: Dict[str, Any]):
        """
        初期化
        
        Args:
            config: 設定辞書
        """
        self.config = config
        self.llm_config = config.get('llm', {})
        
        # 辞書情報を読み込み
        self.dictionary_rules = self._load_dictionary_rules()
        
        # プロバイダー設定
        provider_name = self.llm_config.get('provider', 'anthropic')
        self.provider = LLMProvider(provider_name)
        
        # モデル設定
        self.model = self.llm_config.get('model', 'claude-3-sonnet-20240229')
        self.max_tokens = self.llm_config.get('max_tokens', 1000)
        self.temperature = self.llm_config.get('temperature', 0.1)
        self.batch_size = self.llm_config.get('batch_size', 20)  # シート単位では使用しない（後方互換性のため保持）
        self.min_confidence = self.llm_config.get('min_confidence', 0.70)
        
        # APIクライアントの初期化
        self._init_api_client()
        
        # プロンプトテンプレート
        self.system_prompt = self._get_system_prompt()
        
        # LLMリクエストログの初期化
        self._init_request_logger()
        
        # リクエスト回数カウンター
        self.request_count = 0
    
    def _load_dictionary_rules(self) -> str:
        """
        辞書ファイルから正準形ルールを読み込んでプロンプト用文字列を生成
        
        Returns:
            プロンプトに含める辞書情報の文字列
        """
        try:
            import yaml
            import os
            
            canonicals_path = os.path.join(os.path.dirname(os.path.dirname(__file__)), 'dict', 'canonicals.yml')
            
            if not os.path.exists(canonicals_path):
                logger.warning(f"Canonicals file not found: {canonicals_path}")
                return ""
            
            with open(canonicals_path, 'r', encoding='utf-8') as f:
                data = yaml.safe_load(f)
            
            if not data or 'rules' not in data:
                return ""
            
            rules_text = ["プロジェクト表記統一ルール:"]
            for rule in data['rules']:
                expected = rule.get('expected', '')
                patterns = rule.get('patterns', [])
                if expected and patterns:
                    patterns_str = ', '.join(f'"{p}"' for p in patterns)
                    rules_text.append(f"- {patterns_str} → \"{expected}\"")
            
            return '\n'.join(rules_text)
            
        except Exception as e:
            logger.warning(f"Failed to load dictionary rules: {e}")
            return ""
    
    def _init_api_client(self):
        """APIクライアントを初期化"""
        try:
            if self.provider == LLMProvider.ANTHROPIC:
                api_key = os.getenv('ANTHROPIC_API_KEY')
                if not api_key:
                    logger.error("ANTHROPIC_API_KEY not found in environment")
                    self.client = None
                    return
                self.client = anthropic.Anthropic(api_key=api_key)
                
            elif self.provider == LLMProvider.OPENAI:
                api_key = os.getenv('OPENAI_API_KEY')
                if not api_key:
                    logger.error("OPENAI_API_KEY not found in environment")
                    self.client = None
                    return
                # OpenAIクライアントを単純に初期化
                self.client = openai.OpenAI(api_key=api_key)
                
            logger.info(f"Initialized {self.provider.value} client")
            
        except Exception as e:
            logger.error(f"Failed to initialize LLM client: {e}")
            logger.debug(f"Error details: {str(e)}")
            self.client = None
    
    def _init_request_logger(self):
        """LLMリクエスト専用ログの初期化"""
        llm_log_path = self.config.get('logging', {}).get('llm_requests', 'data/output/llm_requests.log')
        
        # ログディレクトリを作成
        os.makedirs(os.path.dirname(llm_log_path), exist_ok=True)
        
        # LLM専用ロガーを追加
        logger.add(
            llm_log_path,
            level="INFO",
            format="{time:YYYY-MM-DD HH:mm:ss} | {message}",
            rotation="10 MB",
            retention="30 days",
            encoding="utf-8",
            filter=lambda record: record["extra"].get("llm_request", False)
        )
        
        logger.info("LLM request logging initialized")
    
    def _log_llm_request(self, request: 'LLMReviewRequest', user_prompt: str, response: str = None, error: str = None):
        """LLMリクエストをログに記録"""
        try:
            cell_info = f"{request.cell_data.file_name}:{request.cell_data.sheet_name}:{request.cell_data.cell_address}" if hasattr(request, 'cell_data') and request.cell_data else "位置不明"
            
            log_entry = {
                "timestamp": datetime.now().isoformat(),
                "provider": self.provider.value,
                "model": self.model,
                "cell_location": cell_info,
                "system_prompt": self.system_prompt,
                "user_prompt": user_prompt,
                "response": response,
                "error": error
            }
            
            # JSON形式でログ出力
            logger.bind(llm_request=True).info(json.dumps(log_entry, ensure_ascii=False, indent=2))
            
        except Exception as e:
            logger.warning(f"Failed to log LLM request: {e}")
    
    def _get_system_prompt(self) -> str:
        """システムプロンプトを取得"""
        dictionary_info = self.dictionary_rules
        
        base_prompt = """あなたはシステム設計書専門の誤字脱字検出アシスタントです。
以下の思考プロセスに従って、誤字・誤変換・脱字を検出し、適切な修正案を提供してください。

【思考プロセス】
1. 語句の意味理解: その語句がシステム設計書の文脈で意味をなすか判断
2. 同音異義語チェック: 似た音の別の語句が文脈により適切ではないか検証
3. 専門用語確認: ビジネス・IT・会計分野の専門用語として正しいか評価
4. 完全性検証: 語句が完全で、文字の欠落や余分な文字がないか確認
5. 修正案生成: 文脈に最も適した完全で自然な語句を提案

【判定基準】
- typo: システム設計書の文脈で不自然・不適切な語句
- none: 文脈に適した正しい語句、または表記統一範囲

【修正案生成の原則】
✓ 文脈適合性: システム設計書として最も自然な語句を選択
✓ 完全性: 部分修正ではなく、完全で意味の通る語句に修正  
✓ 専門性: ビジネス・IT分野の適切な専門用語を使用
✓ 論理性: 前後の文脈や項目との整合性を考慮

【よくある誤変換パターンの判定方法】
- 同音異義語: 文脈で意味が通らない場合、類似音の適切な語句を検討
- 業務用語: ビジネス・会計・IT分野の専門用語として不自然な場合を特定
- 文字混入: 余分な文字や間違った文字が混入していないか確認
- 欠落補完: 意味を完成させるために必要な文字が欠けていないか確認

重要な制限事項:
- 技術仕様（DB設計、API仕様等）でのアルファベット表記は変更しない
- テーブル名・カラム名・関数名・変数名等の技術用語は英語のまま維持  
- 文書内容・説明文でのみ表記統一ルールを適用
- 信頼度0.7以上のもののみ出力
- 推測ではなく、明確に不適切と判断できる場合のみ修正提案"""

        if dictionary_info:
            return f"{base_prompt}\n\n【表記統一ルール - 文書内容・説明文のみ適用】\n技術仕様書の説明文や業務説明で以下のルールに該当する場合のみ修正提案：\n{dictionary_info}\n\n注意：テーブル名・カラム名・API名等の技術用語は英語のまま維持してください。"
        else:
            return base_prompt
    
    def _get_user_prompt(self, request: LLMReviewRequest) -> str:
        """ユーザープロンプトを生成"""
        related_str = ', '.join(request.related) if request.related else '[]'
        canonical_str = request.canonical if request.canonical else 'null'
        
        prompt = f"""【文脈】{request.context}
【候補語】{request.original}
【正準（あれば）】{canonical_str}
【関連語】{related_str}

JSONスキーマ:
{{
  "issue_type": "variant" | "typo" | "none",
  "original": "string",
  "suggested_fix": "string | null",
  "canonical": "string | null", 
  "reason": "string",
  "confidence": 0.0
}}"""
        
        return prompt
    
    def review_all_cells(self, cells: List[CellData]) -> List[LLMReviewResponse]:
        """
        全セルをシート単位でLLMレビュー（誤字脱字検出）
        
        Args:
            cells: セルデータのリスト
            
        Returns:
            LLMレビュー結果のリスト
        """
        if not self.client:
            logger.error("LLM client not initialized")
            return []
        
        # 空のセルや短すぎるセルを除外
        valid_cells = [cell for cell in cells if cell.text and len(cell.text.strip()) >= 2]
        
        if not valid_cells:
            logger.info("No valid cells found for LLM review")
            return []
        
        logger.info(f"Reviewing {len(valid_cells)} cells with LLM (sheet-based batching)")
        
        # シート単位でグルーピング
        sheet_groups = self._group_cells_by_sheet(valid_cells)
        
        all_responses = []
        total_requests = len(sheet_groups)
        
        for i, (sheet_key, sheet_cells) in enumerate(sheet_groups.items(), 1):
            logger.info(f"Processing sheet {i}/{total_requests}: {sheet_key} ({len(sheet_cells)} cells)")
            
            # シート内のセルを一度にLLMに送信
            sheet_responses = self._process_cell_batch(sheet_cells)
            all_responses.extend(sheet_responses)
            
            logger.debug(f"Completed sheet {sheet_key}: {len(sheet_responses)} issues found")
        
        logger.info(f"Completed LLM review: {total_requests} requests, {len(all_responses)} total issues")
        return all_responses
    
    def _group_cells_by_sheet(self, cells: List[CellData]) -> Dict[str, List[CellData]]:
        """
        セルをシート単位でグルーピング
        
        Args:
            cells: セルデータのリスト
            
        Returns:
            シート単位でグルーピングされたセル辞書
        """
        from collections import defaultdict
        
        sheet_groups = defaultdict(list)
        
        for cell in cells:
            # ファイル名:シート名 でグルーピング
            sheet_key = f"{cell.file_name}:{cell.sheet_name}"
            sheet_groups[sheet_key].append(cell)
        
        # セル数でソート（小さいシートから処理）
        sorted_groups = dict(sorted(sheet_groups.items(), key=lambda x: len(x[1])))
        
        return sorted_groups

    def review_detection_results(self, results: List[DetectionResult]) -> List[LLMReviewResponse]:
        """
        検出結果をLLMでレビュー（従来方式・後方互換性のため残す）
        
        Args:
            results: 検出結果のリスト
            
        Returns:
            LLMレビュー結果のリスト
        """
        if not self.client:
            logger.error("LLM client not initialized")
            return []
        
        # 迷う候補を選別
        ambiguous_results = self._select_ambiguous_cases(results)
        
        if not ambiguous_results:
            logger.info("No ambiguous cases found for LLM review")
            return []
        
        logger.info(f"Reviewing {len(ambiguous_results)} ambiguous cases with LLM")
        
        # バッチ処理
        all_responses = []
        
        for i in range(0, len(ambiguous_results), self.batch_size):
            batch = ambiguous_results[i:i + self.batch_size]
            batch_responses = self._process_batch(batch)
            all_responses.extend(batch_responses)
            
            logger.debug(f"Processed batch {i//self.batch_size + 1}/{(len(ambiguous_results)-1)//self.batch_size + 1}")
        
        return all_responses
    
    def _select_ambiguous_cases(self, results: List[DetectionResult]) -> List[DetectionResult]:
        """
        全セルをLLMレビューの対象とする（汎用的な誤字検出のため）
        
        Args:
            results: 検出結果のリスト
            
        Returns:
            レビュー対象の候補リスト（全セル）
        """
        # 全セルをLLMレビューの対象とする
        # 明らかに問題ないもの以外は全てレビュー
        ambiguous = []
        
        for result in results:
            # 極めて高い信頼度（0.99以上）の表記ゆれのみ除外
            if result.issue_type == IssueType.VARIANT and result.confidence >= 0.99:
                continue
            
            # その他は全てLLMでレビュー
            ambiguous.append(result)
        
        logger.info(f"Selected {len(ambiguous)} cases for LLM review (comprehensive approach)")
        return ambiguous
    
    
    def _process_cell_batch(self, batch: List[CellData]) -> List[LLMReviewResponse]:
        """
        セルバッチの直接処理 - 複数セルを1つのLLMリクエストで処理
        
        Args:
            batch: セルデータのバッチ
            
        Returns:
            LLMレビュー結果のリスト
        """
        if not batch:
            return []
        
        try:
            # 1回のLLMコールで全セルを処理
            responses = self._call_llm_for_cells(batch)
            return responses
            
        except Exception as e:
            logger.error(f"Error processing cell batch: {e}")
            return []

    def _process_batch(self, batch: List[DetectionResult]) -> List[LLMReviewResponse]:
        """
        真のバッチ処理 - 複数候補を1つのLLMリクエストで処理
        
        Args:
            batch: 検出結果のバッチ
            
        Returns:
            LLMレビュー結果のリスト
        """
        if not batch:
            return []
        
        try:
            # バッチリクエストを作成
            batch_requests = []
            for i, result in enumerate(batch):
                request = LLMReviewRequest(
                    context=result.context,
                    original=result.original,
                    canonical=result.canonical,
                    related=result.related_terms
                )
                request.cell_data = result.cell_data
                request.batch_index = i  # バッチ内のインデックス
                batch_requests.append(request)
            
            # 1回のLLMコールで全候補を処理
            responses = self._call_llm_batch(batch_requests, batch)
            return responses
            
        except Exception as e:
            logger.error(f"Error processing batch: {e}")
            return []
    
    def _call_llm_batch(self, batch_requests: List[LLMReviewRequest], source_detections: List[DetectionResult]) -> List[LLMReviewResponse]:
        """
        バッチでLLM APIを呼び出し
        
        Args:
            batch_requests: レビューリクエストのリスト
            source_detections: 元の検出結果のリスト
            
        Returns:
            LLMレビュー結果のリスト
        """
        if not batch_requests:
            return []
        
        user_prompt = None
        response_text = None
        
        try:
            # バッチ用のプロンプトを生成
            user_prompt = self._get_batch_user_prompt(batch_requests)
            
            # リクエスト回数をカウント（バッチでも1回）
            self.request_count += 1
            
            if self.provider == LLMProvider.ANTHROPIC:
                response_text = self._call_anthropic(user_prompt)
            elif self.provider == LLMProvider.OPENAI:
                response_text = self._call_openai(user_prompt)
            else:
                error_msg = f"Unsupported provider: {self.provider}"
                logger.error(error_msg)
                self._log_batch_llm_request(batch_requests, user_prompt, error=error_msg)
                return []
            
            # 成功時のログ記録
            self._log_batch_llm_request(batch_requests, user_prompt, response_text)
            
            return self._parse_batch_llm_response(response_text, batch_requests, source_detections)
            
        except Exception as e:
            error_msg = f"Batch LLM API call failed: {e}"
            logger.error(error_msg)
            # エラー時のログ記録
            self._log_batch_llm_request(batch_requests, user_prompt or "バッチプロンプト生成失敗", error=error_msg)
            return []
    
    def _call_llm(self, request: LLMReviewRequest, source_detection: DetectionResult = None) -> Optional[LLMReviewResponse]:
        """
        LLM APIを呼び出し
        
        Args:
            request: レビューリクエスト
            source_detection: 元の検出結果
            
        Returns:
            LLMレビュー結果
        """
        user_prompt = None
        response_text = None
        
        try:
            user_prompt = self._get_user_prompt(request)
            
            # リクエスト回数をカウント
            self.request_count += 1
            
            if self.provider == LLMProvider.ANTHROPIC:
                response_text = self._call_anthropic(user_prompt)
            elif self.provider == LLMProvider.OPENAI:
                response_text = self._call_openai(user_prompt)
            else:
                error_msg = f"Unsupported provider: {self.provider}"
                logger.error(error_msg)
                self._log_llm_request(request, user_prompt, error=error_msg)
                return None
            
            # 成功時のログ記録
            self._log_llm_request(request, user_prompt, response_text)
            
            return self._parse_llm_response(response_text, request.original, source_detection)
            
        except Exception as e:
            error_msg = f"LLM API call failed: {e}"
            logger.error(error_msg)
            # エラー時のログ記録
            self._log_llm_request(request, user_prompt or "プロンプト生成失敗", error=error_msg)
            return None
    
    def _call_anthropic(self, user_prompt: str) -> str:
        """Anthropic APIを呼び出し"""
        message = self.client.messages.create(
            model=self.model,
            max_tokens=self.max_tokens,
            temperature=self.temperature,
            system=self.system_prompt,
            messages=[
                {
                    "role": "user",
                    "content": user_prompt
                }
            ]
        )
        
        return message.content[0].text
    
    def _call_openai(self, user_prompt: str) -> str:
        """OpenAI APIを呼び出し"""
        response = self.client.chat.completions.create(
            model=self.model,
            max_tokens=self.max_tokens,
            temperature=self.temperature,
            messages=[
                {"role": "system", "content": self.system_prompt},
                {"role": "user", "content": user_prompt}
            ]
        )
        
        return response.choices[0].message.content
    
    def _parse_llm_response(self, response: str, original: str, source_detection: DetectionResult = None) -> Optional[LLMReviewResponse]:
        """
        LLM応答をパース
        
        Args:
            response: LLM応答文字列
            original: 元のテキスト
            source_detection: 元の検出結果
            
        Returns:
            パースされたレビュー結果
        """
        try:
            # JSONを抽出（```json```で囲まれている場合に対応）
            json_str = response.strip()
            
            if '```json' in json_str:
                start = json_str.find('```json') + 7
                end = json_str.find('```', start)
                json_str = json_str[start:end].strip()
            elif '```' in json_str:
                start = json_str.find('```') + 3
                end = json_str.find('```', start)
                json_str = json_str[start:end].strip()
            
            # JSONパース
            data = json.loads(json_str)
            
            # レスポンスオブジェクト作成
            llm_response = LLMReviewResponse(
                issue_type=data.get('issue_type', 'none'),
                original=data.get('original', original),
                suggested_fix=data.get('suggested_fix'),
                canonical=data.get('canonical'),
                reason=data.get('reason', ''),
                confidence=float(data.get('confidence', 0.0)),
                source_detection=source_detection  # 元の検出結果を保持
            )
            
            return llm_response
            
        except json.JSONDecodeError as e:
            logger.error(f"Failed to parse JSON response: {e}")
            logger.debug(f"Response: {response[:200]}...")
            return None
        except Exception as e:
            logger.error(f"Error parsing LLM response: {e}")
            logger.debug(f"Response text (first 200 chars): {response[:200]}...")
            return None
    
    def _extract_json_from_response(self, response_text: str) -> str:
        """
        レスポンステキストからJSONを抽出
        
        Args:
            response_text: LLMからの応答テキスト
            
        Returns:
            抽出されたJSON文字列
        """
        if "```json" in response_text:
            json_start = response_text.find("```json") + 7
            json_end = response_text.find("```", json_start)
            if json_end == -1:
                # 終了の```がない場合、テキストの最後まで
                return response_text[json_start:].strip()
            return response_text[json_start:json_end].strip()
        elif "[" in response_text and "]" in response_text:
            # 生JSONを探す
            json_start = response_text.find("[")
            json_end = response_text.rfind("]") + 1
            return response_text[json_start:json_end]
        else:
            return ""
    
    def _repair_json(self, json_text: str) -> str:
        """
        不完全なJSONを修復
        
        Args:
            json_text: JSON文字列
            
        Returns:
            修復されたJSON文字列
        """
        try:
            # 既に有効なJSONかチェック
            json.loads(json_text)
            return json_text
        except json.JSONDecodeError:
            pass
        
        # 一般的な修復処理
        repaired = json_text.strip()
        
        # 末尾の不完全なオブジェクトを削除
        if repaired.endswith(','):
            repaired = repaired.rstrip(',')
        
        # 不完全なオブジェクトを検出して削除
        lines = repaired.split('\n')
        valid_lines = []
        brace_count = 0
        bracket_count = 0
        
        for line in lines:
            line = line.strip()
            if not line:
                continue
                
            # 括弧のカウント
            brace_count += line.count('{') - line.count('}')
            bracket_count += line.count('[') - line.count(']')
            
            valid_lines.append(line)
            
            # 不正な状態の検出
            if brace_count < 0 or bracket_count < 0:
                break
        
        # 配列を正しく閉じる
        repaired = '\n'.join(valid_lines)
        if bracket_count > 0:
            repaired += ']'
        if brace_count > 0:
            repaired += '}' * brace_count
        
        return repaired
    
    def save_review_results(self, responses: List[LLMReviewResponse], output_path: str):
        """
        レビュー結果をJSONファイルに保存
        
        Args:
            responses: LLMレビュー結果のリスト
            output_path: 出力ファイルパス
        """
        try:
            data = {
                'timestamp': str(pd.Timestamp.now()),
                'provider': self.provider.value,
                'model': self.model,
                'total_reviews': len(responses),
                'reviews': [response.to_dict() for response in responses]
            }
            
            with open(output_path, 'w', encoding='utf-8') as f:
                json.dump(data, f, ensure_ascii=False, indent=2)
            
            logger.info(f"LLM review results saved to {output_path}")
            
        except Exception as e:
            logger.error(f"Failed to save LLM review results: {e}")
    
    def get_review_statistics(self, responses: List[LLMReviewResponse]) -> Dict[str, Any]:
        """
        レビュー統計を取得
        
        Args:
            responses: LLMレビュー結果のリスト
            
        Returns:
            統計情報辞書
        """
        if not responses:
            return {}
        
        issue_types = [r.issue_type for r in responses]
        confidences = [r.confidence for r in responses]
        
        stats = {
            'total_reviews': len(responses),
            'issue_type_counts': {
                'variant': issue_types.count('variant'),
                'typo': issue_types.count('typo'),
                'none': issue_types.count('none')
            },
            'avg_confidence': sum(confidences) / len(confidences),
            'high_confidence': len([c for c in confidences if c >= 0.7]),
            'provider': self.provider.value,
            'model': self.model,
            'request_count': self.request_count  # リクエスト回数を追加
        }
        
        return stats

    def _call_llm_for_cells(self, cells: List[CellData]) -> List[LLMReviewResponse]:
        """
        セルを直接LLMで処理
        
        Args:
            cells: セルデータのリスト
            
        Returns:
            LLMレビュー結果のリスト
        """
        if not cells:
            return []
        
        user_prompt = None
        response_text = None
        
        try:
            # セル用のプロンプトを生成
            user_prompt = self._get_cell_batch_prompt(cells)
            
            # リクエスト回数をカウント（バッチでも1回）
            self.request_count += 1
            
            if self.provider == LLMProvider.ANTHROPIC:
                response_text = self._call_anthropic(user_prompt)
            elif self.provider == LLMProvider.OPENAI:
                response_text = self._call_openai(user_prompt)
            else:
                error_msg = f"Unsupported provider: {self.provider}"
                logger.error(error_msg)
                self._log_cell_batch_request(cells, user_prompt, error=error_msg)
                return []
            
            # 成功時のログ記録
            self._log_cell_batch_request(cells, user_prompt, response_text)
            
            return self._parse_cell_batch_response(response_text, cells)
            
        except Exception as e:
            error_msg = f"Cell batch LLM API call failed: {e}"
            logger.error(error_msg)
            # エラー時のログ記録
            self._log_cell_batch_request(cells, user_prompt or "セルバッチプロンプト生成失敗", error=error_msg)
            return []

    def _get_batch_user_prompt(self, batch_requests: List[LLMReviewRequest]) -> str:
        """
        バッチ用のユーザープロンプトを生成
        
        Args:
            batch_requests: レビューリクエストのリスト
            
        Returns:
            バッチ用プロンプト
        """
        prompt_parts = ["以下の候補について誤字脱字のみを判定してください。表記ゆれは別途処理済みです。\n"]
        
        for i, request in enumerate(batch_requests):
            canonical_text = f"正準（あれば）: {request.canonical}" if request.canonical else "正準（あれば）: なし"
            related_text = f"関連語: {', '.join(request.related)}" if request.related else "関連語: なし"
            
            prompt_parts.append(f"""
項目{i+1}:
【文脈】{request.context}
【候補語】{request.original}
【{canonical_text}】
【{related_text}】
""")
        
        prompt_parts.append("""
各項目について以下のJSONスキーマで応答してください：
[
  {
    "item_index": 1,
    "issue_type": "typo" | "none",
    "original": "string",
    "suggested_fix": "string | null",
    "canonical": "string | null", 
    "reason": "string",
    "confidence": 0.0
  },
  ...
]

重要な制限事項：
- 明らかな誤字・誤変換（必要→比様、頻度→振動など）のみ "typo" として検出
- 表記統一（データベース→DB、ログイン→loginなど）は別途処理済みのため検出しない
- アルファベット→日本語の変更は提案しない
- 信頼度0.7以上のもののみ指摘""")
        
        return "".join(prompt_parts)

    def _parse_batch_llm_response(self, response_text: str, batch_requests: List[LLMReviewRequest], source_detections: List[DetectionResult]) -> List[LLMReviewResponse]:
        """
        バッチLLMレスポンスをパース
        
        Args:
            response_text: LLMからの応答テキスト
            batch_requests: 元のリクエストリスト
            source_detections: 元の検出結果リスト
            
        Returns:
            LLMレビュー結果のリスト（修正が必要な項目のみ）
        """
        try:
            # JSON抽出機能を使用
            json_text = self._extract_json_from_response(response_text)
            if not json_text:
                logger.error("JSON format not found in batch response")
                return []
            
            # JSON修復を試行
            json_text = self._repair_json(json_text)
            
            # JSONパース
            parsed_items = json.loads(json_text)
            
            responses = []
            for item in parsed_items:
                item_index = item.get("item_index", 1) - 1  # 0ベースに変換
                issue_type = item.get("issue_type", "none")
                confidence = float(item.get("confidence", 0.0))
                
                # 修正が必要な項目のみを処理（typo/variantで信頼度が0.7以上）
                if issue_type in ["typo", "variant"] and confidence >= 0.7 and 0 <= item_index < len(source_detections):
                    original_text = item.get("original", "")
                    suggested_fix = item.get("suggested_fix")
                    canonical = item.get("canonical")
                    
                    # 元テキストと修正案が同じ場合は除外（LLMの誤判定）
                    if original_text == suggested_fix:
                        logger.debug(f"Skipping batch identical suggestion: '{original_text}' == '{suggested_fix}'")
                        continue
                    
                    # canonicalが設定されている場合は表記ゆれ（variant）として扱う
                    final_issue_type = "variant" if canonical else issue_type
                    
                    response = LLMReviewResponse(
                        issue_type=final_issue_type,
                        original=original_text,
                        suggested_fix=suggested_fix,
                        canonical=canonical,
                        reason=item.get("reason", ""),
                        confidence=confidence,
                        source_detection=source_detections[item_index]
                    )
                    responses.append(response)
                else:
                    # 修正不要の場合はログに記録してスキップ
                    if 0 <= item_index < len(source_detections):
                        logger.debug(f"Skipping batch item {item_index}: {issue_type} (confidence: {confidence})")
            
            logger.info(f"Filtered batch LLM results: {len(responses)} items require fixes from {len(parsed_items)} reviewed")
            return responses
            
        except (json.JSONDecodeError, KeyError, ValueError, IndexError) as e:
            logger.error(f"Failed to parse batch LLM response: {e}")
            logger.debug(f"Response text (first 500 chars): {response_text[:500]}...")
            # JSON修復後のテキストもログ出力
            try:
                json_text = self._extract_json_from_response(response_text)
                if json_text:
                    repaired_json = self._repair_json(json_text)
                    logger.debug(f"Repaired JSON (first 300 chars): {repaired_json[:300]}...")
            except Exception:
                pass
            return []

    def _log_batch_llm_request(self, batch_requests: List[LLMReviewRequest], user_prompt: str, response: str = None, error: str = None):
        """
        バッチLLMリクエストをログに記録
        
        Args:
            batch_requests: レビューリクエストのリスト
            user_prompt: ユーザープロンプト
            response: LLMからの応答
            error: エラーメッセージ
        """
        try:
            if not hasattr(self, 'request_logger'):
                return
            
            # バッチの場所情報を収集
            locations = []
            for request in batch_requests:
                if hasattr(request, 'cell_data') and request.cell_data:
                    location = f"{request.cell_data.file_name}:{request.cell_data.sheet_name}:{request.cell_data.cell_address}"
                    locations.append(location)
                else:
                    locations.append("不明")
            
            log_entry = {
                "timestamp": datetime.now().isoformat(),
                "provider": self.provider.value,
                "model": self.model,
                "batch_size": len(batch_requests),
                "cell_locations": locations,
                "system_prompt": self.system_prompt,
                "user_prompt": user_prompt,
                "response": response,
                "error": error
            }
            
            self.request_logger.info(json.dumps(log_entry, ensure_ascii=False, indent=2))
            
        except Exception as e:
            logger.error(f"Failed to log batch LLM request: {e}")

    def _get_cell_batch_prompt(self, cells: List[CellData]) -> str:
        """
        セルバッチ用のプロンプトを生成
        
        Args:
            cells: セルデータのリスト
            
        Returns:
            プロンプト文字列
        """
        prompt_parts = ["以下はシステム設計書の一部です。誤字・誤変換・脱字および表記統一ルールに該当する語句を検出してください。\n"]
        
        for i, cell in enumerate(cells):
            location = f"{cell.file_name}:{cell.sheet_name}:{cell.cell_address}"
            prompt_parts.append(f"\n項目{i+1} ({location}):\n{cell.text}")
        
        prompt_parts.append("""\n
各項目について以下のJSONスキーマで応答してください：
[
  {
    "item_index": 1,
    "issue_type": "typo" | "none",
    "original": "string",
    "suggested_fix": "string | null",
    "canonical": "string | null", 
    "reason": "string",
    "confidence": 0.0
  },
  ...
]

重要な判定・修正ルール：
【思考プロセス】各語句について以下を順次確認
1. システム設計書の文脈で意味が通るか
2. 同音異義語で、より適切な語句があるか  
3. ビジネス・IT・会計分野の専門用語として正しいか
4. 文字の欠落・余剰・誤変換がないか
5. 完全で自然な修正案を提案できるか

【修正原則】
✓ 文脈に最も適した完全な語句を提案
✓ 部分削除ではなく、意味の通る完全な表現に修正
✓ システム設計書として自然で専門的な語句を選択
- 文書内容・説明文でのみ表記統一ルールを適用、技術仕様の英語表記は維持
- 問題がない場合は "none"、信頼度0.7以上のもののみ指摘""")
        
        return "".join(prompt_parts)

    def _parse_cell_batch_response(self, response_text: str, cells: List[CellData]) -> List[LLMReviewResponse]:
        """
        セルバッチLLMレスポンスをパース
        
        Args:
            response_text: LLMからの応答テキスト
            cells: 元のセルリスト
            
        Returns:
            LLMレビュー結果のリスト（修正が必要な項目のみ）
        """
        try:
            # JSON抽出機能を使用
            json_text = self._extract_json_from_response(response_text)
            if not json_text:
                logger.error("JSON format not found in cell batch response")
                return []
            
            # JSON修復を試行
            json_text = self._repair_json(json_text)
            
            # JSONパース
            parsed_items = json.loads(json_text)
            
            responses = []
            for item in parsed_items:
                item_index = item.get("item_index", 1) - 1  # 0ベースに変換
                issue_type = item.get("issue_type", "none")
                confidence = float(item.get("confidence", 0.0))
                
                # 修正が必要な項目のみを処理（typo/variantで信頼度が0.7以上）
                if (issue_type in ["typo", "variant"] and confidence >= 0.7 and 0 <= item_index < len(cells)):
                    original_text = item.get("original", "")
                    suggested_fix = item.get("suggested_fix", "")
                    canonical = item.get("canonical")
                    
                    # 元テキストと修正案が同じ場合は除外（LLMの誤判定）
                    if original_text == suggested_fix:
                        logger.debug(f"Skipping identical suggestion: '{original_text}' == '{suggested_fix}'")
                        continue
                    
                    # canonicalが設定されている場合は表記ゆれ（variant）として扱う
                    final_issue_type = "variant" if canonical else issue_type
                    final_issue_enum = IssueType.VARIANT if canonical else IssueType.TYPO
                    
                    cell = cells[item_index]
                    # CellDataから仮のDetectionResultを作成
                    fake_detection = DetectionResult(
                        cell_data=cell,
                        issue_type=final_issue_enum,
                        original=cell.text,
                        suggested_fix=suggested_fix,
                        canonical=canonical,
                        confidence=confidence,
                        reason="LLM直接レビュー",
                        context=cell.text,
                        related_terms=[]
                    )
                    
                    response = LLMReviewResponse(
                        issue_type=final_issue_type,
                        original=original_text,
                        suggested_fix=suggested_fix,
                        canonical=canonical,
                        reason=item.get("reason", ""),
                        confidence=confidence,
                        source_detection=fake_detection
                    )
                    responses.append(response)
                else:
                    # 修正不要の場合はログに記録してスキップ
                    if 0 <= item_index < len(cells):
                        logger.debug(f"Skipping item {item_index}: {issue_type} (confidence: {confidence})")
            
            logger.info(f"Filtered LLM results: {len(responses)} items require fixes from {len(parsed_items)} reviewed")
            return responses
            
        except (json.JSONDecodeError, KeyError, ValueError, IndexError) as e:
            logger.error(f"Failed to parse cell batch LLM response: {e}")
            logger.debug(f"Response text (first 500 chars): {response_text[:500]}...")
            # JSON修復後のテキストもログ出力
            try:
                json_text = self._extract_json_from_response(response_text)
                if json_text:
                    repaired_json = self._repair_json(json_text)
                    logger.debug(f"Repaired JSON (first 300 chars): {repaired_json[:300]}...")
            except Exception:
                pass
            return []

    def _log_cell_batch_request(self, cells: List[CellData], user_prompt: str, response: str = None, error: str = None):
        """
        セルバッチLLMリクエストをログに記録
        
        Args:
            cells: セルデータのリスト
            user_prompt: ユーザープロンプト
            response: LLMからの応答
            error: エラーメッセージ
        """
        try:
            # セルの場所情報を収集
            locations = []
            for cell in cells:
                location = f"{cell.file_name}:{cell.sheet_name}:{cell.cell_address}"
                locations.append(location)
            
            log_entry = {
                "timestamp": datetime.now().isoformat(),
                "provider": self.provider.value,
                "model": self.model,
                "batch_size": len(cells),
                "cell_locations": locations,
                "system_prompt": self.system_prompt,
                "user_prompt": user_prompt,
                "response": response,
                "error": error
            }
            
            # JSON形式でログ出力
            logger.bind(llm_request=True).info(json.dumps(log_entry, ensure_ascii=False, indent=2))
            
        except Exception as e:
            logger.error(f"Failed to log cell batch LLM request: {e}")


if __name__ == "__main__":
    # テスト実行
    import pandas as pd
    import yaml
    # DetectionResult と IssueType は既にこのファイルで定義済み
    from .extract import CellData
    
    # 設定読み込み
    with open("config.yml", "r", encoding="utf-8") as f:
        config = yaml.safe_load(f)
    
    # レビューワー初期化
    reviewer = LLMReviewer(config)
    
    # テストデータ
    test_cell = CellData("test.xlsx", "Sheet1", "A1", 1, 1, "感情科目", "感情科目")
    test_result = DetectionResult(
        cell_data=test_cell,
        issue_type=IssueType.TYPO,
        original="感情科目",
        suggested_fix="勘定科目",
        canonical="勘定科目",
        confidence=0.75,
        reason="類似度75%で勘定科目に近い",
        context="感情科目の管理について",
        related_terms=["仕訳", "貸借", "元帳", "会計"]
    )
    
    # LLMレビュー実行（APIキーが設定されている場合）
    if reviewer.client:
        responses = reviewer.review_detection_results([test_result])
        
        if responses:
            print("=== LLMレビュー結果 ===")
            for response in responses:
                print(f"判定: {response.issue_type}")
                print(f"元: {response.original}")
                print(f"修正案: {response.suggested_fix}")
                print(f"理由: {response.reason}")
                print(f"信頼度: {response.confidence}")
    else:
        print("LLMクライアントが初期化されていません（APIキーを確認してください）")