import pandas as pd
import os

# 修正版のテストケース仕様書（表記ゆれを統一）
test_cases = {
    'テストケースID': ['TC001', 'TC002', 'TC003', 'TC004', 'TC005'],
    '機能名': ['ログイン', 'ログイン', 'パスワード変更', 'ユーザ管理', 'セッション管理'],  # 表記ゆれ: ユーザ
    'テスト項目': [
        '正しいユーザーIDとパスワードでログインできること',
        '間違ったパスワードでログイン失敗すること',
        'パスワードポリシーに適合すること',
        '管理者がユーザを追加できること',  # 表記ゆれ: ユーザ
        'セッションタイムアウトが正しく動作すること'
    ],
    '期待結果': [
        'ログイン成功しダッシュボードへ遷移',
        'エラーメッセージ表示',
        'パスワードが正常に更新される',
        'ユーザが追加される',  # 表記ゆれ: ユーザ
        'ログイン画面へリダイレクト'
    ],
    '前提条件': [
        'ユーザーがシステムに未ログイン',
        'ユーザーがシステムに未ログイン',
        '一般ユーザーでログイン済み',
        '管理者権限でログイン済み',
        'ユーザがログイン済み'  # 表記ゆれ: ユーザ
    ],
    'テストデータ': [
        'ID:admin001, パスワード:Admin@123',
        'ID:admin001, パスワード:wrongpass',
        '新パスワード:NewPass@456',
        'ユーザ名:testuser001, パスワード:TestPass@123',  # 表記ゆれ: ユーザ
        'セッションタイムアウト:30分'
    ],
    '実行日': ['2025/01/15', '2025/01/15', '2025/01/16', '2025/01/16', '2025/01/17'],
    '結果': ['OK', 'OK', 'OK', 'NG', '未実施'],
    '備考': ['', '', '', 'メール送信エラー', '']
}

# 在庫管理API詳細設計書（表記ゆれを統一、誤字を1つ追加）
api_design = {
    'API名': ['商品一覧取得', '商品詳細取得', '在庫更新', '商品登録', '商品削除'],
    'メソッド': ['GET', 'GET', 'PUT', 'POST', 'DELETE'],
    'エンドポイント': [
        '/api/products',
        '/api/products/{id}',
        '/api/products/{id}/stock',
        '/api/products',
        '/api/products/{id}'
    ],
    'パラメータ': [
        'page, limit, category',
        'id',
        'id, quantity',
        'name, price, category',
        'id'
    ],
    'レスポンス': [
        '商品リスト',
        '商品詳細情報',
        '更新結果',
        '登録結果',
        '削除結果'
    ],
    '権限': ['全ユーザ', '全ユーザ', '管理者', '管理者', '管理者'],  # 表記ゆれ: ユーザ
    'エラーコード': ['400, 500', '404, 500', '400, 403, 404', '400, 403', '403, 404'],
    '備考': ['ページネーション対応', '', '在庫数の検証あり', 'バリデーション実装', '論理削除を使用']  # "論理削除"は正しい
}

# 顧客管理システム設計書（表記ゆれを統一、誤字を1つ追加）
customer_system = {
    '画面名': ['ログイン画面', 'ダッシュボード', '顧客一覧', '顧客詳細', '顧客登録'],
    '機能': [
        'ユーザ認証',  # 表記ゆれ: ユーザ
        '統計情報表示',
        '顧客データベース検索',  # "顧客データーベース"を"顧客データベース"に修正
        '顧客情報閲覧',
        '新規顧客登録'
    ],
    'アクセス権限': ['全ユーザ', '一般ユーザ以上', '一般ユーザ以上', '一般ユーザ以上', '管理者のみ'],  # 表記ゆれ: ユーザ
    '主要項目': [
        'ユーザID, 認証情報',  # 表記ゆれ: ユーザ
        '売上集計, 顧客数, 感情科目',  # 誤字: 勘定科目 → 感情科目
        '顧客名, 電話番号, 連絡先',  # "メールアドレス"を避ける
        '基本情報, 取引履歴',
        '顧客情報, 担当者'
    ],
    '備考': [
        '二要素認証対応',
        'リアルタイム更新',
        'ページネーション対応',
        '履歴は過去1年分',
        '入力検証必須'
    ]
}

# DB設計（1つだけ意図的な誤字を追加）
db_design = {
    'テーブル名': ['users', 'products', 'orders', 'customers'],
    'カラム': [
        'user_id, username, pass_hash, email_addr',  # 英語表記で統一、カタカナとの混在を避ける
        'product_id, name, price, stock_count',
        'order_id, customer_id, total_amount, order_status',
        'customer_id, name, contact_info, address'  # 誤字を削除、感情科目の方を重視
    ],
    '主キー': ['user_id', 'product_id', 'order_id', 'customer_id'],
    'インデックス': ['username', 'name', 'customer_id', 'name'],
    '備考': ['認証情報はハッシュ化', '在庫数は0以上', 'ステータスは列挙型', '住所は正規化']  # "パスワード"を避ける
}

# ファイル作成
os.makedirs('C:\\develop\\claude-code\\excel-checker\\data\\input', exist_ok=True)

# テストケース仕様書
with pd.ExcelWriter('C:\\develop\\claude-code\\excel-checker\\data\\input\\テストケース仕様書.xlsx', engine='openpyxl') as writer:
    df = pd.DataFrame(test_cases)
    df.to_excel(writer, sheet_name='機能テスト', index=False)

# 在庫管理API詳細設計書
with pd.ExcelWriter('C:\\develop\\claude-code\\excel-checker\\data\\input\\在庫管理API詳細設計書.xlsx', engine='openpyxl') as writer:
    df = pd.DataFrame(api_design)
    df.to_excel(writer, sheet_name='API仕様', index=False)

# 顧客管理システム設計書
with pd.ExcelWriter('C:\\develop\\claude-code\\excel-checker\\data\\input\\顧客管理システム設計書.xlsx', engine='openpyxl') as writer:
    df1 = pd.DataFrame(customer_system)
    df1.to_excel(writer, sheet_name='画面一覧', index=False)
    
    df2 = pd.DataFrame(db_design)
    df2.to_excel(writer, sheet_name='DB設計', index=False)

print("修正完了: data/inputフォルダのExcelファイルを更新しました")