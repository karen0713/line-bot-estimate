from flask import Flask, request, abort, redirect, url_for
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import (
    Configuration, ApiClient, MessagingApi, ReplyMessageRequest, TextMessage, FlexMessage, FlexContainer
)
from linebot.v3.webhooks import MessageEvent, TextMessageContent, PostbackEvent
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re
import os
import json
import logging
from user_management import UserManager
from stripe_payment import StripePayment
from excel_online import ExcelOnlineManager
import sqlite3

# ログ設定
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.FileHandler('app.log'),
        logging.StreamHandler()  # コンソールにも出力
    ]
)
logger = logging.getLogger(__name__)

app = Flask(__name__)

# 環境変数から設定を取得
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN', 'Khehk/dQp536fyLT0u0UVSzBWh7zCNYDGPODIi5KtpNmkp1QJXc5kDKVlTaavNYW/12lK/HLF001axW4WLfoOXqLxTNMaXb6E6BnqtrAIxyoRP56Nw0g41L6JT2An3cA86Nl6tHqUY8ul5gP+9L8BgdB04t89/1O/w1cDnyilFU=')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET', '8326aecb26b4e9c41ef8d702b73c6617')

# Google Sheets設定
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
# 共有スプレッドシートの設定
SHARED_SPREADSHEET_ID = os.environ.get('SHARED_SPREADSHEET_ID', '1GkJ8OYwIIMnYqxcwVBNArvk2byFL3UlGHgkyTiV6QU0')
DEFAULT_SHEET_NAME = os.environ.get('DEFAULT_SHEET_NAME', '比較見積書 ロング')

# 従来の設定（後方互換性のため保持）
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID', SHARED_SPREADSHEET_ID)
SHEET_NAME = os.environ.get('SHEET_NAME', DEFAULT_SHEET_NAME)

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

# ユーザー管理システムの初期化
try:
    user_manager = UserManager()
    logger.info("User management system initialized successfully")
except Exception as e:
    logger.error(f"User management system initialization error: {e}")
    user_manager = None

# Stripe決済システムの初期化
try:
    stripe_payment = StripePayment()
    logger.info("Stripe payment system initialized successfully")
except Exception as e:
    logger.error(f"Stripe payment system initialization error: {e}")
    stripe_payment = None

# Excel Onlineシステムの初期化
try:
    # 環境変数のデバッグログ
    logger.info(f"MS_CLIENT_ID: {os.environ.get('MS_CLIENT_ID', 'NOT_SET')}")
    logger.info(f"MS_CLIENT_SECRET: {os.environ.get('MS_CLIENT_SECRET', 'NOT_SET')[:10]}..." if os.environ.get('MS_CLIENT_SECRET') else 'NOT_SET')
    logger.info(f"MS_TENANT_ID: {os.environ.get('MS_TENANT_ID', 'NOT_SET')}")
    
    excel_online_manager = ExcelOnlineManager()
    logger.info("Excel Online system initialized successfully")
except Exception as e:
    logger.error(f"Excel Online system initialization error: {e}")
    excel_online_manager = None

# ユーザーセッション管理（簡易版）
user_sessions = {}

# 商品テンプレート
PRODUCT_TEMPLATES = {
    "Tシャツ": {"sizes": ["S", "M", "L", "XL"], "prices": [1500, 1500, 1500, 1500]},
    "ポロシャツ": {"sizes": ["S", "M", "L", "XL"], "prices": [2500, 2500, 2500, 2500]},
    "作業服": {"sizes": ["S", "M", "L", "XL"], "prices": [3000, 3000, 3000, 3000]},
    "帽子": {"sizes": ["FREE", "L"], "prices": [800, 800]},
    "タオル": {"sizes": ["FREE"], "prices": [500]},
    "その他": {"sizes": ["FREE"], "prices": [1000]}
}

# --- SHEET_WRITE_CONFIGを4シート名ごとに分岐 ---
SHEET_WRITE_CONFIG = {
    "比較見積書 ロング": {
        "company": "A2:H3",
        "date": "M2:Q2",
        "product": {
            "現状": {"name": ["A", "B"], "price": "C", "quantity": "D", "cycle": "G", "row_start": 19, "row_end": 36},
            "当社": {"name": ["I", "J"], "price": "K", "quantity": "L", "cycle": "O", "row_start": 19, "row_end": 36}
        }
    },
    "比較御見積書　ショート": {
        "company": "A2:H3",
        "date": "M2:Q2",
        "product": {
            "現状": {"name": ["A", "B"], "price": "C", "quantity": "D", "cycle": "G", "row_start": 19, "row_end": 28},
            "当社": {"name": ["I", "J"], "price": "K", "quantity": "L", "cycle": "O", "row_start": 19, "row_end": 28}
        }
    },
    "新規見積書　ショート": {
        "company": "B5:G7",
        "date": "I2:J3",
        "product": {
            "default": {"name": ["B", "C", "D"], "cycle": "E", "quantity": "F", "price": "G", "row_start": 24, "row_end": 30}
        }
    },
    "新規見積書　ロング": {
        "company": "B5:G7",
        "date": "I2:J3",
        "product": {
            "default": {"name": ["B", "C"], "place": "D", "cycle": "E", "quantity": "F", "price": "G", "row_start": 27, "row_end": 48}
        }
    }
}

def setup_google_sheets():
    """Google Sheets APIの設定"""
    try:
        # 環境変数からサービスアカウント情報を取得
        service_account_info = os.environ.get('GOOGLE_SHEETS_CREDENTIALS')
        if service_account_info:
            creds = Credentials.from_service_account_info(
                json.loads(service_account_info), scopes=SCOPES)
        else:
            # ローカル開発用（ファイルから読み込み）
            creds = Credentials.from_service_account_file(
                'gsheet_service_account.json', scopes=SCOPES)
        client = gspread.authorize(creds)
        return client
    except Exception as e:
        print(f"Google Sheets setup error: {e}")
        return None

def parse_estimate_data(text):
    """1行ずつ項目名:値を抽出し、柔軟に辞書化"""
    data = {}
    lines = text.replace('\r', '').split('\n')
    for line in lines:
        line = line.strip()
        if not line:
            continue
        # コロンで分割（全角・半角両対応）
        if ':' in line:
            key, value = line.split(':', 1)
        elif '：' in line:
            key, value = line.split('：', 1)
        else:
            continue
        key = key.strip()
        value = value.strip()
        # 抽出対象の項目を拡張
        if key in ['社名', '会社名', '商品名', '単価', '数量', '日付', 'サイクル', '設置場所']:
            data[key] = value
    # 社名と会社名を統一
    if '会社名' in data and '社名' not in data:
        data['社名'] = data['会社名']
    # 料金を計算
    if '単価' in data and '数量' in data:
        try:
            unit_price = int(re.sub(r'[^0-9]', '', data['単価']))
            quantity = int(re.sub(r'[^0-9]', '', data['数量']))
            data['料金'] = unit_price * quantity
        except ValueError:
            data['料金'] = 0
    print(f"parse_estimate_data: {data}")
    return data

def extract_spreadsheet_id(url):
    """GoogleスプレッドシートURLからIDを抽出"""
    import re
    pattern = r'/spreadsheets/d/([a-zA-Z0-9-_]+)'
    match = re.search(pattern, url)
    return match.group(1) if match else None

def write_to_spreadsheet(data, user_id=None):
    """スプレッドシートまたはExcel Onlineにデータを書き込み（シート名・項目別対応）"""
    try:
        print(f"開始: データ書き込み処理")
        
        # まずExcel Online設定をチェック
        excel_online_enabled = False
        excel_url = None
        excel_file_id = None
        excel_sheet_name = None
        
        if user_id and user_manager and excel_online_manager:
            excel_url, excel_file_id, excel_sheet_name = user_manager.get_user_excel_online(user_id)
            if excel_url and excel_file_id:
                excel_online_enabled = True
                print(f"Excel Online設定を検出: {excel_url}")
        
        # Excel Onlineが有効な場合はExcel Onlineに書き込み
        if excel_online_enabled:
            return write_to_excel_online(data, excel_file_id, excel_sheet_name, user_id)
        
        # 従来のGoogle Sheets処理
        return write_to_google_sheets(data, user_id)
        
    except Exception as e:
        print(f"データ書き込みエラー: {e}")
        return False, f"データ書き込みエラー: {e}"

def write_to_excel_online(data, file_id, sheet_name, user_id=None):
    """Excel Onlineにデータを書き込み"""
    try:
        print(f"開始: Excel Online書き込み処理")
        print(f"file_id: {file_id}, sheet_name: {sheet_name}")
        
        if not excel_online_manager:
            return False, "Excel Onlineシステムが利用できません"
        
        # 商品データの書き込み
        if '商品名' in data and '単価' in data and '数量' in data:
            # 空いている行を探す
            row_number = 19  # デフォルトの開始行
            
            # 既存データを確認して空いている行を探す
            existing_data, error = excel_online_manager.read_range(file_id, sheet_name, 'A19:G36')
            if existing_data:
                for i, row in enumerate(existing_data):
                    if not any(cell for cell in row[:3] if cell):  # 最初の3列が空の場合
                        row_number = 19 + i
                        break
                else:
                    row_number = 19 + len(existing_data)  # 最後の行の次の行
            
            # 商品データを書き込み
            success, error = excel_online_manager.write_product_data_excel(data, file_id, sheet_name, row_number)
            if not success:
                return False, f"商品データの書き込みに失敗: {error}"
            
            print(f"商品データを行 {row_number} に書き込みました")
            
        # 会社情報の更新
        if '社名' in data or '日付' in data:
            success, error = excel_online_manager.update_company_info_excel(data, file_id, sheet_name)
            if not success:
                return False, f"会社情報の更新に失敗: {error}"
            
            print("会社情報を更新しました")
        
        return True, "Excel Onlineにデータを書き込みました"
        
    except Exception as e:
        print(f"Excel Online書き込みエラー: {e}")
        return False, f"Excel Online書き込みエラー: {e}"

def write_to_google_sheets(data, user_id=None):
    """Google Sheetsにデータを書き込み（従来の処理）"""
    try:
        print(f"開始: Google Sheets書き込み処理")
        
        # 顧客のスプレッドシートIDを取得
        if user_id and user_manager:
            spreadsheet_id, sheet_name = user_manager.get_user_spreadsheet(user_id)
            if not spreadsheet_id:
                # ユーザーがスプレッドシートを登録していない場合は共有スプレッドシートを使用
                spreadsheet_id = SHARED_SPREADSHEET_ID
                sheet_name = DEFAULT_SHEET_NAME
                print(f"ユーザーがスプレッドシートを登録していないため、共有スプレッドシートを使用: {spreadsheet_id}")
        else:
            spreadsheet_id = SHARED_SPREADSHEET_ID
            sheet_name = DEFAULT_SHEET_NAME
        
        # --- シート名を正規化 ---
        # normalize_sheet_nameを削除
        # sheet_name = normalize_sheet_name(sheet_name)
        
        client = setup_google_sheets()
        if not client:
            print("エラー: Google Sheets接続失敗")
            return False, "Google Sheets接続エラー"
        
        print(f"成功: Google Sheets接続")
        sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
        print(f"成功: シート '{sheet_name}' を開きました")
        
        # シート名に対応する設定を取得
        sheet_config = SHEET_WRITE_CONFIG.get(sheet_name)
        if not sheet_config:
            print(f"警告: シート '{sheet_name}' の設定が見つかりません。デフォルト設定を使用します。")
            # デフォルト設定（比較見積書 ロング）
            sheet_config = SHEET_WRITE_CONFIG["比較御見積書　ショート"]
        
        print(f"SHEET_WRITE_CONFIG.keys(): {list(SHEET_WRITE_CONFIG.keys())}")
        print(f"sheet_name: '{sheet_name}'")

        # 商品名から「現状」「当社」などの語尾を除去し、商品タイプを判定
        product_name = data.get('商品名', '')
        product_type = "default"  # デフォルト
        if product_name:
            import re
            m = re.match(r"^(.*?)[\s　]*(現状|当社)$", product_name)
            if m:
                product_type = m.group(2)
                data['商品名'] = m.group(1)
            # elseはdefaultのまま

        # 商品設定を取得
        product_config = sheet_config.get('product', {}).get(product_type)
        if not product_config:
            # デフォルト設定を使用
            available_configs = list(sheet_config.get('product', {}).values())
            if available_configs:
                product_config = available_configs[0]
            else:
                print(f"エラー: シート '{sheet_name}' に商品設定が見つかりません")
                return False, f"シート '{sheet_name}' の設定エラー"
        
        print(f"商品タイプ: {product_type}")
        print(f"商品設定: {product_config}")
        print(f"利用可能な設定: {list(sheet_config.get('product', {}).keys())}")
        
        # 既存データの行数を確認
        existing_data = sheet.get_all_values()
        print(f"既存データ行数: {len(existing_data)}")
        
        # 使用済み行数を確認（商品タイプに応じた列のみ）
        row_start = product_config.get('row_start', 19)
        row_end = product_config.get('row_end', 36)
        used_rows = 0
        
        # 商品タイプに応じた列のみをチェック
        check_columns = []
        for col_key in ['name', 'option', 'price', 'quantity', 'cycle', 'place']:
            if col_key in product_config:
                col_value = product_config[col_key]
                if isinstance(col_value, list):
                    check_columns.extend(col_value)
                else:
                    check_columns.append(col_value)
        
        print(f"チェック対象列: {check_columns}")
        
        for row in range(row_start - 1, min(row_end, len(existing_data))):
            # 該当する列にデータがあるかチェック
            has_data = False
            for col_letter in check_columns:
                col_index = ord(col_letter) - ord('A')
                if col_index < len(existing_data[row]) and existing_data[row][col_index]:
                    has_data = True
                    break
            if has_data:
                used_rows += 1
        
        print(f"使用済み行数: {used_rows} (行範囲: {row_start}-{row_end})")
        print(f"チェック対象列: {check_columns}")
        
        # 次の書き込み行を決定
        next_row = row_start + used_rows
        if next_row > row_end:
            print(f"警告: 行数上限 {row_end} を超えています。{row_end}行目に書き込みます。")
            next_row = row_end
        
        print(f"書き込み行: {next_row}")
        
        # 商品名（複数列対応）
        if data.get('商品名', '') and 'name' in product_config:
            name_cols = product_config['name']
            if isinstance(name_cols, list):
                for col in name_cols:
                    sheet.update(values=[[data.get('商品名', '')]], range_name=f"{col}{next_row}")
                    print(f"{col}{next_row} に {data.get('商品名', '')} を書き込みます")
        else:
                sheet.update(values=[[data.get('商品名', '')]], range_name=f"{name_cols}{next_row}")
                print(f"{name_cols}{next_row} に {data.get('商品名', '')} を書き込みます")

        # サイクル（サイクル列が指定されている場合のみ）
        if data.get('サイクル', '') and 'cycle' in product_config:
            cycle_col = product_config['cycle']
            sheet.update(values=[[data.get('サイクル', '')]], range_name=f"{cycle_col}{next_row}")
            print(f"{cycle_col}{next_row} に {data.get('サイクル', '')} を書き込みます")
        
        # 数量
        if data.get('数量', '') and 'quantity' in product_config:
            quantity_col = product_config['quantity']
            sheet.update(values=[[data.get('数量', '')]], range_name=f"{quantity_col}{next_row}")
            print(f"{quantity_col}{next_row} に {data.get('数量', '')} を書き込みます")

        # 単価
        if data.get('単価', '') and 'price' in product_config:
            price_col = product_config['price']
            sheet.update(values=[[data.get('単価', '')]], range_name=f"{price_col}{next_row}")
            print(f"{price_col}{next_row} に {data.get('単価', '')} を書き込みます")
        
        # 設置場所（設置場所列が指定されている場合のみ）
        if data.get('設置場所', '') and 'place' in product_config:
            place_col = product_config['place']
            sheet.update(values=[[data.get('設置場所', '')]], range_name=f"{place_col}{next_row}")
            print(f"{place_col}{next_row} に {data.get('設置場所', '')} を書き込みます")
        
        print(f"成功: データを{next_row}行目に書き込みました")
        return True, f"データを{next_row}行目に正常に書き込みました"
        
    except Exception as e:
        print(f"Spreadsheet write error: {e}")
        return False, f"書き込みエラー: {str(e)}"

def update_company_info(data, user_id=None):
    """会社名と日付を更新（シート名別対応）"""
    try:
        print(f"開始: 会社情報更新処理")
        
        # まずExcel Online設定をチェック
        excel_online_enabled = False
        excel_url = None
        excel_file_id = None
        excel_sheet_name = None
        
        if user_id and user_manager and excel_online_manager:
            excel_url, excel_file_id, excel_sheet_name = user_manager.get_user_excel_online(user_id)
            if excel_url and excel_file_id:
                excel_online_enabled = True
                print(f"Excel Online設定を検出: {excel_url}")
        
        # Excel Onlineが有効な場合はExcel Onlineに更新
        if excel_online_enabled:
            return update_company_info_excel_online(data, excel_file_id, excel_sheet_name, user_id)
        
        # 従来のGoogle Sheets処理
        return update_company_info_google_sheets(data, user_id)
        
    except Exception as e:
        print(f"会社情報更新エラー: {e}")
        return False, f"会社情報更新エラー: {e}"

def update_company_info_excel_online(data, file_id, sheet_name, user_id=None):
    """Excel Onlineの会社情報を更新"""
    try:
        print(f"開始: Excel Online会社情報更新処理")
        
        if not excel_online_manager:
            return False, "Excel Onlineシステムが利用できません"
        
        success, error = excel_online_manager.update_company_info_excel(data, file_id, sheet_name)
        if not success:
            return False, f"会社情報の更新に失敗: {error}"
        
        print("Excel Onlineの会社情報を更新しました")
        return True, "Excel Onlineの会社情報を更新しました"
        
    except Exception as e:
        print(f"Excel Online会社情報更新エラー: {e}")
        return False, f"Excel Online会社情報更新エラー: {e}"

def update_company_info_google_sheets(data, user_id=None):
    """Google Sheetsの会社情報を更新（従来の処理）"""
    try:
        print(f"開始: Google Sheets会社情報更新処理")
        
        # 顧客のスプレッドシートIDを取得
        if user_id and user_manager:
            spreadsheet_id, sheet_name = user_manager.get_user_spreadsheet(user_id)
            if not spreadsheet_id:
                # ユーザーがスプレッドシートを登録していない場合は共有スプレッドシートを使用
                spreadsheet_id = SHARED_SPREADSHEET_ID
                sheet_name = DEFAULT_SHEET_NAME
                print(f"ユーザーがスプレッドシートを登録していないため、共有スプレッドシートを使用: {spreadsheet_id}")
        else:
            spreadsheet_id = SHARED_SPREADSHEET_ID
            sheet_name = DEFAULT_SHEET_NAME
        
        # --- シート名を正規化 ---
        # normalize_sheet_nameを削除
        # sheet_name = normalize_sheet_name(sheet_name)
        
        client = setup_google_sheets()
        if not client:
            print("エラー: Google Sheets接続失敗")
            return False, "Google Sheets接続エラー"
        
        sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
        
        # シート名に対応する設定を取得
        sheet_config = SHEET_WRITE_CONFIG.get(sheet_name)
        if not sheet_config:
            print(f"警告: シート '{sheet_name}' の設定が見つかりません。デフォルト設定を使用します。")
            sheet_config = SHEET_WRITE_CONFIG["比較御見積書　ショート"]
        
        updates = []
        
        # 会社名を更新
        if '社名' in data:
            company_range = sheet_config.get('company', 'A2:H3')
            # 範囲から列数を計算
            import re
            range_match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', company_range)
            if range_match:
                start_col = range_match.group(1)
                end_col = range_match.group(3)
                # 列数を計算（A=1, B=2, ...）
                start_col_num = sum((ord(c) - ord('A') + 1) * (26 ** i) for i, c in enumerate(reversed(start_col)))
                end_col_num = sum((ord(c) - ord('A') + 1) * (26 ** i) for i, c in enumerate(reversed(end_col)))
                col_count = end_col_num - start_col_num + 1
            else:
                col_count = 8  # デフォルト
            
            # 会社名の書き込み形式を決定
            company_values = [
                [data['社名']] + [''] * (col_count - 1),
                [''] * col_count
            ]
            sheet.update(values=company_values, range_name=company_range)
            updates.append(f"会社名: {data['社名']}")
            print(f"会社名を更新: {data['社名']} (範囲: {company_range}, 列数: {col_count})")
        
        # 日付を更新
        if '日付' in data:
            date_range = sheet_config.get('date', 'M2:Q2')
            # 範囲から列数を計算
            range_match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', date_range)
            if range_match:
                start_col = range_match.group(1)
                end_col = range_match.group(3)
                # 列数を計算（A=1, B=2, ...）
                def col_to_num(col):
                    result = 0
                    for i, c in enumerate(reversed(col)):
                        result += (ord(c) - ord('A') + 1) * (26 ** i)
                    return result
                
                start_col_num = col_to_num(start_col)
                end_col_num = col_to_num(end_col)
                col_count = end_col_num - start_col_num + 1
                print(f"日付範囲計算: {start_col}({start_col_num}) から {end_col}({end_col_num}) = {col_count}列")
            else:
                col_count = 5  # デフォルト
                print(f"日付範囲の正規表現マッチ失敗: {date_range}, デフォルト列数: {col_count}")
            
            # 日付の書き込み形式を決定
            date_values = [
                [data['日付']] + [''] * (col_count - 1)
            ]
            sheet.update(values=date_values, range_name=date_range)
            updates.append(f"日付: {data['日付']}")
            print(f"日付を更新: {data['日付']} (範囲: {date_range}, 列数: {col_count})")
        
        if updates:
            return True, f"更新完了: {', '.join(updates)}"
        else:
            return False, "更新するデータがありません"
        
    except Exception as e:
        print(f"Company info update error: {e}")
        return False, f"更新エラー: {str(e)}"

def create_main_menu():
    """メインメニューのFlex Messageを作成"""
    return {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "text",
                    "text": "見積書作成システム",
                    "weight": "bold",
                    "size": "lg",
                    "align": "center"
                },
                {
                    "type": "text",
                    "text": "何をしますか？",
                    "margin": "md",
                    "align": "center",
                    "color": "#666666"
                }
            ]
        },
        "footer": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "商品を追加",
                        "data": "action=add_product"
                    },
                    "style": "primary",
                    "margin": "sm"
                },
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "会社情報を更新",
                        "data": "action=update_company"
                    },
                    "style": "secondary",
                    "margin": "sm"
                },
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "利用状況確認",
                        "data": "action=check_usage"
                    },
                    "style": "secondary",
                    "margin": "sm"
                },
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "プランアップグレード",
                        "data": "action=upgrade_plan"
                    },
                    "style": "primary",
                    "margin": "sm"
                },
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "見積書を確認",
                        "data": "action=view_estimate"
                    },
                    "style": "secondary",
                    "margin": "sm"
                }
            ]
        }
    }

def create_product_selection():
    """商品選択のFlex Messageを作成"""
    return {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "text",
                    "text": "商品を追加",
                    "weight": "bold",
                    "size": "lg",
                    "align": "center"
                },
                {
                    "type": "text",
                    "text": "カスタム商品を入力してください",
                    "margin": "md",
                    "align": "center",
                    "color": "#666666"
                }
            ]
        },
        "footer": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "カスタム商品を入力",
                        "data": "action=custom_product"
                    },
                    "style": "primary",
                    "margin": "sm"
                }
            ]
        }
    }

def create_size_selection(product):
    """サイズ選択のFlex Messageを作成"""
    sizes = PRODUCT_TEMPLATES[product]["sizes"]
    buttons = []
    
    for i, size in enumerate(sizes):
        price = PRODUCT_TEMPLATES[product]["prices"][i]
        buttons.append({
            "type": "button",
            "action": {
                "type": "postback",
                "label": f"{size} ({price}円)",
                "data": f"action=select_size&product={product}&size={size}&price={price}"
            },
            "style": "secondary",
            "margin": "sm"
        })
    
    # カスタム価格入力ボタンを追加
    buttons.append({
        "type": "button",
        "action": {
            "type": "postback",
            "label": "カスタム価格を入力",
            "data": f"action=custom_price&product={product}"
        },
        "style": "primary",
        "margin": "sm"
    })
    
    return {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "text",
                    "text": f"{product}のサイズを選択",
                    "weight": "bold",
                    "size": "lg",
                    "align": "center"
                }
            ]
        },
        "footer": {
            "type": "box",
            "layout": "vertical",
            "contents": buttons
        }
    }

def create_quantity_selection(product, size, price):
    """数量選択のFlex Messageを作成"""
    buttons = []
    quantities = [1, 2, 3, 5, 10, 20, 50, 100]
    
    for qty in quantities:
        total = int(price) * qty
        buttons.append({
            "type": "button",
            "action": {
                "type": "postback",
                "label": f"{qty}個 ({total}円)",
                "data": f"action=select_quantity&product={product}&size={size}&price={price}&quantity={qty}"
            },
            "style": "secondary",
            "margin": "sm"
        })
    
    return {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "text",
                    "text": f"{product} {size} ({price}円)",
                    "weight": "bold",
                    "size": "lg",
                    "align": "center"
                },
                {
                    "type": "text",
                    "text": "数量を選択してください",
                    "margin": "md",
                    "align": "center",
                    "color": "#666666"
                }
            ]
        },
        "footer": {
            "type": "box",
            "layout": "vertical",
            "contents": buttons
        }
    }

def create_plan_selection():
    """プラン選択のFlex Messageを作成"""
    return {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "text",
                    "text": "プランアップグレード",
                    "weight": "bold",
                    "size": "lg",
                    "align": "center"
                },
                {
                    "type": "text",
                    "text": "プランを選択してください",
                    "margin": "md",
                    "align": "center",
                    "color": "#666666"
                }
            ]
        },
        "footer": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "ベーシックプラン (月額500円)",
                        "data": "action=select_plan&plan=basic"
                    },
                    "style": "primary",
                    "margin": "sm"
                },
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "プロプラン (月額1,000円)",
                        "data": "action=select_plan&plan=pro"
                    },
                    "style": "primary",
                    "margin": "sm"
                }
            ]
        }
    }

def create_sheet_selection():
    """シート選択のFlex Messageを作成"""
    return {
        "type": "bubble",
        "body": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "text",
                    "text": "シートを選択",
                    "weight": "bold",
                    "size": "lg",
                    "align": "center"
                },
                {
                    "type": "text",
                    "text": "使用するシートを選択してください",
                    "margin": "md",
                    "align": "center",
                    "color": "#666666"
                }
            ]
        },
        "footer": {
            "type": "box",
            "layout": "vertical",
            "contents": [
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "比較見積書 ロング",
                        "data": "action=select_sheet&sheet=比較見積書 ロング"
                    },
                    "style": "primary",
                    "margin": "sm"
                },
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "比較御見積書　ショート",
                        "data": "action=select_sheet&sheet=比較御見積書　ショート"
                    },
                    "style": "secondary",
                    "margin": "sm"
                },
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "新規見積書　ショート",
                        "data": "action=select_sheet&sheet=新規見積書　ショート"
                    },
                    "style": "secondary",
                    "margin": "sm"
                },
                {
                    "type": "button",
                    "action": {
                        "type": "postback",
                        "label": "新規見積書　ロング",
                        "data": "action=select_sheet&sheet=新規見積書　ロング"
                    },
                    "style": "secondary",
                    "margin": "sm"
                }
            ]
        }
    }

def create_rich_menu():
    """リッチメニューを作成"""
    try:
        with ApiClient(configuration) as api_client:
            messaging_api = MessagingApi(api_client)
            
            # 既存のリッチメニューを削除
            rich_menus = messaging_api.get_rich_menu_list()
            deleted_count = 0
            for rich_menu in rich_menus.richmenus:
                messaging_api.delete_rich_menu(rich_menu.rich_menu_id)
                deleted_count += 1
                logger.info(f"Deleted rich menu: {rich_menu.rich_menu_id}")
            
            rich_menu_dict = {
                "size": {"width": 1200, "height": 405},
                "selected": False,
                "name": "見積書作成メニュー",
                "chatBarText": "メニュー",
                "areas": [
                    {
                        "bounds": {"x": 0, "y": 0, "width": 400, "height": 405},
                        "action": {"type": "message", "label": "商品を追加", "text": "商品を追加"}
                    },
                    {
                        "bounds": {"x": 400, "y": 0, "width": 400, "height": 405},
                        "action": {"type": "message", "label": "スプレッドシート登録", "text": "スプレッドシート登録"}
                    },
                    {
                        "bounds": {"x": 800, "y": 0, "width": 400, "height": 405},
                        "action": {"type": "postback", "label": "シート選択", "data": "action=show_sheet_selection"}
                    }
                ]
            }
            rich_menu_id = messaging_api.create_rich_menu(rich_menu_dict).rich_menu_id
            messaging_api.set_default_rich_menu(rich_menu_id)
            print(f"Rich menu created and set as default: {rich_menu_id}")
            return rich_menu_id
    except Exception as e:
        print(f"Rich menu creation error: {e}")
        return None

@app.route("/", methods=['GET'])
def index():
    return "LINE Bot is running!"

@app.route("/env-check", methods=['GET'])
def env_check():
    """環境変数の確認用エンドポイント（デバッグ用）"""
    env_vars = {
        'MS_CLIENT_ID': os.environ.get('MS_CLIENT_ID', 'NOT_SET'),
        'MS_CLIENT_SECRET': 'SET' if os.environ.get('MS_CLIENT_SECRET') else 'NOT_SET',
        'MS_TENANT_ID': os.environ.get('MS_TENANT_ID', 'NOT_SET'),
        'LINE_CHANNEL_ACCESS_TOKEN': 'SET' if os.environ.get('LINE_CHANNEL_ACCESS_TOKEN') else 'NOT_SET',
        'LINE_CHANNEL_SECRET': 'SET' if os.environ.get('LINE_CHANNEL_SECRET') else 'NOT_SET',
        'SHARED_SPREADSHEET_ID': os.environ.get('SHARED_SPREADSHEET_ID', 'NOT_SET'),
        'DEFAULT_SHEET_NAME': os.environ.get('DEFAULT_SHEET_NAME', 'NOT_SET'),
        'STRIPE_SECRET_KEY': 'SET' if os.environ.get('STRIPE_SECRET_KEY') else 'NOT_SET',
        'STRIPE_WEBHOOK_SECRET': 'SET' if os.environ.get('STRIPE_WEBHOOK_SECRET') else 'NOT_SET',
        'GOOGLE_SHEETS_CREDENTIALS': 'SET' if os.environ.get('GOOGLE_SHEETS_CREDENTIALS') else 'NOT_SET',
        'PORT': os.environ.get('PORT', 'NOT_SET'),
        'FLASK_ENV': os.environ.get('FLASK_ENV', 'NOT_SET')
    }
    
    html = """
    <html>
    <head><title>環境変数確認</title></head>
    <body style="font-family: Arial, sans-serif; padding: 20px;">
        <h1>環境変数確認</h1>
        <table border="1" style="border-collapse: collapse; width: 100%;">
            <tr><th>変数名</th><th>値</th></tr>
    """
    
    for key, value in env_vars.items():
        status_color = "green" if value != "NOT_SET" else "red"
        html += f'<tr><td>{key}</td><td style="color: {status_color};">{value}</td></tr>'
    
    html += """
        </table>
        <p><small>※ セキュリティのため、一部の値は「SET」と表示されます</small></p>
    </body>
    </html>
    """
    
    return html

@app.route("/create-rich-menu", methods=['GET'])
def create_rich_menu_endpoint():
    """リッチメニュー作成エンドポイント"""
    try:
        rich_menu_id = create_rich_menu()
        if rich_menu_id:
            return f"Rich menu created successfully! ID: {rich_menu_id}"
        else:
            return "Failed to create rich menu"
    except Exception as e:
        return f"Error: {str(e)}"

@app.route("/delete-rich-menu", methods=['GET'])
def delete_rich_menu_endpoint():
    """リッチメニュー削除エンドポイント"""
    try:
        with ApiClient(configuration) as api_client:
            messaging_api = MessagingApi(api_client)
            
            # 既存のリッチメニューを削除
            rich_menus = messaging_api.get_rich_menu_list()
            deleted_count = 0
            for rich_menu in rich_menus.richmenus:
                messaging_api.delete_rich_menu(rich_menu.rich_menu_id)
                deleted_count += 1
                logger.info(f"Deleted rich menu: {rich_menu.rich_menu_id}")
            
            return f"Deleted {deleted_count} rich menus successfully"
    except Exception as e:
        return f"Error deleting rich menus: {str(e)}"

@app.route("/callback", methods=['POST'])
def callback():
    logger.info("Webhook受信")
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    logger.info(f"Received webhook: {body[:100]}...")  # ログ追加
    try:
        handler.handle(body, signature)
    except InvalidSignatureError as e:
        logger.error(f"Invalid signature error: {e}")  # ログ追加
        abort(400)
    except Exception as e:
        logger.error(f"Unexpected error: {e}")  # ログ追加
        abort(500)
    return 'OK'

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_text = event.message.text.strip()
    user_id = event.source.user_id
    logger.info(f"Received message from {user_id}: {user_text}")
    
    # reply変数を初期化
    reply = ""
    
    # ユーザー登録（初回利用時）
    if user_manager:
        user_info = user_manager.get_user_info(user_id)
        if not user_info:
            # 新規ユーザー登録
            success, message = user_manager.register_user(user_id, "LINE User")
            if success:
                logger.info(f"New user registered: {user_id}")
            else:
                logger.error(f"User registration failed: {message}")
    else:
        logger.warning("User management system not available")

    # リッチメニューやテキストコマンドに応じた返答
    if user_text in ["商品を追加"]:
        # ユーザーの状態を商品追加に設定
        set_user_state(user_id, 'product_add')
        # シート選択画面を表示
        flex_message = FlexMessage(
            alt_text="シート選択",
            contents=FlexContainer.from_dict(create_sheet_selection())
        )
        send_flex_message(event.reply_token, flex_message)
        return
    elif user_text in ["スプレッドシート登録"]:
        # ユーザーの状態をスプレッドシート登録に設定
        set_user_state(user_id, 'spreadsheet_register')
        # シート選択画面を表示
        flex_message = FlexMessage(
            alt_text="シート選択",
            contents=FlexContainer.from_dict(create_sheet_selection())
        )
        send_flex_message(event.reply_token, flex_message)
        return
    elif user_text in ["会社情報を更新"]:
        reply = "会社情報を更新するには、以下の形式で入力してください：\n\n"
        reply += "会社名:○○株式会社\n"
        reply += "日付:2024/01/15\n\n"
        reply += "または、\n"
        reply += "会社名:○○株式会社 日付:2024/01/15"
        send_text_message(event.reply_token, reply)
        return
    elif user_text in ["利用状況確認"]:
        if user_manager:
            summary = user_manager.get_usage_summary(user_id)
            send_text_message(event.reply_token, summary)
        else:
            send_text_message(event.reply_token, "利用状況の取得に失敗しました。")
        return
    elif user_text in ["プランアップグレード"]:
        flex_message = FlexMessage(
            alt_text="プラン選択",
            contents=FlexContainer.from_dict(create_plan_selection())
        )
        send_flex_message(event.reply_token, flex_message)
        return
    elif user_text in ["見積書を確認"]:
        reply = "現在の見積書を確認するには、Googleスプレッドシートを直接確認してください。\n\n"
        reply += "📊 共有スプレッドシートURL:\n"
        reply += f"https://docs.google.com/spreadsheets/d/{SHARED_SPREADSHEET_ID}\n\n"
        reply += "💡 独自のスプレッドシートを登録している場合は、そのスプレッドシートを確認してください。"
        send_text_message(event.reply_token, reply)
        return

    # スプレッドシート管理機能
    print(f"user_text: {user_text}")
    if re.search(r"スプレッドシート[\s　]*登録[：:]", user_text):
        print("スプレッドシート登録コマンドを検出")
        # 2行・1行両対応: 行ごとにURLとシート名を抽出
        url = None
        sheet_name = None
        for line in user_text.splitlines():
            if not url:
                m_url = re.search(r"https?://[\w\-./?%&=:#]+", line)
                if m_url:
                    url = m_url.group(0).strip()
            if not sheet_name:
                m_sheet = re.search(r"シート名[：:]?[\s　]*(.+)", line)
                if m_sheet:
                    sheet_name = m_sheet.group(1).strip()
        print(f"url: {url}, sheet_name: {sheet_name}")
        spreadsheet_id = extract_spreadsheet_id(url) if url else None
        if spreadsheet_id:
            # シート名が指定されていない場合は実際のシート名を取得
            if not sheet_name:
                try:
                    client = setup_google_sheets()
                    if client:
                        spreadsheet = client.open_by_key(spreadsheet_id)
                        # 最初のシートの名前を取得
                        first_sheet = spreadsheet.get_worksheet(0)
                        sheet_name = first_sheet.title
                        print(f"取得したシート名: {sheet_name}")
                    else:
                        sheet_name = "比較御見積書　ショート"  # フォールバック
                except Exception as e:
                    print(f"シート名取得エラー: {e}")
                    sheet_name = "比較御見積書　ショート"  # フォールバック
            success, message = user_manager.set_user_spreadsheet(user_id, spreadsheet_id, sheet_name)
            if success:
                reply = f"✅ スプレッドシートを登録しました！\n\n"
                reply += f"📊 スプレッドシートURL:\n"
                reply += f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}\n\n"
                reply += f"📋 シート名: {sheet_name}\n\n"
                reply += "これで商品データがこのスプレッドシートに反映されます。"
            else:
                reply = f"❌ 登録エラー: {message}"
        else:
            reply = "❌ スプレッドシートURLが正しくありません。\n\n"
            reply += "正しい形式：\n"
            reply += "スプレッドシート登録:https://docs.google.com/spreadsheets/d/xxxxxxx\n\n"
            reply += "または、シート名を指定：\n"
            reply += "スプレッドシート登録:https://docs.google.com/spreadsheets/d/xxxxxxx シート名:見積書\n\n"
            reply += "⚠️ 重要：\n"
            reply += "• 新しいスプレッドシートを作成してください\n"
            reply += "• スプレッドシートは共有設定で「編集者」に設定してください\n"
            reply += "• シート名を指定しない場合は、最初のシートが使用されます\n\n"
            reply += "📋 手順：\n"
            reply += "1. Googleスプレッドシートを新規作成\n"
            reply += "2. シート名を変更（例：「見積書」）\n"
            reply += "3. 共有設定で「編集者」に設定\n"
            reply += "4. URLをコピーして以下の形式で送信：\n"
            reply += "スプレッドシート登録:【URL】 シート名:【シート名】"
        send_text_message(event.reply_token, reply)
        return

    elif user_text == "スプレッドシート確認":
        print(f"スプレッドシート確認処理開始: user_id={user_id}")
        if user_manager:
            spreadsheet_id, sheet_name = user_manager.get_user_spreadsheet(user_id)
            print(f"取得結果: spreadsheet_id={spreadsheet_id}, sheet_name={sheet_name}")
            if spreadsheet_id:
                reply = f"📊 あなたのスプレッドシート\n\n"
                reply += f"スプレッドシートURL:\n"
                reply += f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}\n\n"
                reply += f"シート名: {sheet_name}"
            else:
                reply = f"📊 共有スプレッドシートを使用中\n\n"
                reply += f"スプレッドシートURL:\n"
                reply += f"https://docs.google.com/spreadsheets/d/{SHARED_SPREADSHEET_ID}\n\n"
                reply += f"シート名: {DEFAULT_SHEET_NAME}\n\n"
                reply += "💡 独自のスプレッドシートを使用したい場合は、以下の形式で登録してください：\n"
                reply += "スプレッドシート登録:https://docs.google.com/spreadsheets/d/xxxxxxx シート名:見積書"
        else:
            print("user_manager is None")
            reply = "❌ システムエラー: ユーザー管理システムが利用できません。"
        send_text_message(event.reply_token, reply)
        return

    elif user_text == "スプレッドシート登録":
        # ユーザーの状態をスプレッドシート登録に設定
        set_user_state(user_id, 'spreadsheet_register')
        # シート選択画面を表示
        flex_message = FlexMessage(
            alt_text="シート選択",
            contents=FlexContainer.from_dict(create_sheet_selection())
        )
        send_flex_message(event.reply_token, flex_message)
        return

    # Excel Online URLの処理
    elif re.search(r"Excel[\s　]*Online[\s　]*登録[：:]", user_text) or re.search(r"エクセル[\s　]*オンライン[\s　]*登録[：:]", user_text):
        print("Excel Online登録コマンドを検出")
        # URLを抽出
        url = None
        sheet_name = None
        for line in user_text.splitlines():
            if not url:
                m_url = re.search(r"https?://[\w\-./?%&=:#]+", line)
                if m_url:
                    url = m_url.group(0).strip()
            if not sheet_name:
                m_sheet = re.search(r"シート名[：:]?[\s　]*(.+)", line)
                if m_sheet:
                    sheet_name = m_sheet.group(1).strip()
        
        print(f"Excel Online URL: {url}, sheet_name: {sheet_name}")
        
        if url and excel_online_manager:
            # URLの妥当性をチェック
            is_valid, error_msg = excel_online_manager.validate_excel_url(url)
            if not is_valid:
                reply = f"❌ Excel Online URLが正しくありません: {error_msg}\n\n"
                reply += "正しい形式：\n"
                reply += "Excel Online登録:https://unimatlifejp-my.sharepoint.com/...\n\n"
                reply += "または、シート名を指定：\n"
                reply += "Excel Online登録:https://unimatlifejp-my.sharepoint.com/... シート名:見積書"
                send_text_message(event.reply_token, reply)
                return
            
            # ファイルIDを抽出
            file_id = excel_online_manager.extract_file_id_from_url(url)
            if not file_id:
                reply = "❌ Excel Online URLからファイルIDを抽出できませんでした。\n\n"
                reply += "正しいSharePoint/OneDrive URLを入力してください。"
                send_text_message(event.reply_token, reply)
                return
            
            # シート名が指定されていない場合は実際のシート名を取得
            if not sheet_name:
                try:
                    worksheets, error = excel_online_manager.get_worksheets(file_id)
                    if worksheets and not error:
                        sheet_name = worksheets[0]  # 最初のシートを使用
                        print(f"取得したシート名: {sheet_name}")
                    else:
                        sheet_name = "Sheet1"  # フォールバック
                        print(f"シート名取得エラー: {error}")
                except Exception as e:
                    print(f"シート名取得エラー: {e}")
                    sheet_name = "Sheet1"  # フォールバック
            
            # ユーザーのExcel Online設定を保存
            success, message = user_manager.set_user_excel_online(user_id, url, file_id, sheet_name)
            if success:
                reply = f"✅ Excel Onlineファイルを登録しました！\n\n"
                reply += f"📊 Excel Online URL:\n"
                reply += f"{url}\n\n"
                reply += f"📋 シート名: {sheet_name}\n\n"
                reply += "これで商品データがこのExcel Onlineファイルに反映されます。"
            else:
                reply = f"❌ 登録エラー: {message}"
        else:
            reply = "❌ Excel Online URLが正しくありません。\n\n"
            reply += "正しい形式：\n"
            reply += "Excel Online登録:https://unimatlifejp-my.sharepoint.com/...\n\n"
            reply += "または、シート名を指定：\n"
            reply += "Excel Online登録:https://unimatlifejp-my.sharepoint.com/... シート名:見積書\n\n"
            reply += "⚠️ 重要：\n"
            reply += "• SharePoint/OneDriveのExcel Onlineファイルを使用してください\n"
            reply += "• ファイルは共有設定で「編集者」に設定してください\n"
            reply += "• シート名を指定しない場合は、最初のシートが使用されます"
        send_text_message(event.reply_token, reply)
        return

    elif user_text == "Excel Online確認" or user_text == "エクセルオンライン確認":
        print(f"Excel Online確認処理開始: user_id={user_id}")
        if user_manager:
            excel_url, excel_file_id, excel_sheet_name = user_manager.get_user_excel_online(user_id)
            print(f"取得結果: excel_url={excel_url}, excel_file_id={excel_file_id}, excel_sheet_name={excel_sheet_name}")
            if excel_url:
                reply = f"📊 あなたのExcel Onlineファイル\n\n"
                reply += f"Excel Online URL:\n"
                reply += f"{excel_url}\n\n"
                reply += f"シート名: {excel_sheet_name}"
            else:
                reply = f"📊 共有スプレッドシートを使用中\n\n"
                reply += f"スプレッドシートURL:\n"
                reply += f"https://docs.google.com/spreadsheets/d/{SHARED_SPREADSHEET_ID}\n\n"
                reply += f"シート名: {DEFAULT_SHEET_NAME}\n\n"
                reply += "💡 Excel Onlineファイルを使用したい場合は、以下の形式で登録してください：\n"
                reply += "Excel Online登録:https://unimatlifejp-my.sharepoint.com/... シート名:見積書"
        else:
            print("user_manager is None")
            reply = "❌ システムエラー: ユーザー管理システムが利用できません。"
        send_text_message(event.reply_token, reply)
        return

    # それ以外は従来通りの案内＋データ解析・登録
    data = parse_estimate_data(user_text)
    if data:
        # 会社情報の更新か商品データの書き込みかを判定
        is_company_update = '社名' in data or '会社名' in data or '日付' in data
        is_product_data = '商品名' in data and '単価' in data and '数量' in data

        if is_company_update and not is_product_data:
            # 会社情報の更新
            success, message = update_company_info(data, user_id)
            if success:
                reply = f"会社情報を更新しました！\n\n"
                if '社名' in data:
                    reply += f"会社名: {data['社名']}\n"
                if '日付' in data:
                    reply += f"日付: {data['日付']}\n"
            else:
                reply = f"エラー: {message}"

        elif is_product_data:
            # 利用制限チェック
            if user_manager:
                can_use, limit_message = user_manager.check_usage_limit(user_id)
                if not can_use:
                    reply = f"❌ {limit_message}\n\n"
                    reply += "プランアップグレードをご検討ください。\n"
                    reply += "「メニュー」→「利用状況確認」で詳細を確認できます。"
                    send_text_message(event.reply_token, reply)
                    return
            else:
                print("User management system not available, skipping usage limit check")

            # 商品データの書き込み
            success, message = write_to_spreadsheet(data, user_id)
            if success:
                # 利用回数を記録
                if user_manager:
                    user_manager.increment_usage(user_id, "add_product", data)
                reply = f"✅ 見積書を作成しました！\n\n"
                reply += f"📋 登録内容:\n"
                reply += f"社名: {data.get('社名', 'N/A')}\n"
                reply += f"商品名: {data.get('商品名', 'N/A')}\n"
                reply += f"サイズ: {data.get('サイズ', 'N/A')}\n"
                reply += f"単価: {data.get('単価', 'N/A')}\n"
                reply += f"数量: {data.get('数量', 'N/A')}\n"
                reply += f"料金: {data.get('料金', 'N/A')}\n\n"
                reply += f"📊 スプレッドシートに反映されました。"
            else:
                reply = f"❌ 見積書作成エラー: {message}"
        else:
            reply = "データの形式が正しくありません。\n\n"
            reply += "【会社情報更新】\n"
            reply += "例: 会社名:ABC株式会社 日付:2024/01/15\n\n"
            reply += "【商品データ登録】\n"
            reply += "例: 社名:ABC株式会社 商品名:商品A サイズ:M 単価:1000 数量:5\n\n"
            reply += "【追加項目（シートによって利用可能）】\n"
            reply += "サイクル:月1回 設置場所:1階\n\n"
            reply += "【語尾指定（比較見積書系のみ）】\n"
            reply += "商品名:マット 現状  ← 現状用の列に書き込み\n"
            reply += "商品名:マット 当社  ← 当社用の列に書き込み"
    else:
        # データが解析できない場合のデフォルトメッセージ
        reply = "見積書作成システムへようこそ！\n\n"
        reply += "以下のコマンドが利用できます：\n\n"
        reply += "📝 商品を追加\n"
        reply += "📊 スプレッドシート登録\n"
        reply += "📊 Excel Online登録\n"
        reply += "🏢 会社情報を更新\n"
        reply += "📈 利用状況確認\n"
        reply += "💳 プランアップグレード\n\n"
        reply += "詳細は「メニュー」ボタンからご確認ください。"
    
    send_text_message(event.reply_token, reply)

@handler.add(PostbackEvent)
def handle_postback(event):
    """Postbackイベントの処理（ボタンクリック）"""
    user_id = event.source.user_id
    data = event.postback.data
    print(f"Received postback from {user_id}: {data}")
    
    # データをパース
    params = {}
    for item in data.split('&'):
        if '=' in item:
            key, value = item.split('=', 1)
            params[key] = value
    
    action = params.get('action', '')
    
    if action == 'add_product':
        # 商品選択画面を表示
        flex_message = FlexMessage(
            alt_text="商品選択",
            contents=FlexContainer.from_dict(create_product_selection())
        )
        send_flex_message(event.reply_token, flex_message)
        
    elif action == 'custom_product':
        # カスタム商品名入力の案内
        reply = "カスタム商品を追加するには、以下の形式で入力してください：\n\n"
        reply += "【基本項目】\n"
        reply += "商品名:○○○○\n"
        reply += "サイズ:○○\n"
        reply += "単価:○○○○\n"
        reply += "数量:○○\n\n"
        reply += "【追加項目（シートによって利用可能）】\n"
        reply += "サイクル:○○\n"
        reply += "設置場所:○○\n\n"
        reply += "【語尾指定（比較見積書系のみ）】\n"
        reply += "商品名:マット 現状  ← 現状用の列に書き込み\n"
        reply += "商品名:マット 当社  ← 当社用の列に書き込み\n\n"
        reply += "例：\n"
        reply += "商品名:オリジナルTシャツ\n"
        reply += "サイズ:L\n"
        reply += "単価:2000\n"
        reply += "数量:5"
        send_text_message(event.reply_token, reply)
        
    elif action == 'select_product':
        # サイズ選択画面を表示
        product = params.get('product', '')
        flex_message = FlexMessage(
            alt_text="サイズ選択",
            contents=FlexContainer.from_dict(create_size_selection(product))
        )
        send_flex_message(event.reply_token, flex_message)
        
    elif action == 'custom_price':
        # カスタム価格入力の案内
        product = params.get('product', '')
        reply = f"{product}のカスタム価格を設定するには、以下の形式で入力してください：\n\n"
        reply += f"商品名:{product}\n"
        reply += "サイズ:○○\n"
        reply += "単価:○○○○\n"
        reply += "数量:○○\n\n"
        reply += f"例：\n"
        reply += f"商品名:{product}\n"
        reply += "サイズ:L\n"
        reply += "単価:1800\n"
        reply += "数量:3"
        send_text_message(event.reply_token, reply)
        
    elif action == 'select_quantity':
        # 商品データをスプレッドシートに書き込み
        product = params.get('product', '')
        size = params.get('size', '')
        price = params.get('price', '')
        quantity = params.get('quantity', '')
        
        # デバッグ用ログ
        print(f"Processing quantity selection: product={product}, size={size}, price={price}, quantity={quantity}")
        
        # 利用制限チェック
        if user_manager:
            can_use, limit_message = user_manager.check_usage_limit(user_id)
            if not can_use:
                reply = f"❌ {limit_message}\n\n"
                reply += "プランアップグレードをご検討ください。\n"
                reply += "「メニュー」→「利用状況確認」で詳細を確認できます。"
                send_text_message(event.reply_token, reply)
                return
        else:
            print("User management system not available, skipping usage limit check")
        
        data = {
            '商品名': product,
            'サイズ': size,
            '単価': price,
            '数量': quantity,
            '料金': int(price) * int(quantity)
        }
        
        success, message = write_to_spreadsheet(data, user_id)
        
        if success:
            # 利用回数を記録
            if user_manager:
                user_manager.increment_usage(user_id, "add_product", data)
            
            reply = f"✅ 商品を追加しました！\n\n"
            reply += f"商品名: {product}\n"
            reply += f"サイズ: {size}\n"
            reply += f"単価: {price}円\n"
            reply += f"数量: {quantity}個\n"
            reply += f"合計: {data['料金']}円\n\n"
            reply += "続けて商品を追加する場合は「メニュー」と入力してください。"
        else:
            reply = f"❌ エラー: {message}"
        
        send_text_message(event.reply_token, reply)
        
    elif action == 'check_usage':
        # 利用状況確認
        if user_manager:
            summary = user_manager.get_usage_summary(user_id)
            send_text_message(event.reply_token, summary)
        else:
            print("User management system not available, skipping usage summary")
        
    elif action == 'update_company':
        # 会社情報更新の案内
        reply = "会社情報を更新するには、以下の形式で入力してください：\n\n"
        reply += "会社名:○○株式会社\n"
        reply += "日付:2024/01/15\n\n"
        reply += "または、\n"
        reply += "会社名:○○株式会社 日付:2024/01/15"
        send_text_message(event.reply_token, reply)
        
    elif action == 'view_estimate':
        # 見積書確認の案内
        reply = "現在の見積書を確認するには、Googleスプレッドシートを直接確認してください。\n\n"
        reply += "📊 共有スプレッドシートURL:\n"
        reply += f"https://docs.google.com/spreadsheets/d/{SHARED_SPREADSHEET_ID}\n\n"
        reply += "💡 独自のスプレッドシートを登録している場合は、そのスプレッドシートを確認してください。"
        send_text_message(event.reply_token, reply)
        return

    elif action == 'upgrade_plan':
        # プラン選択画面を表示
        if stripe_payment:
            flex_message = FlexMessage(
                alt_text="プラン選択",
                contents=FlexContainer.from_dict(create_plan_selection())
            )
            send_flex_message(event.reply_token, flex_message)
        else:
            reply = "申し訳ございません。決済システムが利用できません。"
            send_text_message(event.reply_token, reply)
    
    elif action == 'show_sheet_selection':
        # シート選択画面を表示
        flex_message = FlexMessage(
            alt_text="シート選択",
            contents=FlexContainer.from_dict(create_sheet_selection())
        )
        send_flex_message(event.reply_token, flex_message)
    
    elif action == 'select_sheet':
        # シート選択時の処理
        sheet_name = params.get('sheet', '')
        print(f"Sheet selection: {sheet_name} for user {user_id}")
        
        # 現在のスプレッドシート情報を取得
        if user_manager:
            current_spreadsheet_id, current_sheet_name = user_manager.get_user_spreadsheet(user_id)
            user_state = get_user_state(user_id)

            if user_state == 'spreadsheet_register':
                # スプレッドシート登録からの場合はシート変更のみ
                if current_spreadsheet_id and current_sheet_name != sheet_name:
                    success, message = user_manager.set_user_spreadsheet(user_id, current_spreadsheet_id, sheet_name)
                    if not success:
                        reply = f"❌ シート変更エラー: {message}\n\n"
                        reply += "スプレッドシートの登録からやり直してください。"
                        send_text_message(event.reply_token, reply)
                        return
                reply = f"✅ シートを変更しました！\n\n"
                reply += f"📊 スプレッドシートURL:\n"
                reply += f"https://docs.google.com/spreadsheets/d/{current_spreadsheet_id}\n\n"
                reply += f"📋 変更前シート: {current_sheet_name}\n"
                reply += f"📋 変更後シート: {sheet_name}\n\n"
                reply += "これで商品データが選択したシートに反映されます。"
                send_text_message(event.reply_token, reply)
                return
            elif user_state == 'product_add':
                # 商品追加からの場合は入力フォーマットのみ表示
                reply = f"📝 入力フォーマット（{sheet_name}）:\n"
                if sheet_name == "比較見積書 ロング":
                    reply += "商品名:○○○○\n"
                    reply += "単価:○○○○\n"
                    reply += "数量:○○\n"
                    reply += "サイクル:○○\n"
                    reply += "【語尾指定】\n"
                    reply += "商品名:マット 現状  ← 現状用の列に書き込み\n"
                    reply += "商品名:マット 当社  ← 当社用の列に書き込み\n"
                    reply += "例：\n"
                    reply += "商品名:マット 現状\n"
                    reply += "単価:2000\n"
                    reply += "数量:3\n"
                    reply += "サイクル:週2"
                elif sheet_name == "比較御見積書　ショート":
                    reply += "商品名:○○○○\n"
                    reply += "単価:○○○○\n"
                    reply += "数量:○○\n"
                    reply += "サイクル:○○\n"
                    reply += "【語尾指定】\n"
                    reply += "商品名:マット 現状  ← 現状用の列に書き込み\n"
                    reply += "商品名:マット 当社  ← 当社用の列に書き込み\n"
                    reply += "例：\n"
                    reply += "商品名:マット 現状\n"
                    reply += "単価:2000\n"
                    reply += "数量:3\n"
                    reply += "サイクル:週2"
                elif sheet_name == "新規見積書　ショート":
                    reply += "商品名:○○○○\n"
                    reply += "単価:○○○○\n"
                    reply += "数量:○○\n\n"
                    reply += "例：\n"
                    reply += "商品名:マット\n"
                    reply += "単価:2000\n"
                    reply += "数量:3"
                else:
                    reply += "商品名:○○○○\n"
                    reply += "設置場所:○○\n"
                    reply += "サイクル:○○\n"
                    reply += "数量:○○\n"
                    reply += "単価:○○○○\n\n"
                    reply += "例：\n"
                    reply += "商品名:マット\n"
                    reply += "設置場所:玄関\n"
                    reply += "サイクル:週2\n"
                    reply += "数量:3\n"
                    reply += "単価:2000"
                send_text_message(event.reply_token, reply)
                return
        else:
            reply = "❌ システムエラー: ユーザー管理システムが利用できません。"
            send_text_message(event.reply_token, reply)
            return
    
    elif action == 'select_plan':
        # プラン選択時の処理
        plan_type = params.get('plan', '')
        print(f"Plan selection: {plan_type} for user {user_id}")
        
        if stripe_payment and user_manager:
            print("Stripe payment and user manager are available")
            # Stripeチェックアウトセッションを作成
            success, result = stripe_payment.create_checkout_session(plan_type, user_id)
            print(f"Checkout session result: success={success}, result={result}")
            
            if success:
                checkout_url = result['checkout_url']
                plan_info = result['plan_info']
                
                reply = f"💳 {plan_info['name']}の決済\n\n"
                reply += f"料金: {plan_info['price']}円\n"
                reply += f"内容: {plan_info['description']}\n\n"
                reply += "以下のURLから決済を完了してください：\n"
                reply += f"{checkout_url}\n\n"
                reply += "決済完了後、プランが自動的に更新されます。"
                
                # 決済情報をセッションに保存
                user_sessions[user_id] = {
                    'plan_type': plan_type,
                    'session_id': result['session_id']
                }
            else:
                reply = f"決済URLの作成に失敗しました: {result}"
                print(f"Payment URL creation failed: {result}")
        else:
            reply = "申し訳ございません。決済システムが利用できません。"
            print(f"Payment system not available: stripe_payment={stripe_payment}, user_manager={user_manager}")
        
        send_text_message(event.reply_token, reply)

def send_text_message(reply_token, text):
    """テキストメッセージを送信"""
    try:
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            line_bot_api.reply_message_with_http_info(
                ReplyMessageRequest(
                    reply_token=reply_token,
                    messages=[TextMessage(text=text)]
                )
            )
        print(f"Text message sent: {text}")
    except Exception as e:
        print(f"Error sending text message: {e}")

def send_flex_message(reply_token, flex_message):
    """Flexメッセージを送信"""
    try:
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            line_bot_api.reply_message_with_http_info(
                ReplyMessageRequest(
                    reply_token=reply_token,
                    messages=[flex_message]
                )
            )
        print(f"Flex message sent")
    except Exception as e:
        print(f"Error sending flex message: {e}")

@app.route("/payment/success", methods=['GET'])
def payment_success():
    """Stripe決済完了時の処理"""
    user_id = request.args.get('user_id')
    plan_type = request.args.get('plan')
    
    if user_id and plan_type and user_manager:
        # ユーザーのプランを更新
        success = user_manager.upgrade_plan(user_id, plan_type)
        if success:
            return """
            <html>
            <head><title>決済完了</title></head>
            <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
                <h1>✅ 決済が完了しました！</h1>
                <p>プランが正常に更新されました。</p>
                <p>LINE Botに戻ってご確認ください。</p>
                <p><a href="https://line.me/R/ti/p/@your-bot-id">LINE Botに戻る</a></p>
            </body>
            </html>
            """
        else:
            return """
            <html>
            <head><title>エラー</title></head>
            <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
                <h1>❌ エラーが発生しました</h1>
                <p>プランの更新に失敗しました。</p>
                <p>サポートにお問い合わせください。</p>
            </body>
            </html>
            """
    
    return """
    <html>
    <head><title>エラー</title></head>
    <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
        <h1>❌ エラーが発生しました</h1>
        <p>決済情報が見つかりません。</p>
    </body>
    </html>
    """

@app.route("/payment/cancel", methods=['GET'])
def payment_cancel():
    """Stripe決済キャンセル時の処理"""
    return """
    <html>
    <head><title>決済キャンセル</title></head>
    <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
        <h1>❌ 決済がキャンセルされました</h1>
        <p>LINE Botに戻ってお試しください。</p>
        <p><a href="https://line.me/R/ti/p/@your-bot-id">LINE Botに戻る</a></p>
    </body>
    </html>
    """

@app.route("/payment/portal_return", methods=['GET'])
def payment_portal_return():
    """Stripeカスタマーポータルからの戻り"""
    return """
    <html>
    <head><title>設定完了</title></head>
    <body style="font-family: Arial, sans-serif; text-align: center; padding: 50px;">
        <h1>✅ 設定が完了しました</h1>
        <p>LINE Botに戻ってご確認ください。</p>
        <p><a href="https://line.me/R/ti/p/@your-bot-id">LINE Botに戻る</a></p>
    </body>
    </html>
    """

@app.route("/stripe/webhook", methods=['POST'])
def stripe_webhook():
    """Stripe Webhookの処理"""
    payload = request.get_data()
    sig_header = request.headers.get('Stripe-Signature')
    webhook_secret = os.environ.get('STRIPE_WEBHOOK_SECRET', '')
    
    if not webhook_secret:
        return "Webhook secret not configured", 400
    
    if stripe_payment:
        success, result = stripe_payment.handle_webhook(payload, sig_header, webhook_secret)
        if success:
            return "Webhook processed successfully", 200
        else:
            return f"Webhook error: {result}", 400
    
    return "Stripe payment system not available", 500

# ユーザー状態管理
user_states = {}  # user_id -> state (spreadsheet_register, product_add, etc.)

def get_user_state(user_id):
    """ユーザーの現在の状態を取得"""
    return user_states.get(user_id, 'product_add')  # デフォルトは商品追加

def set_user_state(user_id, state):
    """ユーザーの状態を設定"""
    user_states[user_id] = state
    logger.info(f"User {user_id} state set to: {state}")

if __name__ == "__main__":
    logger.info("=== アプリケーション起動開始 ===")
    logger.info("環境変数の確認:")
    logger.info(f"MS_CLIENT_ID: {os.environ.get('MS_CLIENT_ID', 'NOT_SET')}")
    logger.info(f"MS_CLIENT_SECRET: {os.environ.get('MS_CLIENT_SECRET', 'NOT_SET')[:10]}..." if os.environ.get('MS_CLIENT_SECRET') else 'NOT_SET')
    logger.info(f"MS_TENANT_ID: {os.environ.get('MS_TENANT_ID', 'NOT_SET')}")
    logger.info(f"LINE_CHANNEL_ACCESS_TOKEN: {os.environ.get('LINE_CHANNEL_ACCESS_TOKEN', 'NOT_SET')[:10]}..." if os.environ.get('LINE_CHANNEL_ACCESS_TOKEN') else 'NOT_SET')
    logger.info(f"LINE_CHANNEL_SECRET: {os.environ.get('LINE_CHANNEL_SECRET', 'NOT_SET')[:10]}..." if os.environ.get('LINE_CHANNEL_SECRET') else 'NOT_SET')
    logger.info(f"SHARED_SPREADSHEET_ID: {os.environ.get('SHARED_SPREADSHEET_ID', 'NOT_SET')}")
    logger.info(f"DEFAULT_SHEET_NAME: {os.environ.get('DEFAULT_SHEET_NAME', 'NOT_SET')}")
    logger.info(f"STRIPE_SECRET_KEY: {os.environ.get('STRIPE_SECRET_KEY', 'NOT_SET')[:10]}..." if os.environ.get('STRIPE_SECRET_KEY') else 'NOT_SET')
    logger.info(f"STRIPE_WEBHOOK_SECRET: {os.environ.get('STRIPE_WEBHOOK_SECRET', 'NOT_SET')[:10]}..." if os.environ.get('STRIPE_WEBHOOK_SECRET') else 'NOT_SET')
    logger.info(f"GOOGLE_SHEETS_CREDENTIALS: {'SET' if os.environ.get('GOOGLE_SHEETS_CREDENTIALS') else 'NOT_SET'}")
    logger.info("=== アプリケーション起動完了 ===")
    
    port = int(os.environ.get('PORT', 5002))
    debug_mode = os.environ.get('FLASK_ENV') == 'development'
    app.run(host='0.0.0.0', port=port, debug=debug_mode)
