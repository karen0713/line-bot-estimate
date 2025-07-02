from flask import Flask, request, abort
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import Configuration, ApiClient, MessagingApi, ReplyMessageRequest, TextMessage
from linebot.v3.webhooks import MessageEvent, TextMessageContent
import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime
import re
import os

app = Flask(__name__)

# 環境変数から設定を取得
LINE_CHANNEL_ACCESS_TOKEN = os.environ.get('LINE_CHANNEL_ACCESS_TOKEN', 'Khehk/dQp536fyLT0u0UVSzBWh7zCNYDGPODIi5KtpNmkp1QJXc5kDKVlTaavNYW/12lK/HLF001axW4WLfoOXqLxTNMaXb6E6BnqtrAIxyoRP56Nw0g41L6JT2An3cA86Nl6tHqUY8ul5gP+9L8BgdB04t89/1O/w1cDnyilFU=')
LINE_CHANNEL_SECRET = os.environ.get('LINE_CHANNEL_SECRET', '8326aecb26b4e9c41ef8d702b73c6617')

# Google Sheets設定
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
SPREADSHEET_ID = os.environ.get('SPREADSHEET_ID', '1GkJ8OYwIIMnYqxcwVBNArvk2byFL3UlGHgkyTiV6QU0')
SHEET_NAME = os.environ.get('SHEET_NAME', '比較見積書 ロング')

configuration = Configuration(access_token=LINE_CHANNEL_ACCESS_TOKEN)
handler = WebhookHandler(LINE_CHANNEL_SECRET)

def setup_google_sheets():
    """Google Sheets APIの設定"""
    try:
        # 環境変数からサービスアカウント情報を取得
        service_account_info = os.environ.get('GOOGLE_SERVICE_ACCOUNT_JSON')
        if service_account_info:
            import json
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
    """LINEメッセージから見積書データを解析"""
    # 例: "社名:ABC株式会社 商品名:商品A サイズ:M 単価:1000 数量:5"
    # 例: "会社名:ABC株式会社 日付:2024/01/15"
    data = {}
    
    # 改行をスペースに変換して処理しやすくする
    text = text.replace('\n', ' ')
    
    # パターンマッチングでデータを抽出
    patterns = {
        '社名': r'社名[：:]\s*([^\s]+)',
        '会社名': r'会社名[：:]\s*([^\s]+)',
        '商品名': r'商品名[：:]*\s*([^\s]+)',  # コロンが抜けている場合も対応
        'サイズ': r'サイズ[：:]\s*([^\s]+)',
        '単価': r'単価[：:]\s*(\d+)',
        '数量': r'数量[：:]\s*(\d+)',
        '日付': r'日付[：:]\s*([^\s]+)'
    }
    
    for key, pattern in patterns.items():
        match = re.search(pattern, text)
        if match:
            data[key] = match.group(1)
    
    # 社名と会社名を統一
    if '会社名' in data and '社名' not in data:
        data['社名'] = data['会社名']
    
    # 料金を計算（商品データがある場合のみ）
    if '単価' in data and '数量' in data:
        try:
            unit_price = int(data['単価'])
            quantity = int(data['数量'])
            data['料金'] = unit_price * quantity
        except ValueError:
            data['料金'] = 0
    
    return data

def write_to_spreadsheet(data):
    """スプレッドシートにデータを書き込み"""
    try:
        print(f"開始: スプレッドシート書き込み処理")
        client = setup_google_sheets()
        if not client:
            print("エラー: Google Sheets接続失敗")
            return False, "Google Sheets接続エラー"
        
        print(f"成功: Google Sheets接続")
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
        print(f"成功: シート '{SHEET_NAME}' を開きました")
        
        # 現在の日付を取得
        current_date = datetime.now().strftime('%Y/%m/%d')
        
        # 見積書フォーマットに合わせて、A-D列が36行目まで埋まったらI-L列に書き込み
        # A列: 商品名, B列: サイズ, C列: 単価, D列: 数量
        # I列: 商品名, J列: サイズ, K列: 単価, L列: 数量
        
        # 既存データの行数を確認
        existing_data = sheet.get_all_values()
        print(f"既存データ行数: {len(existing_data)}")
        
        # A-D列の使用状況を確認（19行目から36行目まで）
        ad_used_rows = 0
        for row in range(18, min(36, len(existing_data))):  # 19行目から36行目まで
            if any(existing_data[row][:4]):  # A-D列のいずれかにデータがあるかチェック
                ad_used_rows += 1
        
        print(f"A-D列使用済み行数: {ad_used_rows}")
        
        # A-D列が36行目まで埋まっているかチェック
        if ad_used_rows >= 18:  # 19行目から36行目まで = 18行
            # I-L列に書き込み（19行目から開始）
            next_row = 19
            range_name = f"I{next_row}:L{next_row}"
            print(f"A-D列が36行目まで埋まっているため、I-L列の{next_row}行目に書き込み")
        else:
            # A-D列に書き込み（19行目から順番に）
            next_row = 19 + ad_used_rows
            range_name = f"A{next_row}:D{next_row}"
            print(f"A-D列の{next_row}行目に書き込み")
        
        print(f"書き込み行: {next_row} ({range_name})")
        
        # 書き込むデータを準備
        write_data = [[
            data.get('商品名', ''),
            data.get('サイズ', ''),
            data.get('単価', ''),
            data.get('数量', '')
        ]]
        
        print(f"書き込みデータ: {write_data}")
        print(f"書き込み範囲: {range_name}")
        
        # データを書き込み
        sheet.update(range_name, write_data)
        
        print(f"成功: データを{next_row}行目の{range_name}に書き込みました")
        return True, f"データを{next_row}行目の{range_name}に正常に書き込みました"
        
    except Exception as e:
        print(f"Spreadsheet write error: {e}")
        return False, f"書き込みエラー: {str(e)}"

def update_company_info(data):
    """会社名と日付を更新"""
    try:
        print(f"開始: 会社情報更新処理")
        client = setup_google_sheets()
        if not client:
            print("エラー: Google Sheets接続失敗")
            return False, "Google Sheets接続エラー"
        
        sheet = client.open_by_key(SPREADSHEET_ID).worksheet(SHEET_NAME)
        updates = []
        
        # 会社名を更新（A2:H3セル）
        if '社名' in data:
            company_values = [
                [data['社名']] + [''] * 7,
                [''] * 8
            ]
            sheet.update('A2:H3', company_values)
            updates.append(f"会社名: {data['社名']}")
            print(f"会社名を更新: {data['社名']}")
        
        # 日付を更新（M2:Q2セル）
        if '日付' in data:
            date_values = [
                [data['日付']] + [''] * 4
            ]
            sheet.update('M2:Q2', date_values)
            updates.append(f"日付: {data['日付']}")
            print(f"日付を更新: {data['日付']}")
        
        if updates:
            return True, f"更新完了: {', '.join(updates)}"
        else:
            return False, "更新するデータがありません"
        
    except Exception as e:
        print(f"Company info update error: {e}")
        return False, f"更新エラー: {str(e)}"

@app.route("/", methods=['GET'])
def index():
    return "LINE Bot Server is running!"

@app.route("/webhook", methods=['POST'])
def callback():
    signature = request.headers['X-Line-Signature']
    body = request.get_data(as_text=True)
    print(f"Received webhook: {body[:100]}...")  # ログ追加
    try:
        handler.handle(body, signature)
    except InvalidSignatureError as e:
        print(f"Invalid signature error: {e}")  # ログ追加
        abort(400)
    except Exception as e:
        print(f"Unexpected error: {e}")  # ログ追加
        abort(500)
    return 'OK'

@handler.add(MessageEvent, message=TextMessageContent)
def handle_message(event):
    user_text = event.message.text
    print(f"Received message: {user_text}")  # ログ追加
    
    # 見積書データを解析
    data = parse_estimate_data(user_text)
    
    if data:
        # 会社情報の更新か商品データの書き込みかを判定
        is_company_update = '社名' in data or '会社名' in data or '日付' in data
        is_product_data = '商品名' in data and '単価' in data and '数量' in data
        
        if is_company_update and not is_product_data:
            # 会社情報の更新
            success, message = update_company_info(data)
            
            if success:
                reply = f"会社情報を更新しました！\n\n"
                if '社名' in data:
                    reply += f"会社名: {data['社名']}\n"
                if '日付' in data:
                    reply += f"日付: {data['日付']}\n"
            else:
                reply = f"エラー: {message}"
                
        elif is_product_data:
            # 商品データの書き込み
            success, message = write_to_spreadsheet(data)
            
            if success:
                reply = f"見積書データを登録しました！\n\n"
                reply += f"社名: {data.get('社名', 'N/A')}\n"
                reply += f"商品名: {data.get('商品名', 'N/A')}\n"
                reply += f"サイズ: {data.get('サイズ', 'N/A')}\n"
                reply += f"単価: {data.get('単価', 'N/A')}円\n"
                reply += f"数量: {data.get('数量', 'N/A')}\n"
                reply += f"料金: {data.get('料金', 'N/A')}円"
            else:
                reply = f"エラー: {message}"
        else:
            reply = "データの形式が正しくありません。\n\n"
            reply += "【会社情報更新】\n"
            reply += "例: 会社名:ABC株式会社 日付:2024/01/15\n\n"
            reply += "【商品データ登録】\n"
            reply += "例: 社名:ABC株式会社 商品名:商品A サイズ:M 単価:1000 数量:5"
    else:
        reply = "データの形式が正しくありません。\n\n"
        reply += "【会社情報更新】\n"
        reply += "例: 会社名:ABC株式会社 日付:2024/01/15\n\n"
        reply += "【商品データ登録】\n"
        reply += "例: 社名:ABC株式会社 商品名:商品A サイズ:M 単価:1000 数量:5"
    
    try:
        with ApiClient(configuration) as api_client:
            line_bot_api = MessagingApi(api_client)
            line_bot_api.reply_message_with_http_info(
                ReplyMessageRequest(
                    reply_token=event.reply_token,
                    messages=[TextMessage(text=reply)]
                )
            )
        print(f"Reply sent: {reply}")  # ログ追加
    except Exception as e:
        print(f"Error sending reply: {e}")  # ログ追加

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5002))
    app.run(host='0.0.0.0', port=port, debug=False)
