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
from user_management import UserManager
from stripe_payment import StripePayment

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

# ユーザー管理システムの初期化
try:
    user_manager = UserManager()
    print("User management system initialized successfully")
except Exception as e:
    print(f"User management system initialization error: {e}")
    user_manager = None

# Stripe決済システムの初期化
try:
    stripe_payment = StripePayment()
    print("Stripe payment system initialized successfully")
except Exception as e:
    print(f"Stripe payment system initialization error: {e}")
    stripe_payment = None

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
        if key in ['社名', '会社名', '商品名', 'サイズ', '単価', '数量', '日付']:
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
    """スプレッドシートにデータを書き込み（顧客別対応）"""
    try:
        print(f"開始: スプレッドシート書き込み処理")
        
        # 顧客のスプレッドシートIDを取得
        if user_id and user_manager:
            spreadsheet_id, sheet_name = user_manager.get_user_spreadsheet(user_id)
            if not spreadsheet_id:
                spreadsheet_id = SPREADSHEET_ID  # デフォルト
                sheet_name = SHEET_NAME
        else:
            spreadsheet_id = SPREADSHEET_ID
            sheet_name = SHEET_NAME
        
        client = setup_google_sheets()
        if not client:
            print("エラー: Google Sheets接続失敗")
            return False, "Google Sheets接続エラー"
        
        print(f"成功: Google Sheets接続")
        sheet = client.open_by_key(spreadsheet_id).worksheet(sheet_name)
        print(f"成功: シート '{sheet_name}' を開きました")
        
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

def create_rich_menu():
    """リッチメニューを作成"""
    try:
        with ApiClient(configuration) as api_client:
            messaging_api = MessagingApi(api_client)
            
            # 既存のリッチメニューを削除
            try:
                rich_menus = messaging_api.get_rich_menu_list()
                for rich_menu in rich_menus.richmenus:
                    messaging_api.delete_rich_menu(rich_menu.rich_menu_id)
                    print(f"Deleted existing rich menu: {rich_menu.rich_menu_id}")
            except Exception as e:
                print(f"Error deleting existing rich menus: {e}")
            
            rich_menu_dict = {
                "size": {"width": 1200, "height": 405},
                "selected": False,
                "name": "見積書作成メニュー",
                "chatBarText": "メニュー",
                "areas": [
                    {
                        "bounds": {"x": 0, "y": 0, "width": 200, "height": 405},
                        "action": {"type": "message", "label": "商品を追加", "text": "商品を追加"}
                    },
                    {
                        "bounds": {"x": 200, "y": 0, "width": 200, "height": 405},
                        "action": {"type": "message", "label": "プランアップグレード", "text": "プランアップグレード"}
                    },
                    {
                        "bounds": {"x": 400, "y": 0, "width": 200, "height": 405},
                        "action": {"type": "message", "label": "会社情報を更新", "text": "会社情報を更新"}
                    },
                    {
                        "bounds": {"x": 600, "y": 0, "width": 200, "height": 405},
                        "action": {"type": "message", "label": "利用状況確認", "text": "利用状況確認"}
                    },
                    {
                        "bounds": {"x": 800, "y": 0, "width": 200, "height": 405},
                        "action": {"type": "message", "label": "見積書を確認", "text": "見積書を確認"}
                    },
                    {
                        "bounds": {"x": 1000, "y": 0, "width": 200, "height": 405},
                        "action": {"type": "message", "label": "スプレッドシート登録", "text": "スプレッドシート登録"}
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
                print(f"Deleted rich menu: {rich_menu.rich_menu_id}")
            
            return f"Deleted {deleted_count} rich menus successfully"
    except Exception as e:
        return f"Error deleting rich menus: {str(e)}"

@app.route("/callback", methods=['POST'])
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
    user_text = event.message.text.strip()
    user_id = event.source.user_id
    print(f"Received message from {user_id}: {user_text}")
    
    # ユーザー登録（初回利用時）
    if user_manager:
        user_info = user_manager.get_user_info(user_id)
        if not user_info:
            # 新規ユーザー登録
            success, message = user_manager.register_user(user_id, "LINE User")
            if success:
                print(f"New user registered: {user_id}")
            else:
                print(f"User registration failed: {message}")
    else:
        print("User management system not available")

    # リッチメニューやテキストコマンドに応じた返答
    if user_text in ["商品を追加"]:
        reply = "カスタム商品を追加するには、以下の形式で入力してください：\n\n商品名:○○○○\nサイズ:○○\n単価:○○○○\n数量:○○\n\n例：\n商品名:オリジナルTシャツ\nサイズ:L\n単価:2000\n数量:5"
        send_text_message(event.reply_token, reply)
        return
    elif user_text in ["会社情報を更新"]:
        reply = "会社情報を更新するには、以下の形式で入力してください：\n\n会社名:○○株式会社\n日付:2024/01/15\n\nまたは、\n会社名:○○株式会社 日付:2024/01/15"
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
        reply = "現在の見積書を確認するには、Googleスプレッドシートを直接確認してください。\n\nスプレッドシートURL:\nhttps://docs.google.com/spreadsheets/d/1GkJ8OYwIIMnYqxcwVBNArvk2byFL3UlGHgkyTiV6QU0"
        send_text_message(event.reply_token, reply)
        return

    # スプレッドシート管理機能
    elif user_text.startswith("スプレッドシート登録:"):
        # スプレッドシートURLからIDを抽出
        url = user_text.replace("スプレッドシート登録:", "").strip()
        spreadsheet_id = extract_spreadsheet_id(url)
        
        if spreadsheet_id:
            success, message = user_manager.set_user_spreadsheet(user_id, spreadsheet_id)
            if success:
                reply = f"✅ スプレッドシートを登録しました！\n\n"
                reply += f"スプレッドシートURL:\n"
                reply += f"https://docs.google.com/spreadsheets/d/{spreadsheet_id}\n\n"
                reply += "これで商品データがこのスプレッドシートに反映されます。"
            else:
                reply = f"❌ 登録エラー: {message}"
        else:
            reply = "❌ スプレッドシートURLが正しくありません。\n\n"
            reply += "正しい形式：\n"
            reply += "スプレッドシート登録:https://docs.google.com/spreadsheets/d/..."
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
                reply = "❌ スプレッドシートが登録されていません。\n\n"
                reply += "登録方法：\n"
                reply += "スプレッドシート登録:https://docs.google.com/spreadsheets/d/..."
        else:
            print("user_manager is None")
            reply = "❌ システムエラー: ユーザー管理システムが利用できません。"
        send_text_message(event.reply_token, reply)
        return

    elif user_text == "スプレッドシート登録":
        reply = "📝 スプレッドシートを登録してください\n\n"
        reply += "以下の形式でGoogleスプレッドシートのURLを送信してください：\n\n"
        reply += "スプレッドシート登録:https://docs.google.com/spreadsheets/d/xxxxxxx\n\n"
        reply += "⚠️ 重要：\n"
        reply += "• 新しいスプレッドシートを作成してください\n"
        reply += "• スプレッドシートは共有設定で「編集者」に設定してください\n"
        reply += "• 見積書フォーマットのシート名は「比較見積書 ロング」を推奨します\n\n"
        reply += "📋 手順：\n"
        reply += "1. Googleスプレッドシートを新規作成\n"
        reply += "2. シート名を「比較見積書 ロング」に変更\n"
        reply += "3. 共有設定で「編集者」に設定\n"
        reply += "4. URLをコピーして以下の形式で送信：\n"
        reply += "スプレッドシート登録:【あなたのスプレッドシートURL】"
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
            reply += "または「メニュー」と入力してボタン選択式で入力してください。"
        send_text_message(event.reply_token, reply)
        return

    # 何も該当しない場合のみ案内
    reply = "見積書作成システムへようこそ！\n\n"
    reply += "以下の方法で入力できます：\n\n"
    reply += "1️⃣ **ボタン選択式（推奨）**\n"
    reply += "「メニュー」と入力してボタンで選択\n\n"
    reply += "2️⃣ **テキスト入力**\n"
    reply += "【会社情報更新】\n"
    reply += "例: 会社名:ABC株式会社 日付:2024/01/15\n\n"
    reply += "【商品データ登録】\n"
    reply += "例: 社名:ABC株式会社 商品名:商品A サイズ:M 単価:1000 数量:5"
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
        reply += "商品名:○○○○\n"
        reply += "サイズ:○○\n"
        reply += "単価:○○○○\n"
        reply += "数量:○○\n\n"
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
        reply += "スプレッドシートURL:\n"
        reply += "https://docs.google.com/spreadsheets/d/1GkJ8OYwIIMnYqxcwVBNArvk2byFL3UlGHgkyTiV6QU0"
        send_text_message(event.reply_token, reply)

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

if __name__ == "__main__":
    port = int(os.environ.get('PORT', 5002))
    debug_mode = os.environ.get('FLASK_ENV') == 'development'
    app.run(host='0.0.0.0', port=port, debug=debug_mode)
