from flask import Flask, request, abort, redirect, url_for
from linebot.v3 import WebhookHandler
from linebot.v3.exceptions import InvalidSignatureError
from linebot.v3.messaging import Configuration, ApiClient, MessagingApi, ReplyMessageRequest, TextMessage, FlexMessage, FlexContainer
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
    
    # 特殊コマンドの処理
    if user_text.lower() in ['メニュー', 'menu', '開始', 'start']:
        # メインメニューを表示
        flex_message = FlexMessage(
            alt_text="見積書作成システム",
            contents=FlexContainer.from_dict(create_main_menu())
        )
        send_flex_message(event.reply_token, flex_message)
        return
    
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
            success, message = write_to_spreadsheet(data)
            
            if success:
                # 利用回数を記録
                if user_manager:
                    user_manager.increment_usage(user_id, "add_product", data)
                
                reply = f"✅ 見積書データを登録しました！\n\n"
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
            reply += "例: 社名:ABC株式会社 商品名:商品A サイズ:M 単価:1000 数量:5\n\n"
            reply += "または「メニュー」と入力してボタン選択式で入力してください。"
    else:
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
        
        success, message = write_to_spreadsheet(data)
        
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
    app.run(host='0.0.0.0', port=port, debug=False)
