import gspread
from google.oauth2.service_account import Credentials
from datetime import datetime

# 入力データ
company_name = '株式会社サンプル'

# 今日の日付を自動取得
today = datetime.today()
year = today.year
month = today.month
day = today.day

# 商品データ（必要な数だけ追加）
products = [
    {"name": "りんご", "size": "L", "unit_price": 100, "quantity": 5},
    {"name": "みかん", "size": "M", "unit_price": 80, "quantity": 10},
    {"name": "バナナ", "size": "S", "unit_price": 120, "quantity": 3},
    # ここに商品を追加
]

# スコープと認証
SCOPES = ['https://www.googleapis.com/auth/spreadsheets']
creds = Credentials.from_service_account_file('gsheet_service_account.json', scopes=SCOPES)
gc = gspread.authorize(creds)

# スプレッドシートとワークシートの取得
spreadsheet = gc.open_by_url('https://docs.google.com/spreadsheets/d/1GkJ8OYwIIMnYqxcwVBNArvk2byFL3UlGHgkyTiV6QU0/edit?gid=498384544#gid=498384544')
worksheet = spreadsheet.worksheet('比較見積書 ロング')  # シート名は実際のものに合わせてください

# A2:H3の全セルに社名を入力
company_range = [[company_name for _ in range(8)] for _ in range(2)]
worksheet.update('A2:H3', company_range)

# M2=年、N2=月、O2=日として日付を入力（リストで渡す）
worksheet.update('M2', [[year]])
worksheet.update('N2', [[month]])
worksheet.update('O2', [[day]])

print('A2:H3に社名、M2:O2に日付を入力しました')

start_row = 19
max_rows = 18

# 左側（A～D列）
left_block = []
for i in range(max_rows):
    if i < len(products):
        p = products[i]
        left_block.append([p["name"], p["size"], p["unit_price"], p["quantity"]])
    else:
        left_block.append(["", "", "", ""])
worksheet.update(f'A{start_row}:D{start_row+max_rows-1}', left_block)

# 右側（I～L列）
right_block = []
for i in range(max_rows):
    if i < len(products):
        p = products[i]
        right_block.append([p["name"], p["size"], p["unit_price"], p["quantity"]])
    else:
        right_block.append(["", "", "", ""])
worksheet.update(f'I{start_row}:L{start_row+max_rows-1}', right_block)

print('商品データをA19:D36, I19:L36に一括で書き込みました')

print("操作中のシート名:", worksheet.title)
print("操作中のスプレッドシートURL:", spreadsheet.url)

# A19:D36の内容をAPIで取得してprint
print("A19:D36の現在の値:", worksheet.get('A19:D36'))