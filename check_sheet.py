import gspread
from google.oauth2.service_account import Credentials

# Google Sheets設定
SCOPES = [
    'https://www.googleapis.com/auth/spreadsheets',
    'https://www.googleapis.com/auth/drive'
]
SPREADSHEET_ID = '1GkJ8OYwIIMnYqxcwVBNArvk2byFL3UlGHgkyTiV6QU0'

def check_sheets():
    """スプレッドシートのシート名を確認"""
    try:
        creds = Credentials.from_service_account_file(
            'gsheet_service_account.json', scopes=SCOPES)
        client = gspread.authorize(creds)
        
        # スプレッドシートを開く
        spreadsheet = client.open_by_key(SPREADSHEET_ID)
        
        # シート名を取得
        sheet_names = [worksheet.title for worksheet in spreadsheet.worksheets()]
        print("利用可能なシート名:")
        for i, name in enumerate(sheet_names, 1):
            print(f"{i}. {name}")
        
        return sheet_names
        
    except Exception as e:
        print(f"エラー: {e}")
        return None

if __name__ == "__main__":
    check_sheets() 