import os
import json
import requests
import msal
from datetime import datetime
import re

class ExcelOnlineManager:
    def __init__(self):
        """Microsoft Excel Onlineとの連携を管理するクラス"""
        self.client_id = os.environ.get('MS_CLIENT_ID')
        self.client_secret = os.environ.get('MS_CLIENT_SECRET')
        self.tenant_id = os.environ.get('MS_TENANT_ID')
        self.authority = f"https://login.microsoftonline.com/{self.tenant_id}"
        self.scope = ["https://graph.microsoft.com/.default"]
        
    def get_access_token(self):
        """Microsoft Graph APIのアクセストークンを取得"""
        try:
            app = msal.ConfidentialClientApplication(
                self.client_id,
                authority=self.authority,
                client_credential=self.client_secret
            )
            
            result = app.acquire_token_for_client(scopes=self.scope)
            if "access_token" in result:
                return result["access_token"]
            else:
                print(f"トークン取得エラー: {result.get('error_description', 'Unknown error')}")
                return None
        except Exception as e:
            print(f"アクセストークン取得エラー: {e}")
            return None
    
    def extract_file_id_from_url(self, url):
        """SharePoint/OneDrive URLからファイルIDを抽出"""
        # SharePoint URLのパターンを検出
        patterns = [
            r'/personal/[^/]+/Documents/([^/?]+)',
            r'/sites/[^/]+/Shared%20Documents/([^/?]+)',
            r'/drives/[^/]+/items/([^/?]+)',
            r'/sites/[^/]+/lists/[^/]+/items/([^/?]+)'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, url)
            if match:
                return match.group(1)
        
        # URLから直接ファイル名を抽出
        if '/Documents/' in url:
            file_name = url.split('/Documents/')[-1].split('?')[0]
            return file_name
        
        return None
    
    def get_workbook(self, file_id):
        """Excelファイルの情報を取得"""
        try:
            access_token = self.get_access_token()
            if not access_token:
                return None, "アクセストークンの取得に失敗しました"
            
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            
            # ファイル情報を取得
            url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook"
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                return response.json(), None
            else:
                return None, f"ファイル取得エラー: {response.status_code} - {response.text}"
                
        except Exception as e:
            return None, f"ワークブック取得エラー: {e}"
    
    def get_worksheets(self, file_id):
        """Excelファイルのワークシート一覧を取得"""
        try:
            access_token = self.get_access_token()
            if not access_token:
                return None, "アクセストークンの取得に失敗しました"
            
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            
            url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets"
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                worksheets = response.json().get('value', [])
                return [ws['name'] for ws in worksheets], None
            else:
                return None, f"ワークシート取得エラー: {response.status_code} - {response.text}"
                
        except Exception as e:
            return None, f"ワークシート取得エラー: {e}"
    
    def read_range(self, file_id, sheet_name, range_address):
        """指定された範囲のデータを読み取り"""
        try:
            access_token = self.get_access_token()
            if not access_token:
                return None, "アクセストークンの取得に失敗しました"
            
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            
            # シート名をURLエンコード
            import urllib.parse
            encoded_sheet_name = urllib.parse.quote(sheet_name)
            encoded_range = urllib.parse.quote(range_address)
            
            url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets/{encoded_sheet_name}/range(address='{encoded_range}')"
            response = requests.get(url, headers=headers)
            
            if response.status_code == 200:
                data = response.json()
                return data.get('values', []), None
            else:
                return None, f"データ読み取りエラー: {response.status_code} - {response.text}"
                
        except Exception as e:
            return None, f"データ読み取りエラー: {e}"
    
    def write_range(self, file_id, sheet_name, range_address, values):
        """指定された範囲にデータを書き込み"""
        try:
            access_token = self.get_access_token()
            if not access_token:
                return False, "アクセストークンの取得に失敗しました"
            
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            
            # シート名をURLエンコード
            import urllib.parse
            encoded_sheet_name = urllib.parse.quote(sheet_name)
            encoded_range = urllib.parse.quote(range_address)
            
            url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets/{encoded_sheet_name}/range(address='{encoded_range}')"
            
            payload = {
                "values": values
            }
            
            response = requests.patch(url, headers=headers, json=payload)
            
            if response.status_code == 200:
                return True, None
            else:
                return False, f"データ書き込みエラー: {response.status_code} - {response.text}"
                
        except Exception as e:
            return False, f"データ書き込みエラー: {e}"
    
    def clear_range(self, file_id, sheet_name, range_address):
        """指定された範囲のデータをクリア"""
        try:
            access_token = self.get_access_token()
            if not access_token:
                return False, "アクセストークンの取得に失敗しました"
            
            headers = {
                'Authorization': f'Bearer {access_token}',
                'Content-Type': 'application/json'
            }
            
            # シート名をURLエンコード
            import urllib.parse
            encoded_sheet_name = urllib.parse.quote(sheet_name)
            encoded_range = urllib.parse.quote(range_address)
            
            url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets/{encoded_sheet_name}/range(address='{encoded_range}')/clear"
            
            response = requests.post(url, headers=headers)
            
            if response.status_code == 200:
                return True, None
            else:
                return False, f"データクリアエラー: {response.status_code} - {response.text}"
                
        except Exception as e:
            return False, f"データクリアエラー: {e}"
    
    def update_company_info_excel(self, data, file_id, sheet_name):
        """会社情報をExcelファイルに更新"""
        try:
            # 会社名を設定
            company_name = data.get('社名', '')
            if company_name:
                success, error = self.write_range(
                    file_id, 
                    sheet_name, 
                    'A2', 
                    [[company_name]]
                )
                if not success:
                    return False, f"会社名の更新に失敗: {error}"
            
            # 日付を設定
            current_date = datetime.now().strftime('%Y/%m/%d')
            success, error = self.write_range(
                file_id, 
                sheet_name, 
                'M2', 
                [[current_date]]
            )
            if not success:
                return False, f"日付の更新に失敗: {error}"
            
            return True, None
            
        except Exception as e:
            return False, f"会社情報更新エラー: {e}"
    
    def write_product_data_excel(self, data, file_id, sheet_name, row_number):
        """商品データをExcelファイルに書き込み"""
        try:
            product_name = data.get('商品名', '')
            unit_price = data.get('単価', '')
            quantity = data.get('数量', '')
            cycle = data.get('サイクル', '')
            
            # 商品データを配列として準備
            product_data = [
                [product_name, unit_price, quantity, '', '', cycle, '']
            ]
            
            # 行番号を指定して書き込み
            range_address = f'A{row_number}:G{row_number}'
            success, error = self.write_range(
                file_id, 
                sheet_name, 
                range_address, 
                product_data
            )
            
            if not success:
                return False, f"商品データの書き込みに失敗: {error}"
            
            return True, None
            
        except Exception as e:
            return False, f"商品データ書き込みエラー: {e}"
    
    def validate_excel_url(self, url):
        """Excel Online URLの妥当性を検証"""
        if not url:
            return False, "URLが空です"
        
        # SharePoint/OneDrive URLのパターンをチェック
        valid_patterns = [
            r'https://.*\.sharepoint\.com/.*\.xlsx',
            r'https://.*\.sharepoint\.com/.*\.xls',
            r'https://.*\.sharepoint\.com/.*\?rtime=',
            r'https://.*\.sharepoint\.com/.*/Documents/.*\.xlsx',
            r'https://.*\.sharepoint\.com/.*/Documents/.*\.xls'
        ]
        
        for pattern in valid_patterns:
            if re.search(pattern, url, re.IGNORECASE):
                return True, None
        
        return False, "有効なExcel Online URLではありません" 