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
    
    def clear_new_estimate_short_only(self, file_id, sheet_name):
        """新規見積書　ショート専用のリセット（B23:D23には一切触れない）"""
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
            
            print("新規見積書　ショートのリセットを開始します（B23:D23は保護されます）")
            
            # B23:D23のヘッダー行を事前にバックアップ
            print("B23:D23のヘッダー行をバックアップ中...")
            header_backup, error = self.read_range(file_id, sheet_name, 'B23:D23')
            if error:
                print(f"警告: ヘッダー行のバックアップに失敗: {error}")
                header_backup = None
            else:
                print(f"ヘッダー行をバックアップしました: {header_backup}")
            
            # B24:G30の範囲を個別セルで処理（B23:D23は除外）
            target_cells = []
            
            # B列、C列、D列の24行目から30行目（B28:D28も含む）
            for col in ['B', 'C', 'D']:
                for row in range(24, 31):  # 24行目から30行目
                    target_cells.append(f"{col}{row}")
            
            # E列、F列、G列の24行目から30行目
            for col in ['E', 'F', 'G']:
                for row in range(24, 31):  # 24行目から30行目
                    target_cells.append(f"{col}{row}")
            
            print(f"リセット対象セル数: {len(target_cells)}")
            print(f"リセット対象セル: {target_cells}")
            
            # 各セルを個別にクリア
            cleared_count = 0
            for cell_address in target_cells:
                encoded_cell = urllib.parse.quote(cell_address)
                url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets/{encoded_sheet_name}/range(address='{encoded_cell}')"
                
                payload = {
                    "values": [['']]
                }
                
                response = requests.patch(url, headers=headers, json=payload)
                
                if response.status_code != 200:
                    print(f"警告: セル {cell_address} のクリアに失敗: {response.status_code}")
                else:
                    print(f"クリア: セル {cell_address}")
                    cleared_count += 1
                
                # 各セルの処理後に少し待機
                import time
                time.sleep(0.1)
            
            # 処理完了後に少し待機
            import time
            time.sleep(0.5)
            
            # B23:D23のヘッダー行を復元
            if header_backup:
                print("B23:D23のヘッダー行を復元中...")
                success, error = self.write_range(file_id, sheet_name, 'B23:D23', header_backup)
                if success:
                    print("ヘッダー行の復元が完了しました")
                else:
                    print(f"警告: ヘッダー行の復元に失敗: {error}")
                    
                    # 復元に失敗した場合は、再度試行
                    print("ヘッダー行の復元を再試行中...")
                    time.sleep(1)
                    success, error = self.write_range(file_id, sheet_name, 'B23:D23', header_backup)
                    if success:
                        print("ヘッダー行の復元が完了しました（再試行成功）")
                    else:
                        print(f"警告: ヘッダー行の復元に再び失敗: {error}")
            
            print(f"リセット完了: {cleared_count}/{len(target_cells)} セルをクリアしました")
            return True, f"新規見積書　ショートのリセットが完了しました（{cleared_count}セルをクリア、B23:D23は保護）"
                
        except Exception as e:
            return False, f"新規見積書　ショートリセットエラー: {e}"

    def clear_range_safe_for_new_estimate_short(self, file_id, sheet_name):
        """新規見積書　ショート専用の安全なリセット（B23:D23を保護）"""
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
            
            # B23:D23のヘッダー行を事前にバックアップ
            print("B23:D23のヘッダー行をバックアップ中...")
            header_backup, error = self.read_range(file_id, sheet_name, 'B23:D23')
            if error:
                print(f"警告: ヘッダー行のバックアップに失敗: {error}")
                header_backup = None
            else:
                print(f"ヘッダー行をバックアップしました: {header_backup}")
            
            # 新規見積書　ショートのリセット範囲を個別に処理
            # 各列を個別にクリアして安全性を確保
            clear_columns = [
                ('B', 24, 30),  # B列 24行目から30行目
                ('C', 24, 30),  # C列 24行目から30行目
                ('D', 24, 30),  # D列 24行目から30行目
                ('E', 24, 30),  # E列 24行目から30行目
                ('F', 24, 30),  # F列 24行目から30行目
                ('G', 24, 30)   # G列 24行目から30行目
            ]
            
            # 各列を個別にクリア
            for col_letter, start_row, end_row in clear_columns:
                print(f"{col_letter}列 {start_row}行目から{end_row}行目をクリア中...")
                
                # 各セルを個別にクリア
                for row in range(start_row, end_row + 1):
                    cell_address = f"{col_letter}{row}"
                    
                    # B23:D23の範囲は絶対にクリアしない
                    if row == 23 and col_letter in ['B', 'C', 'D']:
                        print(f"保護: セル {cell_address} はヘッダー行のためスキップ")
                        continue
                    
                    encoded_cell = urllib.parse.quote(cell_address)
                    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets/{encoded_sheet_name}/range(address='{encoded_cell}')"
                    
                    payload = {
                        "values": [['']]
                    }
                    
                    response = requests.patch(url, headers=headers, json=payload)
                    
                    if response.status_code != 200:
                        print(f"警告: セル {cell_address} のクリアに失敗: {response.status_code}")
                    else:
                        print(f"クリア: セル {cell_address}")
                
                # 各列の処理後に少し待機
                import time
                time.sleep(0.3)
            
            # リセット処理完了後に少し待機
            import time
            time.sleep(1)
            
            # B23:D23のヘッダー行を復元
            if header_backup:
                print("B23:D23のヘッダー行を復元中...")
                success, error = self.write_range(file_id, sheet_name, 'B23:D23', header_backup)
                if success:
                    print("ヘッダー行の復元が完了しました")
                else:
                    print(f"警告: ヘッダー行の復元に失敗: {error}")
                    
                    # 復元に失敗した場合は、再度試行
                    print("ヘッダー行の復元を再試行中...")
                    time.sleep(2)
                    success, error = self.write_range(file_id, sheet_name, 'B23:D23', header_backup)
                    if success:
                        print("ヘッダー行の復元が完了しました（再試行成功）")
                    else:
                        print(f"警告: ヘッダー行の復元に再び失敗: {error}")
            
            return True, "新規見積書　ショートのリセットが完了しました（B23:D23は保護されました）"
                
        except Exception as e:
            return False, f"新規見積書　ショートリセットエラー: {e}"

    def clear_range(self, file_id, sheet_name, range_address):
        """指定された範囲のデータをクリア（個別セルで安全にクリア）"""
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
            
            # 範囲のサイズを計算
            import re
            range_match = re.match(r'([A-Z]+)(\d+):([A-Z]+)(\d+)', range_address)
            if not range_match:
                return False, "無効な範囲形式です"
            
            start_col = range_match.group(1)
            start_row = int(range_match.group(2))
            end_col = range_match.group(3)
            end_row = int(range_match.group(4))
            
            # 列を数値に変換
            def col_to_num(col):
                result = 0
                for char in col:
                    result = result * 26 + (ord(char) - ord('A') + 1)
                return result
            
            def num_to_col(num):
                result = ""
                while num > 0:
                    num -= 1
                    result = chr(num % 26 + ord('A')) + result
                    num //= 26
                return result
            
            start_col_num = col_to_num(start_col)
            end_col_num = col_to_num(end_col)
            
            # 各セルを個別にクリア（より安全な方法）
            for row in range(start_row, end_row + 1):
                for col_num in range(start_col_num, end_col_num + 1):
                    col_letter = num_to_col(col_num)
                    cell_address = f"{col_letter}{row}"
                    encoded_cell = urllib.parse.quote(cell_address)
                    
                    url = f"https://graph.microsoft.com/v1.0/me/drive/items/{file_id}/workbook/worksheets/{encoded_sheet_name}/range(address='{encoded_cell}')"
                    
                    payload = {
                        "values": [['']]
                    }
                    
                    response = requests.patch(url, headers=headers, json=payload)
                    
                    if response.status_code != 200:
                        print(f"警告: セル {cell_address} のクリアに失敗: {response.status_code}")
                        # 個別セルの失敗は警告として記録するが、処理は続行
            
            return True, None
                
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