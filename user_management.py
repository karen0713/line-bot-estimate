import sqlite3
import os
from datetime import datetime, timedelta
import json

class UserManager:
    def __init__(self, db_path=None):
        if db_path is None:
            # Render環境では/tmpディレクトリを使用
            if os.environ.get('RENDER'):
                self.db_path = '/tmp/users.db'
            else:
                self.db_path = 'users.db'
        else:
            self.db_path = db_path
        self.init_database()
    
    def init_database(self):
        """データベースの初期化"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            
            # ユーザーテーブルの作成（スプレッドシート管理機能を追加）
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS users (
                    user_id TEXT PRIMARY KEY,
                    display_name TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    plan_type TEXT DEFAULT 'free',
                    monthly_usage INTEGER DEFAULT 0,
                    last_reset_date DATE DEFAULT CURRENT_DATE,
                    is_active BOOLEAN DEFAULT 1,
                    spreadsheet_id TEXT,
                    sheet_name TEXT DEFAULT '比較見積書 ロング',
                    excel_online_url TEXT,
                    excel_file_id TEXT,
                    excel_sheet_name TEXT
                )
            ''')
            
            # 既存のテーブルにカラムが存在しない場合は追加
            try:
                cursor.execute('ALTER TABLE users ADD COLUMN spreadsheet_id TEXT')
            except sqlite3.OperationalError:
                pass  # カラムが既に存在する場合
            
            try:
                cursor.execute('ALTER TABLE users ADD COLUMN sheet_name TEXT DEFAULT "比較見積書 ロング"')
            except sqlite3.OperationalError:
                pass  # カラムが既に存在する場合
            
            try:
                cursor.execute('ALTER TABLE users ADD COLUMN excel_online_url TEXT')
            except sqlite3.OperationalError:
                pass  # カラムが既に存在する場合
            
            try:
                cursor.execute('ALTER TABLE users ADD COLUMN excel_file_id TEXT')
            except sqlite3.OperationalError:
                pass  # カラムが既に存在する場合
            
            try:
                cursor.execute('ALTER TABLE users ADD COLUMN excel_sheet_name TEXT')
            except sqlite3.OperationalError:
                pass  # カラムが既に存在する場合
            
            # 利用履歴テーブルの作成
            cursor.execute('''
                CREATE TABLE IF NOT EXISTS usage_history (
                    id INTEGER PRIMARY KEY AUTOINCREMENT,
                    user_id TEXT,
                    action_type TEXT,
                    action_data TEXT,
                    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP,
                    FOREIGN KEY (user_id) REFERENCES users (user_id)
                )
            ''')
            
            conn.commit()
            conn.close()
            print(f"Database initialized successfully at {self.db_path}")
        except Exception as e:
            print(f"Database initialization error: {e}")
            # エラーが発生してもアプリケーションは継続
            pass
    
    def register_user(self, user_id, display_name):
        """新規ユーザー登録"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                INSERT OR IGNORE INTO users (user_id, display_name)
                VALUES (?, ?)
            ''', (user_id, display_name))
            
            conn.commit()
            return True, "ユーザー登録完了"
        except Exception as e:
            return False, f"登録エラー: {str(e)}"
        finally:
            conn.close()
    
    def get_user_info(self, user_id):
        """ユーザー情報取得"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT user_id, display_name, plan_type, monthly_usage, last_reset_date, is_active
            FROM users WHERE user_id = ?
        ''', (user_id,))
        
        result = cursor.fetchone()
        conn.close()
        
        if result:
            return {
                'user_id': result[0],
                'display_name': result[1],
                'plan_type': result[2],
                'monthly_usage': result[3],
                'last_reset_date': result[4],
                'is_active': result[5]
            }
        return None
    
    def check_usage_limit(self, user_id):
        """利用制限チェック（開発者用：一時的に無効化）"""
        # 開発者用：利用制限を一時的に無効化
        return True, "開発者モード：利用制限なし"
        
        # 元のコード（コメントアウト）
        # user_info = self.get_user_info(user_id)
        # if not user_info:
        #     return False, "ユーザーが見つかりません"
        # 
        # # 月次リセット処理
        # self.reset_monthly_usage_if_needed(user_id)
        # 
        # # 利用制限チェック
        # if user_info['plan_type'] == 'free':
        #     limit = 10
        # elif user_info['plan_type'] == 'basic':
        #     limit = 100
        # else:  # pro
        #     limit = 999999
        # 
        # current_usage = self.get_current_monthly_usage(user_id)
        # 
        # if current_usage >= limit:
        #     return False, f"利用制限に達しました（月{limit}件まで）"
        # 
        # return True, f"利用可能（残り{limit - current_usage}件）"
    
    def reset_monthly_usage_if_needed(self, user_id):
        """月次利用回数のリセット"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT last_reset_date FROM users WHERE user_id = ?
        ''', (user_id,))
        
        result = cursor.fetchone()
        if result:
            last_reset = datetime.strptime(result[0], '%Y-%m-%d')
            current_date = datetime.now().date()
            
            # 月が変わったらリセット
            if last_reset.month != current_date.month or last_reset.year != current_date.year:
                cursor.execute('''
                    UPDATE users 
                    SET monthly_usage = 0, last_reset_date = ?
                    WHERE user_id = ?
                ''', (current_date, user_id))
                
                conn.commit()
        
        conn.close()
    
    def get_current_monthly_usage(self, user_id):
        """現在の月次利用回数を取得"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        cursor.execute('''
            SELECT monthly_usage FROM users WHERE user_id = ?
        ''', (user_id,))
        
        result = cursor.fetchone()
        conn.close()
        
        return result[0] if result else 0
    
    def increment_usage(self, user_id, action_type, action_data):
        """利用回数を増加"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            # 利用回数を増加
            cursor.execute('''
                UPDATE users 
                SET monthly_usage = monthly_usage + 1
                WHERE user_id = ?
            ''', (user_id,))
            
            # 利用履歴を記録
            cursor.execute('''
                INSERT INTO usage_history (user_id, action_type, action_data)
                VALUES (?, ?, ?)
            ''', (user_id, action_type, json.dumps(action_data, ensure_ascii=False)))
            
            conn.commit()
            return True, "利用回数を記録しました"
        except Exception as e:
            return False, f"記録エラー: {str(e)}"
        finally:
            conn.close()
    
    def get_usage_summary(self, user_id):
        """利用状況サマリー取得"""
        user_info = self.get_user_info(user_id)
        if not user_info:
            return "ユーザーが見つかりません"
        
        current_usage = self.get_current_monthly_usage(user_id)
        
        if user_info['plan_type'] == 'free':
            limit = 10
            plan_name = "無料プラン"
        elif user_info['plan_type'] == 'basic':
            limit = 100
            plan_name = "ベーシックプラン"
        else:
            limit = 999999
            plan_name = "プロプラン"
        
        remaining = max(0, limit - current_usage)
        
        summary = f"📊 利用状況\n\n"
        summary += f"プラン: {plan_name}\n"
        summary += f"今月の利用回数: {current_usage}回\n"
        summary += f"残り利用回数: {remaining}回\n"
        summary += f"リセット日: {user_info['last_reset_date']}"
        
        return summary
    
    def upgrade_plan(self, user_id, plan_type):
        """プランアップグレード"""
        conn = sqlite3.connect(self.db_path)
        cursor = conn.cursor()
        
        try:
            cursor.execute('''
                UPDATE users 
                SET plan_type = ?
                WHERE user_id = ?
            ''', (plan_type, user_id))
            
            conn.commit()
            return True
        except Exception as e:
            print(f"Plan upgrade error: {e}")
            return False
        finally:
            conn.close()

    def set_user_spreadsheet(self, user_id, spreadsheet_id, sheet_name="比較見積書 ロング"):
        """顧客のスプレッドシートIDを設定"""
        print(f"set_user_spreadsheet: user_id={user_id}, spreadsheet_id={spreadsheet_id}, sheet_name={sheet_name}")
        try:
            # --- シート名を正規化 ---
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE users 
                SET spreadsheet_id = ?, sheet_name = ?
                WHERE user_id = ?
            ''', (spreadsheet_id, sheet_name, user_id))
            conn.commit()
            conn.close()
            return True, "スプレッドシートを登録しました"
        except Exception as e:
            return False, f"登録エラー: {str(e)}"

    def get_user_spreadsheet(self, user_id):
        """顧客のスプレッドシートIDを取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT spreadsheet_id, sheet_name FROM users WHERE user_id = ?', (user_id,))
            result = cursor.fetchone()
            conn.close()
            if result:
                print(f"get_user_spreadsheet: user_id={user_id}, spreadsheet_id={result[0]}, sheet_name={result[1]}")
                # --- シート名を正規化 ---
                return result[0], result[1]
            return (None, None)
        except Exception as e:
            return None, None

    def set_user_excel_online(self, user_id, excel_url, file_id, sheet_name="Sheet1"):
        """顧客のExcel Online設定を保存"""
        print(f"set_user_excel_online: user_id={user_id}, excel_url={excel_url}, file_id={file_id}, sheet_name={sheet_name}")
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('''
                UPDATE users 
                SET excel_online_url = ?, excel_file_id = ?, excel_sheet_name = ?
                WHERE user_id = ?
            ''', (excel_url, file_id, sheet_name, user_id))
            conn.commit()
            conn.close()
            return True, "Excel Online設定を登録しました"
        except Exception as e:
            return False, f"登録エラー: {str(e)}"

    def get_user_excel_online(self, user_id):
        """顧客のExcel Online設定を取得"""
        try:
            conn = sqlite3.connect(self.db_path)
            cursor = conn.cursor()
            cursor.execute('SELECT excel_online_url, excel_file_id, excel_sheet_name FROM users WHERE user_id = ?', (user_id,))
            result = cursor.fetchone()
            conn.close()
            if result:
                print(f"get_user_excel_online: user_id={user_id}, excel_url={result[0]}, file_id={result[1]}, sheet_name={result[2]}")
                return result[0], result[1], result[2]
            return (None, None, None)
        except Exception as e:
            return None, None, None 