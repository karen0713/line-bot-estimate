import requests
import json
import os
import hashlib
import hmac
import time
from datetime import datetime

class LinePayPayment:
    def __init__(self):
        # LINE Pay API設定
        self.channel_id = os.environ.get('LINE_PAY_CHANNEL_ID', '')
        self.channel_secret = os.environ.get('LINE_PAY_CHANNEL_SECRET', '')
        self.is_sandbox = os.environ.get('LINE_PAY_SANDBOX', 'true').lower() == 'true'
        
        if self.is_sandbox:
            self.api_url = 'https://sandbox-api-pay.line.me'
        else:
            self.api_url = 'https://api-pay.line.me'
    
    def create_payment_request(self, amount, currency='JPY', order_id=None, product_name='プランアップグレード'):
        """決済リクエストを作成"""
        if not order_id:
            order_id = f"order_{int(time.time())}"
        
        # リクエストボディ
        body = {
            "amount": amount,
            "currency": currency,
            "orderId": order_id,
            "packages": [
                {
                    "id": "package-1",
                    "amount": amount,
                    "products": [
                        {
                            "name": product_name,
                            "quantity": 1,
                            "price": amount
                        }
                    ]
                }
            ],
            "redirectUrls": {
                "confirmUrl": f"https://line-bot-estimate.onrender.com/payment/confirm",
                "cancelUrl": f"https://line-bot-estimate.onrender.com/payment/cancel"
            }
        }
        
        # ヘッダー
        headers = {
            'Content-Type': 'application/json',
            'X-LINE-ChannelId': self.channel_id,
            'X-LINE-Authorization-Nonce': str(int(time.time())),
            'X-LINE-Authorization': self._generate_signature(body)
        }
        
        try:
            response = requests.post(
                f"{self.api_url}/v3/payments/request",
                headers=headers,
                json=body
            )
            
            if response.status_code == 200:
                result = response.json()
                return True, result
            else:
                return False, f"LINE Pay API error: {response.status_code} - {response.text}"
                
        except Exception as e:
            return False, f"Payment request error: {str(e)}"
    
    def confirm_payment(self, transaction_id, amount, currency='JPY'):
        """決済を確定"""
        body = {
            "amount": amount,
            "currency": currency
        }
        
        headers = {
            'Content-Type': 'application/json',
            'X-LINE-ChannelId': self.channel_id,
            'X-LINE-Authorization-Nonce': str(int(time.time())),
            'X-LINE-Authorization': self._generate_signature(body)
        }
        
        try:
            response = requests.post(
                f"{self.api_url}/v3/payments/{transaction_id}/confirm",
                headers=headers,
                json=body
            )
            
            if response.status_code == 200:
                result = response.json()
                return True, result
            else:
                return False, f"LINE Pay confirmation error: {response.status_code} - {response.text}"
                
        except Exception as e:
            return False, f"Payment confirmation error: {str(e)}"
    
    def _generate_signature(self, body):
        """署名を生成"""
        body_str = json.dumps(body, separators=(',', ':'))
        nonce = str(int(time.time()))
        
        # 署名文字列を作成
        signature_string = self.channel_secret + '/v3/payments/request' + body_str + nonce
        
        # HMAC-SHA256で署名を生成
        signature = hmac.new(
            self.channel_secret.encode('utf-8'),
            signature_string.encode('utf-8'),
            hashlib.sha256
        ).hexdigest()
        
        return signature
    
    def get_plan_info(self, plan_type):
        """プラン情報を取得"""
        plans = {
            'basic': {
                'name': 'ベーシックプラン',
                'price': 500,
                'description': '月100件まで利用可能',
                'features': ['月100件まで', '基本的な見積書機能', '利用履歴管理']
            },
            'pro': {
                'name': 'プロプラン',
                'price': 1000,
                'description': '無制限利用',
                'features': ['無制限利用', '全機能利用', '優先サポート', 'データ分析']
            }
        }
        
        return plans.get(plan_type, None)
    
    def create_payment_url(self, plan_type, user_id):
        """決済URLを作成"""
        plan_info = self.get_plan_info(plan_type)
        if not plan_info:
            return False, "プランが見つかりません"
        
        order_id = f"upgrade_{plan_type}_{user_id}_{int(time.time())}"
        
        success, result = self.create_payment_request(
            amount=plan_info['price'],
            order_id=order_id,
            product_name=plan_info['name']
        )
        
        if success:
            payment_url = result.get('info', {}).get('paymentUrl', {}).get('web', '')
            return True, {
                'payment_url': payment_url,
                'transaction_id': result.get('info', {}).get('transactionId', ''),
                'order_id': order_id,
                'plan_info': plan_info
            }
        else:
            return False, result 