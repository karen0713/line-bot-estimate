import stripe
import os
from datetime import datetime

class StripePayment:
    def __init__(self):
        # Stripe API設定
        self.stripe_secret_key = os.environ.get('STRIPE_SECRET_KEY', '')
        self.stripe_publishable_key = os.environ.get('STRIPE_PUBLISHABLE_KEY', '')
        
        if self.stripe_secret_key:
            stripe.api_key = self.stripe_secret_key
    
    def get_plan_info(self, plan_type):
        """プラン情報を取得"""
        plans = {
            'basic': {
                'name': 'ベーシックプラン',
                'price': 500,
                'description': '月100件まで利用可能',
                'features': ['月100件まで', '基本的な見積書機能', '利用履歴管理'],
                'stripe_price_id': os.environ.get('STRIPE_BASIC_PRICE_ID', '')
            },
            'pro': {
                'name': 'プロプラン',
                'price': 1000,
                'description': '無制限利用',
                'features': ['無制限利用', '全機能利用', '優先サポート', 'データ分析'],
                'stripe_price_id': os.environ.get('STRIPE_PRO_PRICE_ID', '')
            }
        }
        
        return plans.get(plan_type, None)
    
    def create_checkout_session(self, plan_type, user_id, success_url=None, cancel_url=None):
        """チェックアウトセッションを作成"""
        print(f"Creating checkout session for plan: {plan_type}, user: {user_id}")
        print(f"Stripe secret key available: {bool(self.stripe_secret_key)}")
        
        plan_info = self.get_plan_info(plan_type)
        if not plan_info:
            print(f"Plan not found: {plan_type}")
            return False, "プランが見つかりません"
        
        print(f"Plan info: {plan_info}")
        
        if not success_url:
            success_url = "https://line-bot-estimate.onrender.com/payment/success"
        if not cancel_url:
            cancel_url = "https://line-bot-estimate.onrender.com/payment/cancel"
        
        try:
            print("Creating Stripe checkout session...")
            checkout_session = stripe.checkout.Session.create(
                payment_method_types=['card'],
                line_items=[
                    {
                        'price': plan_info['stripe_price_id'],
                        'quantity': 1,
                    },
                ],
                mode='subscription',
                success_url=success_url + f"?user_id={user_id}&plan={plan_type}",
                cancel_url=cancel_url + f"?user_id={user_id}",
                metadata={
                    'user_id': user_id,
                    'plan_type': plan_type
                }
            )
            
            print(f"Checkout session created successfully: {checkout_session.id}")
            return True, {
                'checkout_url': checkout_session.url,
                'session_id': checkout_session.id,
                'plan_info': plan_info
            }
            
        except Exception as e:
            print(f"Stripe checkout session creation error: {str(e)}")
            return False, f"Stripe checkout session creation error: {str(e)}"
    
    def create_customer_portal_session(self, customer_id, return_url=None):
        """カスタマーポータルセッションを作成"""
        if not return_url:
            return_url = "https://line-bot-estimate.onrender.com/payment/portal_return"
        
        try:
            session = stripe.billing_portal.Session.create(
                customer=customer_id,
                return_url=return_url,
            )
            
            return True, {
                'portal_url': session.url
            }
            
        except Exception as e:
            return False, f"Customer portal session creation error: {str(e)}"
    
    def handle_webhook(self, payload, sig_header, webhook_secret):
        """Webhookイベントを処理"""
        try:
            event = stripe.Webhook.construct_event(
                payload, sig_header, webhook_secret
            )
            
            # イベントタイプに応じて処理
            if event['type'] == 'checkout.session.completed':
                return self.handle_checkout_completed(event['data']['object'])
            elif event['type'] == 'customer.subscription.created':
                return self.handle_subscription_created(event['data']['object'])
            elif event['type'] == 'customer.subscription.updated':
                return self.handle_subscription_updated(event['data']['object'])
            elif event['type'] == 'customer.subscription.deleted':
                return self.handle_subscription_deleted(event['data']['object'])
            
            return True, "Webhook processed successfully"
            
        except ValueError as e:
            return False, f"Invalid payload: {str(e)}"
        except stripe.error.SignatureVerificationError as e:
            return False, f"Invalid signature: {str(e)}"
        except Exception as e:
            return False, f"Webhook error: {str(e)}"
    
    def handle_checkout_completed(self, session):
        """チェックアウト完了時の処理"""
        user_id = session.metadata.get('user_id')
        plan_type = session.metadata.get('plan_type')
        
        # ここでユーザーのプランを更新
        # user_manager.upgrade_plan(user_id, plan_type)
        
        return True, f"Checkout completed for user {user_id}, plan {plan_type}"
    
    def handle_subscription_created(self, subscription):
        """サブスクリプション作成時の処理"""
        customer_id = subscription.customer
        # サブスクリプション情報を保存
        
        return True, f"Subscription created for customer {customer_id}"
    
    def handle_subscription_updated(self, subscription):
        """サブスクリプション更新時の処理"""
        customer_id = subscription.customer
        # サブスクリプション情報を更新
        
        return True, f"Subscription updated for customer {customer_id}"
    
    def handle_subscription_deleted(self, subscription):
        """サブスクリプション削除時の処理"""
        customer_id = subscription.customer
        # サブスクリプションを無効化
        
        return True, f"Subscription deleted for customer {customer_id}" 