# Stripe決済システム設定手順

## 1. Stripeアカウント作成
1. [Stripe公式サイト](https://stripe.com)にアクセス
2. アカウントを作成（無料）
3. ダッシュボードにログイン

## 2. APIキーの取得
1. ダッシュボードの「開発者」→「APIキー」に移動
2. 以下のキーをコピー：
   - **秘密鍵**（sk_test_...）
   - **公開鍵**（pk_test_...）

## 3. 商品・価格の作成
### ベーシックプラン
1. 「商品」→「商品を追加」
2. 商品名：「ベーシックプラン」
3. 価格：500円（月額）
4. 価格IDをコピー（price_...）

### プロプラン
1. 「商品」→「商品を追加」
2. 商品名：「プロプラン」
3. 価格：1000円（月額）
4. 価格IDをコピー（price_...）

## 4. Webhook設定
1. 「開発者」→「Webhook」に移動
2. 「エンドポイントを追加」
3. エンドポイントURL：`https://your-domain.com/stripe/webhook`
4. イベントを選択：
   - `checkout.session.completed`
   - `customer.subscription.created`
   - `customer.subscription.updated`
   - `customer.subscription.deleted`
5. Webhook秘密鍵をコピー（whsec_...）

## 5. 環境変数の設定
Renderダッシュボードで以下の環境変数を設定：

```
STRIPE_SECRET_KEY=sk_test_...
STRIPE_PUBLISHABLE_KEY=pk_test_...
STRIPE_WEBHOOK_SECRET=whsec_...
STRIPE_BASIC_PRICE_ID=price_...
STRIPE_PRO_PRICE_ID=price_...
```

## 6. テスト決済
1. テスト用カード番号：`4242 4242 4242 4242`
2. 有効期限：任意の将来の日付
3. CVC：任意の3桁の数字

## 7. 本番環境への移行
1. ダッシュボードで「本番モード」に切り替え
2. 本番用のAPIキーを取得
3. 環境変数を本番用に更新
4. Webhookエンドポイントを本番用に更新

## 注意事項
- テストモードでは実際の課金は発生しません
- 本番環境では適切なセキュリティ設定が必要です
- Webhookの署名検証は自動で行われます 