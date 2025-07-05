# LINE Bot 見積書作成システム

## 概要
LINE Botを使用した見積書作成システムです。顧客が商品情報を入力すると、Googleスプレッドシートに自動反映されます。

## 機能
- 商品データの登録
- 顧客別スプレッドシート管理
- 利用制限機能
- Stripe決済連携
- リッチメニュー対応

## 本番環境デプロイ

### Renderでのデプロイ
1. Renderアカウントを作成
2. GitHubリポジトリと連携
3. 環境変数を設定
4. デプロイ実行

### 必要な環境変数
- `LINE_CHANNEL_ACCESS_TOKEN`: LINE Botのアクセストークン
- `LINE_CHANNEL_SECRET`: LINE Botのチャネルシークレット
- `SPREADSHEET_ID`: デフォルトのスプレッドシートID
- `SHEET_NAME`: シート名（デフォルト: "比較見積書 ロング"）
- `GOOGLE_SHEETS_CREDENTIALS`: Google Sheets API認証情報
- `STRIPE_SECRET_KEY`: Stripe秘密鍵
- `STRIPE_WEBHOOK_SECRET`: Stripe Webhook秘密鍵

## ローカル開発
```bash
python app.py
```

## ライセンス
MIT License 