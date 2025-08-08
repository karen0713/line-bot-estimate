# LINE Bot 見積書作成システム

## 概要
LINE Botを使用した見積書作成システムです。顧客が商品情報を入力すると、GoogleスプレッドシートまたはMicrosoft Excel Onlineに自動反映されます。

## 機能
- 商品データの登録
- 顧客別スプレッドシート管理
- Microsoft Excel Online連携
- 利用制限機能
- Stripe決済連携
- リッチメニュー対応

## Excel Online連携機能

### 対応ファイル形式
- SharePoint/OneDriveのExcel Onlineファイル（.xlsx, .xls）

### 使用方法
1. Excel OnlineファイルのURLを登録：
   ```
   Excel Online登録:https://unimatlifejp-my.sharepoint.com/... シート名:見積書
   ```

2. 登録確認：
   ```
   Excel Online確認
   ```

3. 商品データの自動反映：
   - 商品名、単価、数量、サイクルなどの情報が自動的にExcel Onlineファイルに反映されます
   - 会社情報の更新も対応しています

### 必要な設定
- Microsoft Graph APIのアプリケーション登録
- 適切な権限設定（Files.ReadWrite.All）
- 環境変数の設定

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
- `MS_CLIENT_ID`: Microsoft Graph APIクライアントID
- `MS_CLIENT_SECRET`: Microsoft Graph APIクライアントシークレット
- `MS_TENANT_ID`: MicrosoftテナントID
- `STRIPE_SECRET_KEY`: Stripe秘密鍵
- `STRIPE_WEBHOOK_SECRET`: Stripe Webhook秘密鍵

## ローカル開発
```bash
python app.py
```

## ライセンス
MIT License 