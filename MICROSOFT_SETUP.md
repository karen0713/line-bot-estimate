# Microsoft Graph API設定手順

## 概要
このシステムでMicrosoft Excel Onlineとの連携を有効にするための設定手順です。

## 1. Azure Active Directoryでアプリケーションを登録

### 1.1 Azure Portalにアクセス
1. [Azure Portal](https://portal.azure.com)にアクセス
2. Azure Active Directoryに移動

### 1.2 アプリケーションの登録
1. 「アプリの登録」→「新規登録」をクリック
2. 以下の情報を入力：
   - **名前**: LINE Bot Excel Online連携
   - **サポートされているアカウントの種類**: この組織ディレクトリのみ
   - **リダイレクトURI**: Web → `https://your-app-name.onrender.com/`

### 1.3 アプリケーション情報を記録
登録後、以下の情報を記録してください：
- **アプリケーション（クライアント）ID**
- **ディレクトリ（テナント）ID**

## 2. クライアントシークレットの作成

### 2.1 シークレットの作成
1. 「証明書とシークレット」→「新しいクライアントシークレット」をクリック
2. 説明を入力（例：「LINE Bot用」）
3. 有効期限を選択（推奨：24ヶ月）
4. 「追加」をクリック

### 2.2 シークレット値を記録
作成後、**シークレットの値**を必ず記録してください（後で表示されません）

## 3. API権限の設定

### 3.1 権限の追加
1. 「APIのアクセス許可」→「アクセス許可の追加」をクリック
2. 「Microsoft Graph」を選択
3. 「アプリケーションのアクセス許可」を選択
4. 以下の権限を追加：
   - `Files.ReadWrite.All` - ファイルの読み取りと書き込み
   - `Sites.ReadWrite.All` - サイトの読み取りと書き込み

### 3.2 管理者の同意
1. 「管理者の同意を与える」をクリック
2. 確認ダイアログで「はい」をクリック

## 4. 環境変数の設定

### 4.1 本番環境（Render）での設定
Renderのダッシュボードで以下の環境変数を設定：

```
MS_CLIENT_ID=your_application_client_id
MS_CLIENT_SECRET=your_client_secret_value
MS_TENANT_ID=your_tenant_id
```

### 4.2 ローカル開発での設定
`.env`ファイルまたは環境変数として設定：

```bash
export MS_CLIENT_ID=your_application_client_id
export MS_CLIENT_SECRET=your_client_secret_value
export MS_TENANT_ID=your_tenant_id
```

## 5. Excel Onlineファイルの準備

### 5.1 ファイルの共有設定
1. SharePoint/OneDriveでExcelファイルを作成またはアップロード
2. ファイルを右クリック→「共有」
3. 「編集者」権限で共有設定
4. 共有リンクをコピー

### 5.2 ファイルの構造
推奨するExcelファイルの構造：
- A列: 商品名
- B列: 単価
- C列: 数量
- D列: サイクル
- 19行目以降: 商品データの開始行

## 6. 使用方法

### 6.1 Excel Onlineファイルの登録
LINE Botで以下のコマンドを送信：
```
Excel Online登録:https://your-sharepoint-url/... シート名:Sheet1
```

### 6.2 登録確認
```
Excel Online確認
```

### 6.3 商品データの入力
従来通り商品情報を入力すると、自動的にExcel Onlineファイルに反映されます。

## 7. トラブルシューティング

### 7.1 よくある問題
- **認証エラー**: クライアントシークレットが正しく設定されているか確認
- **権限エラー**: API権限が正しく設定されているか確認
- **ファイルアクセスエラー**: ファイルの共有設定を確認

### 7.2 ログの確認
アプリケーションのログで詳細なエラー情報を確認できます。

## 8. セキュリティに関する注意事項

- クライアントシークレットは絶対に公開しないでください
- 必要最小限の権限のみを付与してください
- 定期的にシークレットを更新してください
- アクセスログを定期的に確認してください 