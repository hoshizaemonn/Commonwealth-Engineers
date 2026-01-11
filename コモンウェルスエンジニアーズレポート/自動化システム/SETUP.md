# セットアップガイド

このガイドでは、GA4とSearch Consoleから自動的にデータを取得するための設定方法を説明します。

---

## 📋 必要なもの

1. **Python 3.7以上** がインストールされていること
2. **Google Cloud Platform (GCP) アカウント**
3. **credentials.json** ファイル（Google APIの認証情報）

---

## 🔑 Step 1: credentials.json を用意する

### credentials.json とは？

GoogleのAPI（GA4やSearch Console）に接続するための「合鍵」のようなものです。
このファイルがないと、プログラムはAPIにアクセスできません。

### 取得方法

1. [Google Cloud Console](https://console.cloud.google.com/) にアクセス
2. 新しいプロジェクトを作成（または既存のプロジェクトを選択）
3. 「APIとサービス」→「ライブラリ」に移動
4. 以下のAPIを有効化：
   - **Google Analytics Data API**
   - **Google Search Console API**
5. 「APIとサービス」→「認証情報」→「認証情報を作成」→「サービスアカウント」
6. サービスアカウントを作成し、JSONキーをダウンロード
7. ダウンロードしたファイルを `credentials.json` にリネーム
8. `credentials.json` をこのフォルダに配置

### ⚠️ 重要：権限の設定

#### GA4の権限設定
1. [Google Analytics](https://analytics.google.com/) にアクセス
2. 左下の「管理」をクリック
3. 「プロパティユーザー管理」をクリック
4. 「+」ボタンをクリックしてユーザーを追加
5. サービスアカウントのメールアドレス（credentials.json内の `client_email`）を入力
6. 「表示者」以上の権限を付与

#### Search Consoleの権限設定
1. [Search Console](https://search.google.com/search-console/) にアクセス
2. プロパティ（サイト）を選択
3. 左メニュー「設定」→「ユーザーと権限」をクリック
4. 「ユーザーを追加」をクリック
5. サービスアカウントのメールアドレス（credentials.json内の `client_email`）を入力
6. 「フル」権限を付与

---

## ⚙️ Step 2: config.json を作成する

### 2-1. テンプレートをコピー

まず、テンプレートファイルをコピーして `config.json` を作成します：

```bash
cp config.json.template config.json
```

### 2-2. GA4のプロパティIDを調べる

**プロパティID** とは、GA4の各サイトに割り当てられた数字のことです。

#### 見つけ方（超簡単！）

1. [Google Analytics](https://analytics.google.com/) にアクセス
2. 左側のメニューから **「管理」** をクリック（⚙️マーク）
3. 「プロパティ」列（真ん中の列）を見る
4. **「プロパティ設定」** をクリック
5. 画面に表示されている **「プロパティID」** をコピー（数字だけ、例：`123456789`）

#### config.json に設定

`config.json` をテキストエディタで開いて、次の部分を書き換えます：

```json
{
  "ga4": {
    "property_id": "123456789",  ← ここにプロパティIDを貼り付け
    "credentials_file": "credentials.json"
  },
  ...
}
```

**例：**
- `"property_id": "YOUR_GA4_PROPERTY_ID"` → `"property_id": "123456789"`

---

### 2-3. Search ConsoleのサイトURLを調べる

**サイトURL** とは、Search Consoleで管理しているサイトのアドレスです。

#### 見つけ方（超簡単！）

1. [Search Console](https://search.google.com/search-console/) にアクセス
2. 左側のメニューからプロパティ（サイト）を選択
3. 画面上部に表示されているサイトのURLをコピー

#### 形式の確認

Search Consoleには2種類の形式があります：

- **ドメインプロパティ**: `sc-domain:cectokyo.com`（ドメイン全体）
- **URLプレフィックスプロパティ**: `https://cectokyo.com/`（特定のURL）

どちらを使っているか確認して、そのままコピーしてください。

#### config.json に設定

`config.json` をテキストエディタで開いて、次の部分を書き換えます：

```json
{
  ...
  "search_console": {
    "site_url": "sc-domain:cectokyo.com",  ← ここにサイトURLを貼り付け
    "credentials_file": "credentials.json"
  },
  ...
}
```

**例：**
- `"site_url": "sc-domain:YOUR_SITE"` → `"site_url": "sc-domain:cectokyo.com"`
- または `"site_url": "https://cectokyo.com/"`

---

## 📦 Step 3: Pythonライブラリをインストールする

必要なライブラリをインストールします。

ターミナル（コマンドプロンプト）を開いて、このフォルダに移動してから：

```bash
pip install -r requirements.txt
```

**Mac/Linuxの場合：**
```bash
pip3 install -r requirements.txt
```

もしエラーが出た場合は：
```bash
python3 -m pip install -r requirements.txt
```

---

## 🧪 Step 4: 接続テストを実行する

設定が正しいか確認するために、まず接続テストを実行します：

```bash
python test_connection.py
```

**Mac/Linuxの場合：**
```bash
python3 test_connection.py
```

### 期待される結果

```
🚀 GA4 & Search Console 接続テスト開始
============================================================
🔍 GA4接続テスト
============================================================
✅ プロパティID: 123456789
✅ 認証情報ファイル: credentials.json
📡 APIに接続中...
✅ GA4への接続に成功しました！
   テストデータ取得: XX 行

🔍 Search Console接続テスト
============================================================
✅ サイトURL: sc-domain:cectokyo.com
✅ 認証情報ファイル: credentials.json
📡 APIに接続中...
✅ Search Consoleへの接続に成功しました！
   サイト権限: フル

============================================================
📊 テスト結果サマリー
============================================================
GA4接続: ✅ 成功
Search Console接続: ✅ 成功

🎉 すべての接続テストに成功しました！
   次は fetch_data.py を実行してデータを取得できます
```

### エラーが出た場合

- **「credentials.json が見つかりません」**  
  → `credentials.json` がこのフォルダにあるか確認してください

- **「プロパティIDが設定されていません」**  
  → `config.json` の `property_id` を確認してください

- **「サイトURLが設定されていません」**  
  → `config.json` の `site_url` を確認してください

- **「権限がありません」**  
  → Step 1の「権限の設定」を確認してください

---

## 📊 Step 5: データを取得する

接続テストが成功したら、実際にデータを取得します：

```bash
python fetch_data.py
```

**Mac/Linuxの場合：**
```bash
python3 fetch_data.py
```

### 実行結果

プログラムは自動的に：
- 今月の1日から最終日までのデータを取得
- `data/2025-XX/` フォルダにCSVファイルとして保存

### 取得されるファイル

- `トラフィック獲得_セッションのメインのチャネル_グループ（デフォルト_チャネル_グループ）.csv`
- `イベント_イベント名.csv`
- `ページとスクリーン_ページパスとスクリーン_クラス.csv`
- `クエリ.csv`
- `ページ.csv`
- `デバイス.csv`

---

## 📝 よくある質問（FAQ）

### Q1: エラーが出るけど、何が間違っているの？

**A:** 以下の順番で確認してください：
1. `credentials.json` が正しい場所にあるか
2. `config.json` の `property_id` と `site_url` が正しく設定されているか
3. Google Analytics と Search Console でサービスアカウントに権限が付与されているか

### Q2: プロパティIDってどこにあるの？

**A:** 
1. Google Analytics → 管理 → プロパティ設定
2. 画面上部に「プロパティID：123456789」と表示されています

### Q3: サイトURLはどっちを使えばいいの？

**A:** Search Consoleに表示されているURLをそのまま使ってください。
- `sc-domain:...` で始まるもの → そのまま
- `https://...` で始まるもの → そのまま

### Q4: データが取得できない！

**A:** 
- 接続テスト（`test_connection.py`）を先に実行してください
- エラーメッセージを確認してください
- 権限が正しく設定されているか確認してください

### Q5: 過去の月のデータを取得したい

**A:** `fetch_data.py` の `main()` 関数を編集して、日付を変更してください：

```python
# 例：2025年10月のデータを取得する場合
start_date = "2025-10-01"
end_date = "2025-10-31"
```

---

## 🎯 次のステップ

データが正常に取得できたら：

1. 取得したCSVファイルを確認
2. `templates/monthly-report.md` を使ってレポートを作成
3. `knowledge/improvement-patterns.md` を参考に分析

---

## 📚 参考資料

- [GA4 データ取得ガイド](knowledge/ga4-guide.md)
- [Search Console データ取得ガイド](knowledge/search-console-guide.md)
- [改善パターン集](knowledge/improvement-patterns.md)

---

**困ったときは、エラーメッセージをよく読んでください。**
**ほとんどの場合は、設定ファイル（config.json）や権限の問題です！**

