#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GA4とSearch Consoleからデータを自動取得するプログラム

このプログラムは、設定ファイル（config.json）の情報を使って
GA4とSearch Consoleから最新のデータを取得し、CSVファイルとして保存します。
"""

import json
import os
import csv
import time
import urllib.parse
import requests
import jwt
from datetime import datetime, timedelta

# ============================================================================
# データ取得期間の設定
# ============================================================================
# 取得するデータの期間を指定します
# 月を変更する場合は、この2つの値を変更してください
START_DATE = "2025-12-01"
END_DATE = "2025-12-31"
# ============================================================================


def load_config():
    """設定ファイルを読み込む"""
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    
    if not os.path.exists(config_path):
        print("❌ エラー: config.json が見つかりません")
        return None
    
    with open(config_path, 'r', encoding='utf-8') as f:
        config = json.load(f)
    
    return config


def get_access_token(credentials_file, scopes):
    """
    Service Account認証情報を使ってアクセストークンを取得
    
    Args:
        credentials_file: credentials.jsonのパス
        scopes: スコープのリスト
    
    Returns:
        access_token: アクセストークン
    """
    creds_path = os.path.join(os.path.dirname(__file__), credentials_file)
    
    with open(creds_path, 'r', encoding='utf-8') as f:
        creds = json.load(f)
    
    # JWT claimsを作成
    now = int(time.time())
    jwt_claims = {
        'iss': creds['client_email'],
        'sub': creds['client_email'],
        'aud': 'https://oauth2.googleapis.com/token',
        'iat': now,
        'exp': now + 3600,  # 1時間有効
        'scope': ' '.join(scopes)
    }
    
    # JWTを署名
    private_key = creds['private_key']
    signed_jwt = jwt.encode(jwt_claims, private_key, algorithm='RS256')
    
    # トークンリクエスト
    token_url = 'https://oauth2.googleapis.com/token'
    token_data = {
        'grant_type': 'urn:ietf:params:oauth:grant-type:jwt-bearer',
        'assertion': signed_jwt
    }
    
    response = requests.post(token_url, data=token_data)
    response.raise_for_status()
    
    token_info = response.json()
    return token_info['access_token']


def fetch_ga4_report(access_token, property_id, request_body, description):
    """
    GA4 APIからレポートを取得
    
    Args:
        access_token: アクセストークン
        property_id: GA4プロパティID
        request_body: リクエストボディ
        description: 説明文
    
    Returns:
        APIレスポンスのJSON
    """
    api_url = f"https://analyticsdata.googleapis.com/v1beta/properties/{property_id}:runReport"
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    response = requests.post(api_url, headers=headers, json=request_body)
    response.raise_for_status()
    
    return response.json()


def fetch_ga4_traffic_acquisition(access_token, property_id, start_date, end_date, output_dir):
    """
    GA4からトラフィック獲得（チャネル別）データを取得
    
    ガイド: knowledge/ga4-guide.md 参照
    """
    print("\n📊 GA4: トラフィック獲得データを取得中...")
    
    request_body = {
        "dateRanges": [{
            "startDate": start_date,
            "endDate": end_date
        }],
        "metrics": [
            {"name": "sessions"},
            {"name": "totalUsers"},
            {"name": "newUsers"},
            {"name": "eventCount"}
        ],
        "dimensions": [
            {"name": "sessionDefaultChannelGroup"},
            {"name": "eventName"}
        ],
        "orderBys": [
            {
                "dimension": {
                    "dimensionName": "sessionDefaultChannelGroup"
                }
            },
            {
                "metric": {
                    "metricName": "sessions"
                },
                "desc": True
            }
        ]
    }
    
    try:
        response = fetch_ga4_report(access_token, property_id, request_body, "トラフィック獲得")
        
        # CSVファイルに保存
        filename = os.path.join(output_dir, "トラフィック獲得_セッションのメインのチャネル_グループ（デフォルト_チャネル_グループ）.csv")
        
        with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            
            # ヘッダー（GA4のエクスポート形式に合わせる）
            writer.writerow([
                "# ----------------------------------------",
                "# トラフィック獲得: セッションのメインのチャネル グループ（デフォルト チャネル グループ）",
                "# アカウント: 株式会社 コモンウェルスエンジニアーズ",
                "# プロパティ: cectokyo.com",
                "# ----------------------------------------"
            ])
            writer.writerow(["# "])
            writer.writerow(["# すべてのユーザー"])
            writer.writerow([f"# 開始日: {start_date.replace('-', '')}"])
            writer.writerow([f"# 終了日: {end_date.replace('-', '')}"])
            
            # データヘッダー
            writer.writerow([
                "セッションのメインのチャネル グループ（デフォルト チャネル グループ）",
                "イベント名",
                "セッション",
                "新規ユーザー数",
                "セッションあたりの平均エンゲージメント時間",
                "セッションあたりのイベント数",
                "イベント数"
            ])
            
            # データ行
            rows = response.get('rows', [])
            for row in rows:
                dimension_values = row.get('dimensionValues', [])
                metric_values = row.get('metricValues', [])
                
                channel = dimension_values[0].get('value', '') if len(dimension_values) > 0 else ''
                event = dimension_values[1].get('value', '') if len(dimension_values) > 1 else ''
                sessions = metric_values[0].get('value', '0') if len(metric_values) > 0 else '0'
                total_users = metric_values[1].get('value', '0') if len(metric_values) > 1 else '0'
                new_users = metric_values[2].get('value', '0') if len(metric_values) > 2 else '0'
                event_count = metric_values[3].get('value', '0') if len(metric_values) > 3 else '0'
                
                # 計算値（簡略化版、実際のGA4と同じ計算が必要な場合は要調整）
                writer.writerow([
                    channel,
                    event,
                    sessions,
                    new_users,
                    0,  # 平均エンゲージメント時間は別途取得が必要
                    0,  # セッションあたりのイベント数は別途計算が必要
                    event_count
                ])
        
        print(f"✅ 保存完了: {filename}")
        print(f"   取得件数: {len(rows)} 行")
        
    except Exception as e:
        print(f"❌ エラー: {str(e)}")
        return False
    
    return True


def fetch_ga4_events(access_token, property_id, start_date, end_date, output_dir):
    """
    GA4からイベント（コンバージョン）データを取得
    
    ガイド: knowledge/ga4-guide.md 参照
    """
    print("\n📊 GA4: イベントデータを取得中...")
    
    request_body = {
        "dateRanges": [{
            "startDate": start_date,
            "endDate": end_date
        }],
        "metrics": [
            {"name": "eventCount"},
            {"name": "totalUsers"},
            {"name": "eventCountPerUser"}
        ],
        "dimensions": [{"name": "eventName"}],
        "orderBys": [
            {
                "metric": {
                    "metricName": "eventCount"
                },
                "desc": True
            }
        ]
    }
    
    try:
        response = fetch_ga4_report(access_token, property_id, request_body, "イベント")
        
        # CSVファイルに保存
        filename = os.path.join(output_dir, "イベント_イベント名.csv")
        
        with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow([
                "# ----------------------------------------",
                "# イベント: イベント名",
                "# アカウント: 株式会社 コモンウェルスエンジニアーズ",
                "# プロパティ: cectokyo.com",
                "# ----------------------------------------"
            ])
            writer.writerow(["# "])
            writer.writerow(["# すべてのユーザー"])
            writer.writerow([f"# 開始日: {start_date.replace('-', '')}"])
            writer.writerow([f"# 終了日: {end_date.replace('-', '')}"])
            
            # データヘッダー
            writer.writerow([
                "イベント名",
                "イベント数",
                "総ユーザー数",
                "アクティブ ユーザーあたりのイベント数",
                "合計収益"
            ])
            
            # データ行
            rows = response.get('rows', [])
            for row in rows:
                dimension_values = row.get('dimensionValues', [])
                metric_values = row.get('metricValues', [])
                
                event_name = dimension_values[0].get('value', '') if len(dimension_values) > 0 else ''
                event_count = metric_values[0].get('value', '0') if len(metric_values) > 0 else '0'
                total_users = metric_values[1].get('value', '0') if len(metric_values) > 1 else '0'
                events_per_user = metric_values[2].get('value', '0') if len(metric_values) > 2 else '0'
                
                writer.writerow([
                    event_name,
                    event_count,
                    total_users,
                    events_per_user,
                    0  # 収益は別途設定が必要
                ])
        
        print(f"✅ 保存完了: {filename}")
        print(f"   取得件数: {len(rows)} 行")
        
    except Exception as e:
        print(f"❌ エラー: {str(e)}")
        return False
    
    return True


def fetch_ga4_pages(access_token, property_id, start_date, end_date, output_dir):
    """
    GA4からページとスクリーンデータを取得
    
    ガイド: knowledge/ga4-guide.md 参照
    """
    print("\n📊 GA4: ページとスクリーンデータを取得中...")
    
    request_body = {
        "dateRanges": [{
            "startDate": start_date,
            "endDate": end_date
        }],
        "metrics": [
            {"name": "screenPageViews"},
            {"name": "totalUsers"},
            {"name": "averageSessionDuration"},
            {"name": "bounceRate"}
        ],
        "dimensions": [{"name": "pagePath"}],
        "orderBys": [
            {
                "metric": {
                    "metricName": "screenPageViews"
                },
                "desc": True
            }
        ],
        "limit": 100  # 上位100ページ
    }
    
    try:
        response = fetch_ga4_report(access_token, property_id, request_body, "ページとスクリーン")
        
        # CSVファイルに保存
        filename = os.path.join(output_dir, "ページとスクリーン_ページパスとスクリーン_クラス.csv")
        
        with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow([
                "# ----------------------------------------",
                "# ページとスクリーン: ページパス",
                "# アカウント: 株式会社 コモンウェルスエンジニアーズ",
                "# プロパティ: cectokyo.com",
                "# ----------------------------------------"
            ])
            writer.writerow(["# "])
            writer.writerow(["# すべてのユーザー"])
            writer.writerow([f"# 開始日: {start_date.replace('-', '')}"])
            writer.writerow([f"# 終了日: {end_date.replace('-', '')}"])
            
            # データヘッダー
            writer.writerow([
                "ページパス",
                "表示回数",
                "ユーザー",
                "平均エンゲージメント時間",
                "直帰率"
            ])
            
            # データ行
            rows = response.get('rows', [])
            for row in rows:
                dimension_values = row.get('dimensionValues', [])
                metric_values = row.get('metricValues', [])
                
                page_path = dimension_values[0].get('value', '') if len(dimension_values) > 0 else ''
                views = metric_values[0].get('value', '0') if len(metric_values) > 0 else '0'
                users = metric_values[1].get('value', '0') if len(metric_values) > 1 else '0'
                avg_duration = metric_values[2].get('value', '0') if len(metric_values) > 2 else '0'
                bounce_rate = metric_values[3].get('value', '0') if len(metric_values) > 3 else '0'
                
                writer.writerow([
                    page_path,
                    views,
                    users,
                    avg_duration,
                    bounce_rate
                ])
        
        print(f"✅ 保存完了: {filename}")
        print(f"   取得件数: {len(rows)} 行")
        
    except Exception as e:
        print(f"❌ エラー: {str(e)}")
        return False
    
    return True


def fetch_search_console_queries(access_token, site_url, start_date, end_date, output_dir):
    """
    Search Consoleから検索クエリデータを取得
    
    ガイド: knowledge/search-console-guide.md 参照
    """
    print("\n📊 Search Console: 検索クエリデータを取得中...")
    
    # URLエンコード
    encoded_site_url = urllib.parse.quote(site_url, safe='')
    
    api_url = f"https://www.googleapis.com/webmasters/v3/sites/{encoded_site_url}/searchAnalytics/query"
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    request_body = {
        'startDate': start_date,
        'endDate': end_date,
        'dimensions': ['query'],
        'rowLimit': 1000
    }
    
    try:
        response = requests.post(api_url, headers=headers, json=request_body)
        response.raise_for_status()
        
        data = response.json()
        
        # CSVファイルに保存
        filename = os.path.join(output_dir, "クエリ.csv")
        
        with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow([
                "上位のクエリ",
                "クリック数",
                "表示回数",
                "CTR",
                "掲載順位"
            ])
            
            # データ行
            rows = data.get('rows', [])
            if rows:
                for row in rows:
                    query = row['keys'][0]
                    clicks = row.get('clicks', 0)
                    impressions = row.get('impressions', 0)
                    ctr = row.get('ctr', 0)
                    position = row.get('position', 0)
                    
                    writer.writerow([
                        query,
                        clicks,
                        impressions,
                        f"{ctr * 100:.2f}%",
                        round(position, 2)
                    ])
                
                print(f"✅ 保存完了: {filename}")
                print(f"   取得件数: {len(rows)} 行")
            else:
                print("⚠️  データがありません")
                writer.writerow([])
        
        return True
                
    except requests.exceptions.HTTPError as e:
        error_detail = ""
        try:
            error_json = e.response.json()
            error_detail = error_json.get('error', {}).get('message', str(e))
        except:
            error_detail = str(e)
        print(f"❌ エラー: {error_detail}")
        return False
    except Exception as e:
        print(f"❌ エラー: {str(e)}")
        return False


def fetch_search_console_pages(access_token, site_url, start_date, end_date, output_dir):
    """
    Search Consoleからページ別パフォーマンスデータを取得
    
    ガイド: knowledge/search-console-guide.md 参照
    """
    print("\n📊 Search Console: ページ別パフォーマンスデータを取得中...")
    
    # URLエンコード
    encoded_site_url = urllib.parse.quote(site_url, safe='')
    
    api_url = f"https://www.googleapis.com/webmasters/v3/sites/{encoded_site_url}/searchAnalytics/query"
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    request_body = {
        'startDate': start_date,
        'endDate': end_date,
        'dimensions': ['page'],
        'rowLimit': 1000
    }
    
    try:
        response = requests.post(api_url, headers=headers, json=request_body)
        response.raise_for_status()
        
        data = response.json()
        
        # CSVファイルに保存
        filename = os.path.join(output_dir, "ページ.csv")
        
        with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow([
                "上位のページ",
                "クリック数",
                "表示回数",
                "CTR",
                "掲載順位"
            ])
            
            # データ行
            rows = data.get('rows', [])
            if rows:
                for row in rows:
                    page = row['keys'][0]
                    clicks = row.get('clicks', 0)
                    impressions = row.get('impressions', 0)
                    ctr = row.get('ctr', 0)
                    position = row.get('position', 0)
                    
                    writer.writerow([
                        page,
                        clicks,
                        impressions,
                        f"{ctr * 100:.2f}%",
                        round(position, 2)
                    ])
                
                print(f"✅ 保存完了: {filename}")
                print(f"   取得件数: {len(rows)} 行")
            else:
                print("⚠️  データがありません")
                writer.writerow([])
        
        return True
                
    except requests.exceptions.HTTPError as e:
        error_detail = ""
        try:
            error_json = e.response.json()
            error_detail = error_json.get('error', {}).get('message', str(e))
        except:
            error_detail = str(e)
        print(f"❌ エラー: {error_detail}")
        return False
    except Exception as e:
        print(f"❌ エラー: {str(e)}")
        return False


def fetch_search_console_devices(access_token, site_url, start_date, end_date, output_dir):
    """
    Search Consoleからデバイス別パフォーマンスデータを取得
    
    ガイド: knowledge/search-console-guide.md 参照
    """
    print("\n📊 Search Console: デバイス別パフォーマンスデータを取得中...")
    
    # URLエンコード
    encoded_site_url = urllib.parse.quote(site_url, safe='')
    
    api_url = f"https://www.googleapis.com/webmasters/v3/sites/{encoded_site_url}/searchAnalytics/query"
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    request_body = {
        'startDate': start_date,
        'endDate': end_date,
        'dimensions': ['device'],
        'rowLimit': 10
    }
    
    try:
        response = requests.post(api_url, headers=headers, json=request_body)
        response.raise_for_status()
        
        data = response.json()
        
        # CSVファイルに保存
        filename = os.path.join(output_dir, "デバイス.csv")
        
        with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
            writer = csv.writer(f)
            
            # ヘッダー
            writer.writerow([
                "デバイス",
                "クリック数",
                "表示回数",
                "CTR",
                "掲載順位"
            ])
            
            # データ行
            rows = data.get('rows', [])
            if rows:
                device_map = {
                    'DESKTOP': 'PC',
                    'MOBILE': 'モバイル',
                    'TABLET': 'タブレット'
                }
                
                for row in rows:
                    device_key = row['keys'][0]
                    device = device_map.get(device_key, device_key)
                    clicks = row.get('clicks', 0)
                    impressions = row.get('impressions', 0)
                    ctr = row.get('ctr', 0)
                    position = row.get('position', 0)
                    
                    writer.writerow([
                        device,
                        clicks,
                        impressions,
                        f"{ctr * 100:.2f}%",
                        round(position, 2)
                    ])
                
                print(f"✅ 保存完了: {filename}")
                print(f"   取得件数: {len(rows)} 行")
            else:
                print("⚠️  データがありません")
                writer.writerow([])
        
        return True
                
    except requests.exceptions.HTTPError as e:
        error_detail = ""
        try:
            error_json = e.response.json()
            error_detail = error_json.get('error', {}).get('message', str(e))
        except:
            error_detail = str(e)
        print(f"❌ エラー: {error_detail}")
        return False
    except Exception as e:
        print(f"❌ エラー: {str(e)}")
        return False


def main():
    """メイン処理"""
    print("🚀 データ取得プログラム開始")
    print("="*60)
    
    # 設定ファイルを読み込む
    config = load_config()
    if not config:
        return
    
    # 日付範囲を設定（定数から取得）
    start_date = START_DATE
    end_date = END_DATE
    
    print(f"📅 対象期間: {start_date} 〜 {end_date}")
    
    # 出力ディレクトリを作成（年月フォルダ名は開始日の年月を使用）
    year_month = start_date[:7].replace('-', '-')
    output_dir = os.path.join(
        os.path.dirname(__file__),
        config.get('output', {}).get('data_dir', 'data'),
        year_month
    )
    os.makedirs(output_dir, exist_ok=True)
    print(f"📁 出力先: {output_dir}")
    
    # GA4データを取得
    ga4_config = config.get('ga4', {})
    property_id = ga4_config.get('property_id')
    
    if property_id and property_id != "YOUR_GA4_PROPERTY_ID":
        try:
            # アクセストークンを取得
            credentials_file = ga4_config.get('credentials_file')
            scopes = ['https://www.googleapis.com/auth/analytics.readonly']
            access_token = get_access_token(credentials_file, scopes)
            
            fetch_ga4_traffic_acquisition(access_token, property_id, start_date, end_date, output_dir)
            fetch_ga4_events(access_token, property_id, start_date, end_date, output_dir)
            fetch_ga4_pages(access_token, property_id, start_date, end_date, output_dir)
            
        except Exception as e:
            print(f"❌ GA4データ取得エラー: {str(e)}")
    else:
        print("⚠️  GA4のプロパティIDが設定されていません。スキップします。")
    
    # Search Consoleデータを取得
    sc_config = config.get('search_console', {})
    site_url = sc_config.get('site_url')
    
    if site_url and site_url != "sc-domain:YOUR_SITE":
        try:
            # アクセストークンを取得
            credentials_file = sc_config.get('credentials_file')
            scopes = ['https://www.googleapis.com/auth/webmasters.readonly']
            access_token = get_access_token(credentials_file, scopes)
            
            fetch_search_console_queries(access_token, site_url, start_date, end_date, output_dir)
            fetch_search_console_pages(access_token, site_url, start_date, end_date, output_dir)
            fetch_search_console_devices(access_token, site_url, start_date, end_date, output_dir)
            
        except Exception as e:
            print(f"❌ Search Consoleデータ取得エラー: {str(e)}")
    else:
        print("⚠️  Search ConsoleのサイトURLが設定されていません。スキップします。")
    
    print("\n" + "="*60)
    print("✅ データ取得が完了しました！")
    print("="*60)


if __name__ == "__main__":
    main()
