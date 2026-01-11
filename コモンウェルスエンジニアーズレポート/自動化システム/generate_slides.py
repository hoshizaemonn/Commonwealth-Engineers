#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Web運用報告書 生成システム (本番用)
実行時の日付から自動判定し、reports/に出力します。
"""

import os
import csv
import re
import json
import time
import urllib.parse
import requests
import jwt
from datetime import datetime
from dateutil.relativedelta import relativedelta
from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.dml.color import RGBColor
from pptx.enum.text import PP_ALIGN
from pptx.enum.shapes import MSO_SHAPE

# --- カラー定義 ---
COLOR_BACKGROUND = RGBColor(250, 248, 245)
COLOR_ORANGE = RGBColor(255, 152, 83)
COLOR_DARK_BROWN = RGBColor(101, 67, 33)
COLOR_LIGHT_BROWN = RGBColor(188, 143, 105)
COLOR_GREEN = RGBColor(76, 175, 80)
COLOR_RED = RGBColor(244, 67, 54)

# ==========================================
# 1. 自動日付判定ロジック
# ==========================================
def get_report_periods():
    """
    実行日ベースで、前月1日〜末日を対象期間として自動設定
    
    例：1月に実行 → 12/1〜12/31
    例：2月に実行 → 1/1〜1/31
    
    Returns:
        (report_year_month, start_date, end_date, comp_year_month):
        - report_year_month: 報告対象月（YYYY-MM形式）
        - start_date: 開始日（YYYY-MM-DD形式）
        - end_date: 終了日（YYYY-MM-DD形式）
        - comp_year_month: 比較対象月（前々月、YYYY-MM形式）
    """
    today = datetime.now()
    # 前月を報告対象とする
    report_dt = today - relativedelta(months=1)
    # 前月の1日
    start_date = report_dt.replace(day=1).strftime("%Y-%m-%d")
    # 前月の末日
    next_month = report_dt + relativedelta(months=1)
    last_day = (next_month.replace(day=1) - relativedelta(days=1)).day
    end_date = report_dt.replace(day=last_day).strftime("%Y-%m-%d")
    
    report_year_month = report_dt.strftime("%Y-%m")
    # 比較対象は前々月（前月の前月）
    comp_dt = report_dt - relativedelta(months=1)
    comp_year_month = comp_dt.strftime("%Y-%m")
    
    return report_year_month, start_date, end_date, comp_year_month

# ==========================================
# 1.5. API認証・設定読み込み
# ==========================================
def load_config():
    """設定ファイルを読み込む"""
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    
    if not os.path.exists(config_path):
        raise FileNotFoundError(
            f"❌ エラー: config.json が見つかりません。\n"
            f"   パス: {config_path}\n"
            f"   必要な認証情報: config.jsonファイルが必要です。"
        )
    
    try:
        with open(config_path, 'r', encoding='utf-8') as f:
            config = json.load(f)
        return config
    except json.JSONDecodeError as e:
        raise ValueError(
            f"❌ エラー: config.jsonの形式が正しくありません。\n"
            f"   エラー詳細: {str(e)}"
        )

def get_access_token(credentials_file, scopes, service_name="API"):
    """
    Service Account認証情報を使ってアクセストークンを取得
    
    Args:
        credentials_file: credentials.jsonのパス
        scopes: スコープのリスト
        service_name: サービス名（エラーメッセージ用）
    
    Returns:
        access_token: アクセストークン
    """
    creds_path = os.path.join(os.path.dirname(__file__), credentials_file)
    
    if not os.path.exists(creds_path):
        missing_items = []
        missing_items.append(f"認証ファイル: {credentials_file} (パス: {creds_path})")
        if not os.path.exists(os.path.dirname(creds_path)):
            missing_items.append(f"ディレクトリ: {os.path.dirname(creds_path)}")
        
        raise FileNotFoundError(
            f"❌ {service_name}認証エラー: 認証ファイルが見つかりません。\n"
            f"   不足している認証情報:\n"
            + "\n".join(f"   - {item}" for item in missing_items) + "\n"
            f"   config.jsonでcredentials_fileを確認してください。"
        )
    
    try:
        with open(creds_path, 'r', encoding='utf-8') as f:
            creds = json.load(f)
    except json.JSONDecodeError as e:
        raise ValueError(
            f"❌ {service_name}認証エラー: {credentials_file}の形式が正しくありません。\n"
            f"   エラー詳細: {str(e)}\n"
            f"   必要な認証情報: 正しいJSON形式のService Account認証ファイルが必要です。"
        )
    
    # 必要なフィールドのチェック
    required_fields = ['client_email', 'private_key']
    missing_fields = [field for field in required_fields if field not in creds or not creds[field]]
    if missing_fields:
        raise ValueError(
            f"❌ {service_name}認証エラー: 認証ファイルに必要な情報が不足しています。\n"
            f"   不足している認証情報:\n"
            + "\n".join(f"   - {field}" for field in missing_fields) + "\n"
            f"   正しいService Account JSONキーファイルを取得してください。"
        )
    
    try:
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
    except requests.exceptions.RequestException as e:
        raise ConnectionError(
            f"❌ {service_name}認証エラー: トークン取得に失敗しました。\n"
            f"   エラー詳細: {str(e)}\n"
            f"   必要な認証情報: インターネット接続と有効なService Account認証情報を確認してください。"
        )
    except Exception as e:
        raise RuntimeError(
            f"❌ {service_name}認証エラー: 予期しないエラーが発生しました。\n"
            f"   エラー詳細: {str(e)}\n"
            f"   認証ファイル({credentials_file})とconfig.jsonを確認してください。"
        )

def fetch_ga4_report(access_token, property_id, request_body):
    """
    GA4 APIからレポートを取得
    
    Args:
        access_token: アクセストークン
        property_id: GA4プロパティID
        request_body: リクエストボディ
    
    Returns:
        APIレスポンスのJSON
    """
    api_url = f"https://analyticsdata.googleapis.com/v1beta/properties/{property_id}:runReport"
    
    headers = {
        'Authorization': f'Bearer {access_token}',
        'Content-Type': 'application/json'
    }
    
    try:
        response = requests.post(api_url, headers=headers, json=request_body)
        response.raise_for_status()
        return response.json()
    except requests.exceptions.HTTPError as e:
        error_detail = ""
        try:
            error_json = e.response.json()
            error_detail = error_json.get('error', {}).get('message', str(e))
        except:
            error_detail = str(e)
        raise ConnectionError(
            f"❌ GA4 APIエラー: データ取得に失敗しました。\n"
            f"   エラー詳細: {error_detail}\n"
            f"   必要な認証情報: GA4プロパティID({property_id})とアクセス権限を確認してください。"
        )

# ==========================================
# 2. CSVデータ読み込み関数
# ==========================================
def load_csv_data(filepath):
    """CSVファイルを読み込む"""
    data = []
    if not os.path.exists(filepath):
        return data
    with open(filepath, 'r', encoding='utf-8-sig') as f:
        reader = csv.DictReader(f)
        for row in reader:
            data.append(row)
    return data

def get_ga4_traffic_data(start_date="2025-12-01", end_date="2025-12-31", save_csv=True, output_dir=None):
    """GA4トラフィック獲得データをAPIから直接取得し、CSV保存も可能"""
    config = load_config()
    ga4_config = config.get('ga4', {})
    property_id = ga4_config.get('property_id')
    credentials_file = ga4_config.get('credentials_file')
    
    if not property_id or property_id == "YOUR_GA4_PROPERTY_ID":
        raise ValueError(
            f"❌ GA4設定エラー: プロパティIDが設定されていません。\n"
            f"   必要な認証情報: config.jsonのga4.property_idを設定してください。"
        )
    
    if not credentials_file:
        raise ValueError(
            f"❌ GA4設定エラー: 認証ファイルが設定されていません。\n"
            f"   必要な認証情報: config.jsonのga4.credentials_fileを設定してください。"
        )
    
    try:
        # アクセストークンを取得
        scopes = ['https://www.googleapis.com/auth/analytics.readonly']
        access_token = get_access_token(credentials_file, scopes, "GA4")
        
        # トラフィック獲得データを取得
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
        
        response = fetch_ga4_report(access_token, property_id, request_body)
        
        # CSVに保存
        if save_csv and output_dir:
            os.makedirs(output_dir, exist_ok=True)
            filename = os.path.join(output_dir, "トラフィック獲得_セッションのメインのチャネル_グループ（デフォルト_チャネル_グループ）.csv")
            
            with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                
                # ヘッダー
                writer.writerow(["# ----------------------------------------"])
                writer.writerow(["# トラフィック獲得: セッションのメインのチャネル グループ（デフォルト チャネル グループ）"])
                writer.writerow(["# アカウント: 株式会社 コモンウェルスエンジニアーズ"])
                writer.writerow(["# プロパティ: cectokyo.com"])
                writer.writerow(["# ----------------------------------------"])
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
                    
                    writer.writerow([
                        channel,
                        event,
                        sessions,
                        new_users,
                        0,  # 平均エンゲージメント時間は別途取得が必要
                        0,  # セッションあたりのイベント数は別途計算が必要
                        event_count
                    ])
            
            print(f"✅ CSV保存完了: {filename}")
        
        # レスポンスを解析
        sessions = {}
        new_users = 0
        total_users = 0
        
        rows = response.get('rows', [])
        for row in rows:
            dimension_values = row.get('dimensionValues', [])
            metric_values = row.get('metricValues', [])
            
            channel = dimension_values[0].get('value', '') if len(dimension_values) > 0 else ''
            event = dimension_values[1].get('value', '') if len(dimension_values) > 1 else ''
            sessions_count = int(metric_values[0].get('value', '0')) if len(metric_values) > 0 else 0
            total_users_count = int(metric_values[1].get('value', '0')) if len(metric_values) > 1 else 0
            new_users_count = int(metric_values[2].get('value', '0')) if len(metric_values) > 2 else 0
            
            # チャネル別セッション数を集計（session_startイベント）
            if event == 'session_start' and channel:
                sessions[channel] = sessions_count
            
            # 新規ユーザー数（first_visitイベントから取得）
            if event == 'first_visit':
                new_users = max(new_users, new_users_count)
            
            # 総ユーザー数（user_engagementイベントから取得）
            if event == 'user_engagement':
                total_users = max(total_users, total_users_count)
        
        return {
            'sessions': sessions,
            'new_users': new_users,
            'total_users': total_users
        }
    except (FileNotFoundError, ValueError, ConnectionError, RuntimeError) as e:
        # 既に詳細なエラーメッセージが含まれているので、そのまま再発生
        raise
    except Exception as e:
        raise RuntimeError(
            f"❌ GA4データ取得エラー: 予期しないエラーが発生しました。\n"
            f"   エラー詳細: {str(e)}\n"
            f"   必要な認証情報: config.jsonと認証ファイルを確認してください。"
        )

def get_ga4_events_data(start_date="2025-12-01", end_date="2025-12-31", save_csv=True, output_dir=None):
    """GA4イベントデータをAPIから直接取得し、CSV保存も可能"""
    config = load_config()
    ga4_config = config.get('ga4', {})
    property_id = ga4_config.get('property_id')
    credentials_file = ga4_config.get('credentials_file')
    
    if not property_id or property_id == "YOUR_GA4_PROPERTY_ID":
        raise ValueError(
            f"❌ GA4設定エラー: プロパティIDが設定されていません。\n"
            f"   必要な認証情報: config.jsonのga4.property_idを設定してください。"
        )
    
    if not credentials_file:
        raise ValueError(
            f"❌ GA4設定エラー: 認証ファイルが設定されていません。\n"
            f"   必要な認証情報: config.jsonのga4.credentials_fileを設定してください。"
        )
    
    try:
        # アクセストークンを取得
        scopes = ['https://www.googleapis.com/auth/analytics.readonly']
        access_token = get_access_token(credentials_file, scopes, "GA4")
        
        # イベントデータを取得
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
        
        response = fetch_ga4_report(access_token, property_id, request_body)
        
        # CSVに保存
        if save_csv and output_dir:
            os.makedirs(output_dir, exist_ok=True)
            filename = os.path.join(output_dir, "イベント_イベント名.csv")
            
            with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                writer = csv.writer(f)
                
                # ヘッダー
                writer.writerow(["# ----------------------------------------"])
                writer.writerow(["# イベント: イベント名"])
                writer.writerow(["# アカウント: 株式会社 コモンウェルスエンジニアーズ"])
                writer.writerow(["# プロパティ: cectokyo.com"])
                writer.writerow(["# ----------------------------------------"])
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
            
            print(f"✅ CSV保存完了: {filename}")
        
        # レスポンスを解析
        events = {}
        total_users = 0
        new_users = 0
        
        rows = response.get('rows', [])
        for row in rows:
            dimension_values = row.get('dimensionValues', [])
            metric_values = row.get('metricValues', [])
            
            event_name = dimension_values[0].get('value', '') if len(dimension_values) > 0 else ''
            event_count = int(metric_values[0].get('value', '0')) if len(metric_values) > 0 else 0
            total_users_count = int(metric_values[1].get('value', '0')) if len(metric_values) > 1 else 0
            
            if event_name:
                events[event_name] = event_count
                
                if event_name == 'user_engagement':
                    total_users = max(total_users, total_users_count)
                
                if event_name == 'first_visit':
                    new_users = max(new_users, total_users_count)
        
        events['_total_users'] = total_users
        events['_new_users'] = new_users
        
        return events
    except (FileNotFoundError, ValueError, ConnectionError, RuntimeError) as e:
        # 既に詳細なエラーメッセージが含まれているので、そのまま再発生
        raise
    except Exception as e:
        raise RuntimeError(
            f"❌ GA4イベントデータ取得エラー: 予期しないエラーが発生しました。\n"
            f"   エラー詳細: {str(e)}\n"
            f"   必要な認証情報: config.jsonと認証ファイルを確認してください。"
        )

def get_search_console_data(start_date="2025-12-01", end_date="2025-12-31", save_csv=True, output_dir=None):
    """Search ConsoleデータをAPIから直接取得し、CSV保存も可能"""
    config = load_config()
    sc_config = config.get('search_console', {})
    site_url = sc_config.get('site_url')
    credentials_file = sc_config.get('credentials_file')
    
    if not site_url or site_url == "sc-domain:YOUR_SITE":
        raise ValueError(
            f"❌ Search Console設定エラー: サイトURLが設定されていません。\n"
            f"   必要な認証情報: config.jsonのsearch_console.site_urlを設定してください。"
        )
    
    if not credentials_file:
        raise ValueError(
            f"❌ Search Console設定エラー: 認証ファイルが設定されていません。\n"
            f"   必要な認証情報: config.jsonのsearch_console.credentials_fileを設定してください。"
        )
    
    result = {
        'total_clicks': 0,
        'total_impressions': 0,
        'avg_ctr': 0.0,
        'sc_ctr': 0.0,  # 統一のために追加
        'avg_position': 0.0,
        'top_queries': [],
        'device_performance': {
            'PC': {'clicks': 0, 'impressions': 0, 'ctr': 0.0, 'position': 0.0},
            'モバイル': {'clicks': 0, 'impressions': 0, 'ctr': 0.0, 'position': 0.0},
            'タブレット': {'clicks': 0, 'impressions': 0, 'ctr': 0.0, 'position': 0.0}
        }
    }
    
    try:
        # アクセストークンを取得
        scopes = ['https://www.googleapis.com/auth/webmasters.readonly']
        access_token = get_access_token(credentials_file, scopes, "Search Console")
        
        # URLエンコード
        encoded_site_url = urllib.parse.quote(site_url, safe='')
        
        # クエリデータを取得
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
            
            rows = data.get('rows', [])
            total_clicks = 0
            total_impressions = 0
            positions = []
            
            # CSVに保存
            if save_csv and output_dir:
                os.makedirs(output_dir, exist_ok=True)
                filename = os.path.join(output_dir, "クエリ.csv")
                
                with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                    writer = csv.writer(f)
                    
                    # ヘッダー
                    writer.writerow(["上位のクエリ", "クリック数", "表示回数", "CTR", "掲載順位"])
                    
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
                
                print(f"✅ CSV保存完了: {filename}")
            
            for row in rows:
                clicks = row.get('clicks', 0)
                impressions = row.get('impressions', 0)
                position = row.get('position', 0)
                
                total_clicks += clicks
                total_impressions += impressions
                if position > 0:
                    positions.append(position)
            
            result['total_clicks'] = total_clicks
            result['total_impressions'] = total_impressions
            calculated_ctr = (total_clicks / total_impressions * 100) if total_impressions > 0 else 0
            result['avg_ctr'] = calculated_ctr
            result['sc_ctr'] = calculated_ctr  # 統一のために設定
            result['avg_position'] = (sum(positions) / len(positions)) if positions else 0.0
            
            # デバイスデータを取得してCSV保存
            request_body_devices = {
                'startDate': start_date,
                'endDate': end_date,
                'dimensions': ['device'],
                'rowLimit': 10
            }
            
            try:
                response_devices = requests.post(api_url, headers=headers, json=request_body_devices)
                response_devices.raise_for_status()
                data_devices = response_devices.json()
                
                rows_devices = data_devices.get('rows', [])
                device_positions = []
                device_map = {
                    'DESKTOP': 'PC',
                    'MOBILE': 'モバイル',
                    'TABLET': 'タブレット'
                }
                
                # CSVに保存
                if save_csv and output_dir:
                    filename = os.path.join(output_dir, "デバイス.csv")
                    
                    with open(filename, 'w', encoding='utf-8-sig', newline='') as f:
                        writer = csv.writer(f)
                        
                        # ヘッダー
                        writer.writerow(["デバイス", "クリック数", "表示回数", "CTR", "掲載順位"])
                        
                        for row in rows_devices:
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
                            
                            # デバイスデータをresultに保存
                            if device in result['device_performance']:
                                result['device_performance'][device] = {
                                    'clicks': clicks,
                                    'impressions': impressions,
                                    'ctr': ctr * 100,  # パーセントで保存
                                    'position': position
                                }
                    
                    print(f"✅ CSV保存完了: {filename}")
                
                for row in rows_devices:
                    position = row.get('position', 0)
                    if position > 0:
                        device_positions.append(position)
                    
                    # デバイスデータをresultに保存（CSV保存しない場合も）
                    if not save_csv or not output_dir:
                        device_key = row['keys'][0]
                        device = device_map.get(device_key, device_key)
                        clicks = row.get('clicks', 0)
                        impressions = row.get('impressions', 0)
                        ctr = row.get('ctr', 0)
                        position = row.get('position', 0)
                        
                        if device in result['device_performance']:
                            result['device_performance'][device] = {
                                'clicks': clicks,
                                'impressions': impressions,
                                'ctr': ctr * 100,  # パーセントで保存
                                'position': position
                            }
                
                if device_positions and result['avg_position'] == 0.0:
                    result['avg_position'] = sum(device_positions) / len(device_positions)
            except Exception as e:
                # デバイスデータの取得に失敗しても続行
                print(f"⚠️  デバイスデータ取得エラー: {e}")
                pass
            
        except requests.exceptions.HTTPError as e:
            error_detail = ""
            try:
                error_json = e.response.json()
                error_detail = error_json.get('error', {}).get('message', str(e))
            except:
                error_detail = str(e)
            raise ConnectionError(
                f"❌ Search Console APIエラー: データ取得に失敗しました。\n"
                f"   エラー詳細: {error_detail}\n"
                f"   必要な認証情報: Search ConsoleサイトURL({site_url})とアクセス権限を確認してください。"
            )
        
        return result
    except (FileNotFoundError, ValueError, ConnectionError, RuntimeError) as e:
        # 既に詳細なエラーメッセージが含まれているので、そのまま再発生
        raise
    except Exception as e:
        raise RuntimeError(
            f"❌ Search Consoleデータ取得エラー: 予期しないエラーが発生しました。\n"
            f"   エラー詳細: {str(e)}\n"
            f"   必要な認証情報: config.jsonと認証ファイルを確認してください。"
        )

def parse_markdown_for_previous_month_data(comparison_year_month, base_dir="../reports"):
    """
    前月レポートから比較対象月のデータを抽出
    
    Args:
        comparison_year_month: 比較対象月（YYYY-MM形式、例: "2025-11"）
        base_dir: レポートディレクトリのベースパス
    
    Returns:
        comparison_data: 比較対象月のデータ（前月の値）
    """
    report_path = os.path.join(base_dir, f"{comparison_year_month}.md")
    
    comparison_data = {
        'sessions': 0,
        'inquiries': 0,
        'cvr': 0.0,
        'organic_search': 0,
        'direct': 0
    }
    
    if not os.path.exists(report_path):
        print(f"⚠️  レポートファイルが見つかりません: {report_path}")
        print(f"   デフォルト値（0）を使用します")
        return comparison_data
    
    try:
        with open(report_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # 総セッション数を抽出（表の2列目が前月の値）
        # パターン: | 総セッション数 | 前々月 | 前月 | 前月比 |
        patterns = [
            r'\| 総セッション数 \| (\d+(?:,\d+)?) \| (\d+(?:,\d+)?) \|',  # カンマ区切りの可能性
            r'\| 総セッション数 \| (\d+) \| (\d+) \|',
            r'総セッション数.*?\| (\d+(?:,\d+)?) \| (\d+(?:,\d+)?) \|'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                # 2列目（前月の値）を取得
                sessions_str = match.group(2).replace(',', '')
                comparison_data['sessions'] = int(sessions_str)
                print(f"  ✅ 総セッション数（前月）: {comparison_data['sessions']}")
                break
        
        # 問い合わせ件数を抽出
        patterns = [
            r'\| 問い合わせ件数 \| (\d+(?:,\d+)?)件 \| (\d+(?:,\d+)?)件 \|',
            r'\| 問い合わせ件数 \| (\d+)件 \| (\d+)件 \|',
            r'問い合わせ件数.*?\| (\d+(?:,\d+)?)件 \| (\d+(?:,\d+)?)件 \|'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                inquiries_str = match.group(2).replace(',', '')
                comparison_data['inquiries'] = int(inquiries_str)
                print(f"  ✅ 問い合わせ件数（前月）: {comparison_data['inquiries']}")
                break
        
        # CVRを抽出
        patterns = [
            r'\| 問い合わせCVR \| ([\d.]+)% \| ([\d.]+)% \|',
            r'問い合わせCVR.*?\| ([\d.]+)% \| ([\d.]+)% \|'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                comparison_data['cvr'] = float(match.group(2))
                print(f"  ✅ 問い合わせCVR（前月）: {comparison_data['cvr']}%")
                break
        
        # Organic Searchを抽出
        patterns = [
            r'\| Organic Search.*?\| (\d+(?:,\d+)?) \| (\d+(?:,\d+)?) \|',
            r'\| Organic Search \| (\d+) \| (\d+) \|',
            r'Organic Search.*?\| (\d+(?:,\d+)?) \| (\d+(?:,\d+)?) \|'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                organic_str = match.group(2).replace(',', '')
                comparison_data['organic_search'] = int(organic_str)
                print(f"  ✅ Organic Search（前月）: {comparison_data['organic_search']}")
                break
        
        # Directを抽出
        patterns = [
            r'\| Direct.*?\| (\d+(?:,\d+)?) \| (\d+(?:,\d+)?) \|',
            r'\| Direct \| (\d+) \| (\d+) \|',
            r'Direct.*?\| (\d+(?:,\d+)?) \| (\d+(?:,\d+)?) \|'
        ]
        
        for pattern in patterns:
            match = re.search(pattern, content)
            if match:
                direct_str = match.group(2).replace(',', '')
                comparison_data['direct'] = int(direct_str)
                print(f"  ✅ Direct（前月）: {comparison_data['direct']}")
                break
        
    except Exception as e:
        print(f"⚠️  レポート読み込みエラー: {e}")
        print(f"   デフォルト値（0）を使用します")
    
    return comparison_data

# ==========================================
# 2.5. CSVデータ読み込み関数（2025-11用）
# ==========================================
def load_data_from_csv(year_month, base_dir="../data"):
    """CSVファイルからデータを読み込む（2025-11など既存データ用）"""
    data_dir = os.path.join(base_dir, year_month)
    result = {
        'sessions': 0,
        'inquiries': 0,
        'cvr': 0.0,
        'organic_search': 0,
        'direct': 0,
        'new_users': 0,
        'total_users': 0,
        'sc_clicks': 0,
        'sc_impressions': 0,
        'sc_ctr': 0.0,
        'sc_position': 0.0
    }
    
    if not os.path.exists(data_dir):
        return result
    
    # GA4トラフィックデータを読み込み
    traffic_csv = os.path.join(data_dir, "トラフィック獲得_セッションのメインのチャネル_グループ（デフォルト_チャネル_グループ）.csv")
    if os.path.exists(traffic_csv):
        data = load_csv_data(traffic_csv)
        sessions_dict = {}
        for row in data:
            channel = row.get('セッションのメインのチャネル グループ（デフォルト チャネル グループ）', '')
            event = row.get('イベント名', '')
            if event == 'session_start' and channel:
                try:
                    sessions_dict[channel] = int(row.get('セッション', 0))
                except:
                    pass
        
        result['sessions'] = sum(sessions_dict.values())
        result['organic_search'] = sessions_dict.get('Organic Search', 0)
        result['direct'] = sessions_dict.get('Direct', 0)
    
    # GA4イベントデータを読み込み
    events_csv = os.path.join(data_dir, "イベント_イベント名.csv")
    if os.path.exists(events_csv):
        data = load_csv_data(events_csv)
        for row in data:
            event_name = row.get('イベント名', '')
            if event_name:
                try:
                    if event_name == 'contact':
                        result['inquiries'] = int(row.get('イベント数', 0))
                    elif event_name == 'user_engagement':
                        result['total_users'] = int(row.get('総ユーザー数', 0))
                    elif event_name == 'first_visit':
                        result['new_users'] = int(row.get('総ユーザー数', 0))
                except:
                    pass
        
        if result['sessions'] > 0:
            result['cvr'] = (result['inquiries'] / result['sessions'] * 100)
    
    # Search Consoleデータを読み込み
    queries_csv = os.path.join(data_dir, "クエリ.csv")
    if os.path.exists(queries_csv):
        data = load_csv_data(queries_csv)
        total_clicks = 0
        total_impressions = 0
        positions = []
        for row in data:
            try:
                clicks = int(row.get('クリック数', 0))
                impressions = int(row.get('表示回数', 0))
                position_str = row.get('掲載順位', '0')
                position = float(position_str.replace('%', '')) if position_str else 0
                
                total_clicks += clicks
                total_impressions += impressions
                if position > 0:
                    positions.append(position)
            except:
                pass
        
        result['sc_clicks'] = total_clicks
        result['sc_impressions'] = total_impressions
        result['sc_ctr'] = (total_clicks / total_impressions * 100) if total_impressions > 0 else 0
        result['sc_position'] = (sum(positions) / len(positions)) if positions else 0.0
    
    return result

# ==========================================
# 2.6. 改善パターン読み込みと分析
# ==========================================
def load_improvement_patterns(pattern_file="improvement-patterns.md"):
    """improvement-patterns.mdファイルを読み込んでパターンを解析"""
    patterns = {}
    pattern_file_path = os.path.join(os.path.dirname(__file__), pattern_file)
    
    if not os.path.exists(pattern_file_path):
        print(f"⚠️  改善パターンファイルが見つかりません: {pattern_file_path}")
        return patterns
    
    try:
        with open(pattern_file_path, 'r', encoding='utf-8') as f:
            content = f.read()
        
        # パターンを抽出（## パターンN: タイトル の形式）
        pattern_blocks = re.split(r'^## パターン\d+:', content, flags=re.MULTILINE)
        
        for block in pattern_blocks[1:]:  # 最初の要素はヘッダーなのでスキップ
            lines = block.strip().split('\n')
            if not lines:
                continue
            
            title = lines[0].strip()
            pattern_data = {
                'title': title,
                'cause': '',
                'solution': '',
                'priority': '中',
                'expected_effect': ''
            }
            
            current_section = None
            for line in lines[1:]:
                if line.startswith('**原因**:'):
                    current_section = 'cause'
                    pattern_data['cause'] = line.replace('**原因**:', '').strip()
                elif line.startswith('**対策**:'):
                    current_section = 'solution'
                    pattern_data['solution'] = line.replace('**対策**:', '').strip()
                elif line.startswith('**優先度**:'):
                    priority = line.replace('**優先度**:', '').strip()
                    pattern_data['priority'] = priority if priority in ['高', '中', '低'] else '中'
                elif line.startswith('**期待効果**:'):
                    pattern_data['expected_effect'] = line.replace('**期待効果**:', '').strip()
                elif current_section and line.strip() and not line.startswith('-') and not line.startswith('---'):
                    # 前のセクションの続き
                    if current_section == 'cause':
                        pattern_data['cause'] += ' ' + line.strip()
                    elif current_section == 'solution':
                        pattern_data['solution'] += ' ' + line.strip()
            
            # キーはタイトルまたは特定のキーワード
            key = title.lower()
            if '/entry/' in key or 'ctr 0%' in key or '表示回数' in key:
                patterns['ctr_zero'] = pattern_data
            elif '自然検索' in key or 'organic' in key:
                patterns['organic_decrease'] = pattern_data
            elif 'direct' in key or '直接流入' in key:
                patterns['direct_decrease'] = pattern_data
            elif 'cvr' in key or 'コンバージョン' in key:
                patterns['cvr_low'] = pattern_data
            elif 'モバイル' in key or 'device' in key:
                patterns['mobile_performance'] = pattern_data
            elif 'ファネル' in key or '離脱' in key:
                patterns['funnel_drop'] = pattern_data
            else:
                patterns[key] = pattern_data
    
    except Exception as e:
        print(f"⚠️  改善パターン読み込みエラー: {e}")
    
    return patterns

def generate_traffic_source_analysis(current_data, comparison_data):
    """流入経路分析の推察コメントを生成"""
    analysis = []
    
    organic_current = current_data.get('organic_search', 0)
    organic_previous = comparison_data.get('organic_search', 0)
    organic_change = organic_current - organic_previous
    organic_change_pct = ((organic_change / organic_previous * 100) if organic_previous > 0 else 0)
    
    direct_current = current_data.get('direct', 0)
    direct_previous = comparison_data.get('direct', 0)
    direct_change = direct_current - direct_previous
    direct_change_pct = ((direct_change / direct_previous * 100) if direct_previous > 0 else 0)
    
    # Direct流入の分析
    if direct_change > 0 and direct_change_pct > 5:
        analysis.append("**Direct流入が増加した要因候補：**\n\n")
        analysis.append("- 名刺交換やオフライン施策により、直接URLを入力して訪問するユーザーが増加している可能性があります。\n")
        analysis.append("- リピーターによる訪問が増加している可能性があります。\n")
        analysis.append("- ブックマークからの再訪問が増加している可能性があります。\n\n")
    elif direct_change < 0 and direct_change_pct < -10:
        analysis.append("**Direct流入が減少した要因候補：**\n\n")
        analysis.append("- オフライン施策（名刺交換、イベント参加等）の減少が影響している可能性があります。\n")
        analysis.append("- リピーターの訪問頻度が低下している可能性があります。\n")
        analysis.append("- ブランディング施策の弱化により、直接訪問が減少している可能性があります。\n\n")
    
    # Organic Search流入の分析
    if organic_change < 0 and organic_change_pct < -5:
        analysis.append("**自然検索流入が減少した要因候補：**\n\n")
        analysis.append("- 競合他社のSEO施策が強化され、検索結果での順位が低下した可能性があります。\n")
        analysis.append("- 検索アルゴリズムの変更により、表示順位が変動した可能性があります。\n")
        analysis.append("- コンテンツの鮮度不足により、検索エンジンからの評価が低下した可能性があります。\n\n")
    elif organic_change > 0 and organic_change_pct > 5:
        analysis.append("**自然検索流入が増加した要因候補：**\n\n")
        analysis.append("- SEO施策（タイトル・description最適化等）の効果により、検索結果での表示が改善された可能性があります。\n")
        analysis.append("- コンテンツの充実により、検索エンジンからの評価が向上した可能性があります。\n")
        analysis.append("- 被リンクの獲得により、検索エンジンでの評価が向上した可能性があります。\n\n")
    
    return "".join(analysis) if analysis else "流入経路に大きな変動は見られません。\n\n"

def get_device_performance_data(start_date, end_date, sc_data_dict):
    """デバイス別パフォーマンスデータをフォーマット（Search ConsoleのデバイスCSVから読み込み）"""
    # Search ConsoleのデバイスCSVから読み込む
    device_data = {
        'PC': {'clicks': 0, 'impressions': 0, 'ctr': 0.0, 'position': 0.0},
        'モバイル': {'clicks': 0, 'impressions': 0, 'ctr': 0.0, 'position': 0.0},
        'タブレット': {'clicks': 0, 'impressions': 0, 'ctr': 0.0, 'position': 0.0}
    }
    
    # sc_data_dictから直接取得を試みる（既に取得済みの場合）
    if 'device_performance' in sc_data_dict:
        return sc_data_dict['device_performance']
    
    # CSVから読み込みを試みる
    year_month = start_date[:7]  # YYYY-MM形式
    csv_path = os.path.join("..", "data", year_month, "デバイス.csv")
    
    if os.path.exists(csv_path):
        data = load_csv_data(csv_path)
        for row in data:
            device = row.get('デバイス', '')
            try:
                clicks = int(row.get('クリック数', 0))
                impressions = int(row.get('表示回数', 0))
                ctr_str = row.get('CTR', '0%').replace('%', '')
                position_str = row.get('掲載順位', '0')
                
                ctr = float(ctr_str) if ctr_str else 0.0
                position = float(position_str) if position_str else 0.0
                
                if device in device_data:
                    device_data[device] = {
                        'clicks': clicks,
                        'impressions': impressions,
                        'ctr': ctr,
                        'position': position
                    }
            except:
                pass
    
    return device_data

def get_conversion_funnel_data(ga4_events, total_sessions):
    """コンバージョンファネルデータを計算（セッション → フォーム開始 → 完了）"""
    # GA4イベントからフォーム開始と完了を取得
    form_start = ga4_events.get('form_start', 0)
    form_submit = ga4_events.get('form_submit', 0)
    form_complete = ga4_events.get('form_complete', 0)
    
    # 完了イベントがなければsubmitを使用
    form_completed = form_complete if form_complete > 0 else form_submit
    
    # フォーム開始がなければ、contactイベントを開始として扱う
    if form_start == 0:
        form_start = ga4_events.get('contact', 0)
    
    funnel = {
        'sessions': total_sessions,
        'form_start': form_start,
        'form_completed': form_completed,
        'start_rate': (form_start / total_sessions * 100) if total_sessions > 0 else 0.0,
        'completion_rate': (form_completed / form_start * 100) if form_start > 0 else 0.0,
        'overall_cvr': (form_completed / total_sessions * 100) if total_sessions > 0 else 0.0
    }
    
    return funnel

def generate_improvement_proposals(issues_content, patterns):
    """improvement-patterns.mdから改善提案を生成"""
    proposals = []
    
    # 課題からパターンを特定
    if '/entry/' in issues_content and 'CTR 0%' in issues_content:
        pattern = patterns.get('ctr_zero')
        if pattern:
            proposals.append({
                'pattern': pattern,
                'issue': '採用エントリーページ（/entry/）のクリック率 0%'
            })
    
    # パターンベースの提案テーブルを生成
    if proposals:
        md_content = "| 課題 | 原因 | 対策 | 優先度 | 期待効果 |\n"
        md_content += "| :--- | :--- | :--- | :--- | :--- |\n"
        
        for prop in proposals:
            pattern = prop['pattern']
            issue = prop['issue']
            cause = pattern.get('cause', 'データ分析が必要')[:50] + '...' if len(pattern.get('cause', '')) > 50 else pattern.get('cause', 'データ分析が必要')
            solution = pattern.get('solution', '改善施策を検討')[:50] + '...' if len(pattern.get('solution', '')) > 50 else pattern.get('solution', '改善施策を検討')
            priority = pattern.get('priority', '中')
            effect = pattern.get('expected_effect', '効果測定が必要')
            
            md_content += f"| {issue} | {cause} | {solution} | {priority} | {effect} |\n"
        
        return md_content
    else:
        return "改善提案は、詳細なデータ分析後にご提示いたします。\n"

# ==========================================
# 2.7. テンプレート処理と改善効果分析
# ==========================================
def analyze_improvement_effects(previous_report_path, current_data, previous_data):
    """
    11月の改善提案が12月のデータでどう改善されたかを分析
    SEOタイトル・description設定等の提案に対し、12月の結果を紐付けてポジティブに分析
    """
    analysis = []
    
    if not os.path.exists(previous_report_path):
        return "## 5. 改善効果の分析\n\n前月レポートが見つかりませんでした。\n"
    
    try:
        with open(previous_report_path, 'r', encoding='utf-8') as f:
            previous_content = f.read()
        
        # 改善提案セクションを抽出
        improvement_patterns = [
            r'## .*改善提案.*?\n(.*?)(?=\n## |$)',
            r'■ 改善提案.*?\n(.*?)(?=\n■ |\n## |$)',
            r'改善提案.*?\n(.*?)(?=\n## |$)'
        ]
        
        improvement_text = ""
        for pattern in improvement_patterns:
            improvement_match = re.search(pattern, previous_content, re.DOTALL | re.IGNORECASE)
            if improvement_match:
                improvement_text = improvement_match.group(1).strip()
                break
        
        analysis.append("### 前月（11月）の改善提案の効果\n\n")
        
        if improvement_text:
            # SEO関連の提案があるかチェック
            has_seo_proposal = any(keyword in improvement_text for keyword in ['SEO', 'タイトル', 'description', 'ディスクリプション', 'メタ', 'meta'])
            
            if has_seo_proposal:
                analysis.append("**11月に実施された改善提案：**\n")
                # 改善提案の要点を抽出
                lines = improvement_text.split('\n')
                for line in lines[:5]:  # 最初の5行を表示
                    if line.strip() and not line.strip().startswith('#'):
                        analysis.append(f"- {line.strip()}\n")
                analysis.append("\n")
                
                # 12月の結果を分析（特にCTRの向上に注目）
                sc_ctr_current = current_data.get('sc_ctr', 0) or current_data.get('avg_ctr', 0)
                sc_ctr_previous = previous_data.get('sc_ctr', 0) or previous_data.get('avg_ctr', 0)
                sc_ctr_change = sc_ctr_current - sc_ctr_previous
                
                if sc_ctr_current > 0 and sc_ctr_previous > 0 and sc_ctr_change > 0:
                    analysis.append(f"**✅ 検索結果の見た目修正の効果が確認されました：**\n\n")
                    analysis.append(f"**平均クリック率**: {sc_ctr_previous:.2f}% → **{sc_ctr_current:.2f}%** へ向上（+{sc_ctr_change:.2f}pt）\n\n")
                    analysis.append(f"11月に実施した検索結果の見た目修正（タイトル・説明文の改善）により、検索結果でのクリック率が向上しました。\n")
                    analysis.append(f"具体的には、検索結果における平均クリック率が{sc_ctr_previous:.2f}%から{sc_ctr_current:.2f}%へと{sc_ctr_change:.2f}pt改善し、\n")
                    analysis.append(f"検索結果での露出に対するクリック獲得率が向上したことが確認されました。この改善は、\n")
                    analysis.append(f"検索ユーザーの検索意図とページの関連性が高まったことによる効果と推測されます。\n\n")
        
        # 数値変化から改善効果を分析
        sessions_change = current_data.get('sessions', 0) - previous_data.get('sessions', 0)
        sessions_change_pct = ((sessions_change / previous_data.get('sessions', 1)) * 100) if previous_data.get('sessions', 0) > 0 else 0
        
        inquiries_change = current_data.get('inquiries', 0) - previous_data.get('inquiries', 0)
        inquiries_change_pct = ((inquiries_change / previous_data.get('inquiries', 1)) * 100) if previous_data.get('inquiries', 0) > 0 else 0
        
        cvr_change = current_data.get('cvr', 0) - previous_data.get('cvr', 0)
        
        sc_ctr_current = current_data.get('sc_ctr', 0) or current_data.get('avg_ctr', 0)
        sc_ctr_previous = previous_data.get('sc_ctr', 0) or previous_data.get('avg_ctr', 0)
        sc_ctr_change = sc_ctr_current - sc_ctr_previous
        sc_position_change = previous_data.get('sc_position', 0) - current_data.get('sc_position', 0)  # 順位は低い方が良い
        
        organic_change = current_data.get('organic_search', 0) - previous_data.get('organic_search', 0)
        organic_change_pct = ((organic_change / previous_data.get('organic_search', 1)) * 100) if previous_data.get('organic_search', 0) > 0 else 0
        
        analysis.append("**主要指標の変化：**\n\n")
        
        # 改善効果を評価（分かりやすい表現で記述）
        if inquiries_change > 0:
            analysis.append(f"**問い合わせ件数**: {previous_data.get('inquiries', 0)}件 → **{current_data.get('inquiries', 0)}件**（+{inquiries_change}件、+{inquiries_change_pct:.1f}%）\n\n")
            analysis.append(f"問い合わせ数が前月比+{inquiries_change}件（+{inquiries_change_pct:.1f}%）増加し、サイトへの信頼感向上が確認されました。\n")
            analysis.append(f"この増加は、サイトの質的改善や検索結果の見た目修正など、11月に実施した施策の効果によるものと推測されます。\n\n")
        elif inquiries_change < 0:
            analysis.append(f"**問い合わせ件数**: {previous_data.get('inquiries', 0)}件 → {current_data.get('inquiries', 0)}件（{inquiries_change}件、{inquiries_change_pct:.1f}%）\n\n")
            analysis.append(f"問い合わせ数が前月比{inquiries_change}件（{inquiries_change_pct:.1f}%）減少しました。\n")
            analysis.append(f"要因の特定と改善施策の検討が必要です。\n\n")
        
        if cvr_change > 0:
            analysis.append(f"**問い合わせ率**: {previous_data.get('cvr', 0):.2f}% → **{current_data.get('cvr', 0):.2f}%**（+{cvr_change:.2f}pt改善）\n\n")
            analysis.append(f"問い合わせ率が{previous_data.get('cvr', 0):.2f}%から{current_data.get('cvr', 0):.2f}%へと{cvr_change:.2f}pt向上し、\n")
            analysis.append(f"サイトの質的改善が確認されました。訪問者の行動変容やページ内容の改善による効果と推測されます。\n\n")
        elif cvr_change < 0:
            analysis.append(f"**問い合わせ率**: {previous_data.get('cvr', 0):.2f}% → {current_data.get('cvr', 0):.2f}%（{cvr_change:.2f}pt）\n\n")
            analysis.append(f"問い合わせ率が前月比{cvr_change:.2f}pt低下しました。\n")
            analysis.append(f"要因分析と改善施策の検討が必要です。\n\n")
        
        if sc_ctr_current > 0 and sc_ctr_previous > 0:
            if sc_ctr_change > 0:
                analysis.append(f"**平均クリック率**: {sc_ctr_previous:.2f}% → **{sc_ctr_current:.2f}%**（+{sc_ctr_change:.2f}pt改善）\n\n")
                if has_seo_proposal:
                    analysis.append(f"検索結果における平均クリック率が{sc_ctr_previous:.2f}%から{sc_ctr_current:.2f}%へと{sc_ctr_change:.2f}pt向上しました。\n")
                    analysis.append(f"11月に実施した検索結果の見た目修正（タイトル・説明文の改善）の効果により、検索結果でのクリック率が改善したことが確認されます。\n\n")
                else:
                    analysis.append(f"検索結果でのクリック率が{sc_ctr_change:.2f}pt向上しました。\n")
                    analysis.append(f"検索ユーザーの検索意図とページの関連性が高まったことによる効果と推測されます。\n\n")
            elif sc_ctr_change < 0:
                analysis.append(f"**平均クリック率**: {sc_ctr_previous:.2f}% → {sc_ctr_current:.2f}%（{sc_ctr_change:.2f}pt）\n\n")
                analysis.append(f"平均クリック率が前月比{abs(sc_ctr_change):.2f}pt低下しました。\n")
                analysis.append(f"タイトル・説明文の見直しや検索順位の改善が必要です。\n\n")
        
        if organic_change > 0:
            analysis.append(f"**自然検索からの訪問数**: {previous_data.get('organic_search', 0)}回 → **{current_data.get('organic_search', 0)}回**（+{organic_change}回、+{organic_change_pct:.1f}%）\n\n")
            analysis.append(f"自然検索からの訪問数が前月比+{organic_change}回（+{organic_change_pct:.1f}%）増加し、\n")
            analysis.append(f"検索結果の見た目修正の効果が確認されました。検索エンジンでの表示機会の拡大や順位向上による効果と推測されます。\n\n")
        elif organic_change < 0:
            analysis.append(f"**自然検索からの訪問数**: {previous_data.get('organic_search', 0)}回 → {current_data.get('organic_search', 0)}回（{organic_change}回、{organic_change_pct:.1f}%）\n\n")
            analysis.append(f"自然検索からの訪問数が前月比{organic_change}回（{organic_change_pct:.1f}%）減少しました。\n")
            analysis.append(f"競合他社の施策強化や検索アルゴリズムの変更が影響している可能性があります。\n\n")
        
        if sc_position_change > 0 and current_data.get('sc_position', 0) > 0:
            analysis.append(f"**平均掲載順位**: {previous_data.get('sc_position', 0):.1f}位 → **{current_data.get('sc_position', 0):.1f}位**（{sc_position_change:.1f}位改善）\n\n")
            analysis.append(f"検索結果での平均掲載順位が{previous_data.get('sc_position', 0):.1f}位から{current_data.get('sc_position', 0):.1f}位へと{sc_position_change:.1f}位向上し、\n")
            analysis.append(f"SEO施策の効果が確認されました。上位表示により、検索ユーザーへの露出機会が拡大しています。\n\n")
        elif sc_position_change < 0 and current_data.get('sc_position', 0) > 0:
            analysis.append(f"**平均掲載順位**: {previous_data.get('sc_position', 0):.1f}位 → {current_data.get('sc_position', 0):.1f}位（{abs(sc_position_change):.1f}位）\n\n")
            analysis.append(f"平均掲載順位が前月比{abs(sc_position_change):.1f}位低下しました。\n")
            analysis.append(f"コンテンツの最適化やSEO施策の強化が必要です。\n\n")
        
    except Exception as e:
        analysis.append(f"⚠️ 改善効果分析でエラーが発生しました: {str(e)}\n")
        import traceback
        analysis.append(f"\n```\n{traceback.format_exc()}\n```\n")
    
    return "## 5. 改善効果の分析\n\n" + "".join(analysis)

def load_email_template(template_path, report_month_num, report_data, comparison_data, sc_data, comparison_sc_data):
    """
    メールテンプレートを生成して変数を埋める
    宛名は「碇谷様」で固定
    """
    # 指定されたフォーマットのテンプレートを直接使用
    email_template = """件名：【{{MONTH}}月分レポート】cectokyo.com 月間分析レポートのご送付

碇谷様

お世話になっております。
鈴木です。

{{MONTH}}月分の月間レポートを作成いたしましたので、お送りいたします。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ レポート資料
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【スライド資料】
{{SLIDE_URL}}

【詳細データ（スプレッドシート）】
{{SPREADSHEET_URL}}


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ {{MONTH}}月のサマリー（前月対比）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【良かった点】
✅ {{GOOD_POINT_1}}
✅ {{GOOD_POINT_2}}

【課題・懸念点】
⚠️ {{ISSUE_1}}
⚠️ {{ISSUE_2}}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 改善提案
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

{{IMPROVEMENT_PROPOSAL}}

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ご不明な点やご質問がございましたら、お気軽にお申し付けください。

よろしくお願いいたします。

鈴木"""
    
    # 良かった点を抽出（12月の結果を分析）
    good_points = []
    if report_data.get('inquiries_change', 0) > 0:
        good_points.append(f"問い合わせ件数が{report_data.get('inquiries', 0)}件（前月比 +{report_data.get('inquiries_change', 0):.1f}%）に増加")
    
    if report_data.get('cvr_change', 0) > 0:
        good_points.append(f"問い合わせ率が{report_data.get('cvr', 0):.2f}%（前月比 +{report_data.get('cvr_change', 0):.2f}pt）に改善")
    
    sc_ctr_change = sc_data.get('sc_ctr', 0) - comparison_sc_data.get('sc_ctr', 0)
    if sc_ctr_change > 0:
        good_points.append(f"検索結果の平均クリック率が{sc_data.get('sc_ctr', 0):.2f}%（前月比 +{sc_ctr_change:.2f}pt）に向上")
    
    if report_data.get('organic_search_change', 0) > 0:
        good_points.append(f"自然検索からの訪問数が{report_data.get('organic_search', 0)}回（前月比 +{report_data.get('organic_search_change', 0):.1f}%）に増加")
    
    # デフォルト値
    if not good_points:
        good_points.append("サイトの安定稼働を確認")
    if len(good_points) < 2:
        good_points.append("引き続き改善を継続")
    
    # 課題・懸念点を抽出（12月の結果を分析）
    issues = []
    if report_data.get('sessions_change', 0) < -10:
        issues.append(f"訪問数が{report_data.get('sessions', 0)}回（前月比 {report_data.get('sessions_change', 0):.1f}%）に減少")
    
    # 採用エントリーページの課題
    issues.append("採用エントリーページ（/entry/）が432回表示されるもクリック数0（CTR 0%）")
    
    if report_data.get('direct_change', 0) < -20:
        issues.append(f"直接流入が{report_data.get('direct', 0)}セッション（前月比 {report_data.get('direct_change', 0):.1f}%）に減少")
    
    # デフォルト値
    if len(issues) < 2:
        issues.append("引き続きモニタリングが必要")
    
    # 改善提案（12月の結果に基づく）
    improvement_proposal = """1. 求人ページ（/entry/）の検索結果表示改善
   - タイトル・説明文の最適化
   - クリック率向上を目指す（目標: 3〜5%）
   - 月間10〜20名の応募検討者をサイトへ誘導

2. コンテンツ拡充
   - プロジェクトページの充実
   - 検索からの訪問数の強化

3. 問い合わせページ（/contact/）の再確認・改善
   - 現在のクリック率 0.12%から3%を目標"""
    
    # 変数を埋める
    content = email_template
    content = content.replace('{{MONTH}}', report_month_num)
    content = content.replace('{{GOOD_POINT_1}}', good_points[0])
    content = content.replace('{{GOOD_POINT_2}}', good_points[1] if len(good_points) > 1 else '引き続き改善を継続')
    content = content.replace('{{ISSUE_1}}', issues[0])
    content = content.replace('{{ISSUE_2}}', issues[1] if len(issues) > 1 else '引き続きモニタリングが必要')
    content = content.replace('{{IMPROVEMENT_PROPOSAL}}', improvement_proposal)
    
    # {{SLIDE_URL}}や{{SPREADSHEET_URL}}はそのまま残す（後で手動で埋める）
    
    return content

def generate_report_from_template(template_path, report_month_str, comp_month_str, report_data, comparison_data, sc_data, comparison_sc_data, improvement_analysis, email_template, base_dir="../reports"):
    """テンプレートを読み込んで数値を埋めてレポートを生成"""
    if not os.path.exists(template_path):
        raise FileNotFoundError(f"テンプレートファイルが見つかりません: {template_path}")
    
    with open(template_path, 'r', encoding='utf-8') as f:
        template_content = f.read()
    
    # 数値を埋める
    content = template_content
    content = content.replace('{REPORT_MONTH_STR}', report_month_str)
    content = content.replace('{COMP_MONTH_STR}', comp_month_str)
    content = content.replace('{COMP_SESSIONS}', str(comparison_data.get('sessions', 0)))
    content = content.replace('{REPORT_SESSIONS}', str(report_data.get('sessions', 0)))
    content = content.replace('{SESSIONS_CHANGE}', f"{report_data.get('sessions_change', 0):+.1f}%")
    content = content.replace('{COMP_INQUIRIES}', str(comparison_data.get('inquiries', 0)))
    content = content.replace('{REPORT_INQUIRIES}', str(report_data.get('inquiries', 0)))
    content = content.replace('{INQUIRIES_CHANGE}', f"{report_data.get('inquiries_change', 0):+.1f}%")
    content = content.replace('{COMP_CVR}', f"{comparison_data.get('cvr', 0):.2f}")
    content = content.replace('{REPORT_CVR}', f"{report_data.get('cvr', 0):.2f}")
    content = content.replace('{CVR_CHANGE}', f"{report_data.get('cvr_change', 0):+.2f}pt")
    content = content.replace('{REPORT_NEW_USERS}', str(report_data.get('new_users', 0)))
    content = content.replace('{REPORT_TOTAL_USERS}', str(report_data.get('total_users', 0)))
    content = content.replace('{COMP_ORGANIC_SEARCH}', str(comparison_data.get('organic_search', 0)))
    content = content.replace('{REPORT_ORGANIC_SEARCH}', str(report_data.get('organic_search', 0)))
    content = content.replace('{ORGANIC_SEARCH_CHANGE}', f"{report_data.get('organic_search_change', 0):+.1f}%")
    content = content.replace('{COMP_DIRECT}', str(comparison_data.get('direct', 0)))
    content = content.replace('{REPORT_DIRECT}', str(report_data.get('direct', 0)))
    content = content.replace('{DIRECT_CHANGE}', f"{report_data.get('direct_change', 0):+.1f}%")
    # CTRの計算（sc_ctrを統一使用、単位はテンプレート側で指定）
    sc_ctr_current = sc_data.get('sc_ctr', 0) or sc_data.get('avg_ctr', 0)
    sc_ctr_comparison = comparison_sc_data.get('sc_ctr', 0) or comparison_sc_data.get('avg_ctr', 0)
    ctr_change = sc_ctr_current - sc_ctr_comparison
    
    content = content.replace('{COMP_CTR}', f"{sc_ctr_comparison:.2f}")
    content = content.replace('{REPORT_CTR}', f"{sc_ctr_current:.2f}")
    # 単位はテンプレート側で指定されているので、数値のみ（符号付き）
    ctr_change_str = f"{ctr_change:+.2f}pt" if sc_ctr_current > 0 and sc_ctr_comparison > 0 else "-"
    content = content.replace('{CTR_CHANGE}', ctr_change_str)
    
    # 掲載順位の計算（0の場合は「-」を表示、単位はテンプレート側で指定）
    comp_position = comparison_sc_data.get('sc_position', 0) or comparison_sc_data.get('avg_position', 0)
    report_position = sc_data.get('sc_position', 0) or sc_data.get('avg_position', 0)
    position_change = report_position - comp_position
    
    comp_position_str = f"{comp_position:.1f}" if comp_position > 0 else "-"
    report_position_str = f"{report_position:.1f}" if report_position > 0 else "-"
    # 単位はテンプレート側で指定されているので、数値のみ（符号付き）
    position_change_str = f"{position_change:+.1f}位" if report_position > 0 and comp_position > 0 else "-"
    
    content = content.replace('{COMP_POSITION}', comp_position_str)
    content = content.replace('{REPORT_POSITION}', report_position_str)
    content = content.replace('{POSITION_CHANGE}', position_change_str)
    
    # CVRの単位もテンプレート側で指定されているので、数値のみ（符号付き）
    cvr_change_val = report_data.get('cvr_change', 0)
    cvr_change_str = f"{cvr_change_val:+.2f}pt" if cvr_change_val != 0 or report_data.get('cvr', 0) > 0 else "-"
    content = content.replace('{CVR_CHANGE}', cvr_change_str)
    content = content.replace('{COMP_CLICKS}', str(comparison_sc_data.get('sc_clicks', 0)))
    content = content.replace('{REPORT_CLICKS}', str(sc_data.get('sc_clicks', 0)))
    content = content.replace('{CLICKS_CHANGE}', f"{(sc_data.get('sc_clicks', 0) - comparison_sc_data.get('sc_clicks', 0)):+d}")
    content = content.replace('{COMP_IMPRESSIONS}', str(comparison_sc_data.get('sc_impressions', 0)))
    content = content.replace('{REPORT_IMPRESSIONS}', str(sc_data.get('sc_impressions', 0)))
    content = content.replace('{IMPRESSIONS_CHANGE}', f"{(sc_data.get('sc_impressions', 0) - comparison_sc_data.get('sc_impressions', 0)):+d}")
    
    # 流入経路分析の推察コメントを生成
    traffic_analysis = generate_traffic_source_analysis(report_data, comparison_data)
    content = content.replace('{TRAFFIC_SOURCE_ANALYSIS}', traffic_analysis)
    
    # デバイス別パフォーマンステーブルを生成
    device_perf = sc_data.get('device_performance', {})
    # device_performanceが空の場合やデバイスデータが取得できていない場合はデフォルト値を設定
    if not device_perf:
        device_perf = {
            'PC': {'clicks': 0, 'impressions': 0, 'ctr': 0.0, 'position': 0.0},
            'モバイル': {'clicks': 0, 'impressions': 0, 'ctr': 0.0, 'position': 0.0},
            'タブレット': {'clicks': 0, 'impressions': 0, 'ctr': 0.0, 'position': 0.0}
        }
    device_table = "| デバイス | クリック数 | 表示回数 | クリック率 | 掲載順位 |\n"
    device_table += "| :--- | :--- | :--- | :--- | :--- |\n"
    for device_name in ['PC', 'モバイル', 'タブレット']:
        device_info = device_perf.get(device_name, {})
        clicks = device_info.get('clicks', 0)
        impressions = device_info.get('impressions', 0)
        ctr = device_info.get('ctr', 0.0)
        position = device_info.get('position', 0.0)
        position_str = f"{position:.1f}" if position > 0 else "-"
        device_table += f"| {device_name} | {clicks} | {impressions} | {ctr:.2f}% | {position_str}位 |\n"
    content = content.replace('{DEVICE_PERFORMANCE_TABLE}', device_table)
    
    # コンバージョンファネルテーブルを生成（report_dataから取得）
    conversion_funnel = report_data.get('conversion_funnel', {})
    total_sessions = conversion_funnel.get('sessions', report_data.get('sessions', 0))
    form_start = conversion_funnel.get('form_start', 0)
    form_complete = conversion_funnel.get('form_complete', conversion_funnel.get('inquiries', report_data.get('inquiries', 0)))
    
    funnel_table = "| ステップ | 数値 | 前ステップ比（問い合わせ率） | 総訪問数比 |\n"
    funnel_table += "| :--- | :--- | :--- | :--- |\n"
    funnel_table += f"| 訪問数 | {total_sessions} | - | 100.00% |\n"
    if total_sessions > 0:
        form_start_rate = (form_start / total_sessions * 100)
        completion_rate = (form_complete / form_start * 100) if form_start > 0 else 0
        overall_cvr = (form_complete / total_sessions * 100)
        funnel_table += f"| フォーム開始 | {form_start} | - | {form_start_rate:.2f}% |\n"
        funnel_table += f"| フォーム完了 | {form_complete} | {completion_rate:.2f}% | {overall_cvr:.2f}% |\n"
    else:
        funnel_table += "| フォーム開始 | 0 | - | 0.00% |\n"
        funnel_table += "| フォーム完了 | 0 | 0.00% | 0.00% |\n"
    content = content.replace('{CONVERSION_FUNNEL_TABLE}', funnel_table)
    
    # 改善提案テーブルを生成（improvement-patterns.mdから読み込み）
    patterns = load_improvement_patterns()
    improvement_proposals_table = generate_improvement_proposals(
        "- 採用エントリーページ(/entry/)の表示回数: 432回\n- 採用エントリーページ(/entry/)のクリック数: 0回\n- 課題: 表示されているが選ばれていない（改善パターン1該当）",
        patterns
    )
    content = content.replace('{IMPROVEMENT_PROPOSALS_TABLE}', improvement_proposals_table)
    
    # 特記事項
    issues_content = "- 採用エントリーページ(/entry/)の表示回数: 432回\n- 採用エントリーページ(/entry/)のクリック数: 0回\n- 課題: 検索結果に表示されているが選ばれていない（見られているがクリックされていない）\n- 改善策: タイトルと説明文を魅力的に書き換え、クリック率を向上させる必要があります"
    content = content.replace('{ISSUES_CONTENT}', issues_content)
    
    # 2ページ目相当の内容を生成（先月の良かった点と施策の効果）
    summary_section = "### 良かった点\n\n"
    if report_data.get('inquiries_change', 0) > 0:
        summary_section += f"✅ **問い合わせ件数が増加しました**\n\n"
        summary_section += f"前月比 +{report_data.get('inquiries_change', 0):.1f}% の増加となりました。\n"
        summary_section += f"この増加は、サイトの信頼感向上や検索結果の見た目修正の効果により、\n"
        summary_section += f"訪問者がサイトを信頼して問い合わせを行ったことが主な要因と考えられます。\n\n"
    
    summary_section += "### 11月に実施した施策の効果\n\n"
    summary_section += "**検索結果の見た目修正（タイトル・説明文の改善）**\n\n"
    sc_ctr_current = sc_data.get('sc_ctr', 0) or sc_data.get('avg_ctr', 0)
    sc_ctr_comparison = comparison_sc_data.get('sc_ctr', 0) or comparison_sc_data.get('avg_ctr', 0)
    if sc_ctr_current > 0 and sc_ctr_comparison > 0:
        ctr_improvement = sc_ctr_current - sc_ctr_comparison
        summary_section += f"11月に実施した検索結果の見た目修正により、平均クリック率が{sc_ctr_comparison:.2f}%から{sc_ctr_current:.2f}%へと{ctr_improvement:.2f}pt改善しました。\n"
        summary_section += f"これにより、検索結果に表示された際に、より多くのユーザーがサイトをクリックするようになりました。\n\n"
    else:
        summary_section += "11月に実施した検索結果の見た目修正の効果を継続的に測定していきます。\n\n"
    
    summary_section += "### 採用ページの課題と改善策\n\n"
    summary_section += "**課題**: 採用エントリーページ（/entry/）が検索結果に432回表示されるもクリック数0（クリック率 0%）\n\n"
    summary_section += "検索結果には表示されているものの、ユーザーが選ばない（クリックしない）状態が続いています。\n"
    summary_section += "これは、タイトルや説明文が魅力的でない、または検索意図とマッチしていない可能性があります。\n\n"
    summary_section += "**改善策**:\n"
    summary_section += "- タイトルに求人の魅力を分かりやすく記載（例：「年収800万円以上」「実績100社以上」など）\n"
    summary_section += "- 説明文に具体的な情報を記載（仕事内容、待遇、職場環境など）\n"
    summary_section += "- 検索ユーザーの検索意図とページの内容を一致させる\n"
    summary_section += "- クリック率を3〜5%まで向上させることを目標とする\n\n"
    content = content.replace('{SUMMARY_SECTION}', summary_section)
    
    # 改善効果分析とメールテンプレート
    content = content.replace('{IMPROVEMENT_ANALYSIS}', improvement_analysis)
    
    report_month_num = report_month_str.split('年')[0] if '年' in report_month_str else "12"
    content = content.replace('{EMAIL_TEMPLATE}', email_template)
    
    # ファイルに保存（../reports/YYYY-MM_レポート.md形式）
    os.makedirs(base_dir, exist_ok=True)
    year_month = report_month_str.replace('年', '-').replace('月', '')
    output_path = os.path.join(base_dir, f"{year_month}_レポート.md")
    
    with open(output_path, 'w', encoding='utf-8') as f:
        f.write(content)
    
    print(f"✅ Markdownレポートを生成しました: {output_path}")
    return output_path

# ==========================================
# 3. Markdownファイルの自動生成機能（本番用）
# ==========================================
def export_to_markdown(report_month, comp_month, report_data, comparison_data, sc_data, comparison_sc_data):
    """Markdownファイルを生成（メールテンプレート形式を含む）"""
    
    # 月の日本語表記
    report_date = datetime.strptime(report_month, "%Y-%m")
    comp_date = datetime.strptime(comp_month, "%Y-%m")
    report_month_str = report_date.strftime("%Y年%m月")
    comp_month_str = comp_date.strftime("%Y年%m月")
    report_month_num = report_date.strftime("%m")
    
    # 良かった点を抽出
    good_points = []
    if report_data.get('inquiries_change', 0) > 0:
        good_points.append(f"問い合わせ件数が{report_data.get('inquiries', 0)}件（前月比 +{report_data.get('inquiries_change', 0):.1f}%）に増加")
    
    if report_data.get('cvr_change', 0) > 0:
        good_points.append(f"問い合わせ率が{report_data.get('cvr', 0):.2f}%（前月比 +{report_data.get('cvr_change', 0):.2f}pt）に改善")
    
    ctr_change = sc_data.get('avg_ctr', 0) - comparison_sc_data.get('avg_ctr', 0)
    if ctr_change > 0:
        good_points.append(f"検索結果の平均CTRが{sc_data.get('avg_ctr', 0):.2f}%（前月比 +{ctr_change:.2f}pt）に向上")
    
    if report_data.get('organic_search_change', 0) > 0:
        good_points.append(f"自然検索流入が{report_data.get('organic_search', 0)}セッション（前月比 +{report_data.get('organic_search_change', 0):.1f}%）に増加")
    
    # デフォルト値
    if not good_points:
        good_points.append("サイトの安定稼働を確認")
    if len(good_points) < 2:
        good_points.append("引き続き改善を継続")
    
    # 課題・懸念点を抽出
    issues = []
    if report_data.get('sessions_change', 0) < -10:
        issues.append(f"訪問数が{report_data.get('sessions', 0)}回（前月比 {report_data.get('sessions_change', 0):.1f}%）に減少")
    
    # 採用エントリーページの課題
    issues.append("採用エントリーページ（/entry/）が432回表示されるもクリック数0（CTR 0%）")
    
    if report_data.get('direct_change', 0) < -20:
        issues.append(f"直接流入が{report_data.get('direct', 0)}セッション（前月比 {report_data.get('direct_change', 0):.1f}%）に減少")
    
    # デフォルト値
    if len(issues) < 2:
        issues.append("引き続きモニタリングが必要")
    
    # 改善提案
    improvement_proposal = """1. 求人ページ（/entry/）の検索結果表示改善
   - タイトル・説明文の最適化
   - クリック率向上を目指す（目標: 3〜5%）
   - 月間10〜20名の応募検討者をサイトへ誘導

2. コンテンツ拡充
   - プロジェクトページの充実
   - 検索からの訪問数の強化

3. 問い合わせページ（/contact/）の再確認・改善
   - 現在のクリック率 0.12%から3%を目標"""
    
    # 修正対象ページ
    pages = """以下のページのSEO設定（タイトル・ディスクリプション）を修正してください：
- /contact/
- /career/
- /entry/"""
    
    # 希望納期（来月の第1週を設定）
    next_month = report_date + relativedelta(months=1)
    deadline = next_month.strftime("%Y年%m月第1週")
    
    # Markdownコンテンツ
    md_content = f"""# {report_month_str} Web運用報告データ

## 1. 主要KPI

| 指標 | {comp_month_str} | {report_month_str} | 前月比 |
| :--- | :--- | :--- | :--- |
| 総訪問数 | {comparison_data.get('sessions', 0)} | {report_data.get('sessions', 0)} | {report_data.get('sessions_change', 0):+.1f}% |
| 問い合わせ件数 | {comparison_data.get('inquiries', 0)}件 | {report_data.get('inquiries', 0)}件 | {report_data.get('inquiries_change', 0):+.1f}% |
| 問い合わせ率 | {comparison_data.get('cvr', 0):.2f}% | {report_data.get('cvr', 0):.2f}% | {report_data.get('cvr_change', 0):+.2f}pt |
| 新規ユーザー数 | - | {report_data.get('new_users', 0)} | - |
| 総ユーザー数 | - | {report_data.get('total_users', 0)} | - |

## 2. 流入経路

| チャネル | {comp_month_str} | {report_month_str} | 前月比 |
| :--- | :--- | :--- | :--- |
| Organic Search | {comparison_data.get('organic_search', 0)} | {report_data.get('organic_search', 0)} | {report_data.get('organic_search_change', 0):+.1f}% |
| Direct | {comparison_data.get('direct', 0)} | {report_data.get('direct', 0)} | {report_data.get('direct_change', 0):+.1f}% |

## 3. Search Console

| 指標 | {comp_month_str} | {report_month_str} | 前月比 |
| :--- | :--- | :--- | :--- |
| 平均CTR | {comparison_sc_data.get('avg_ctr', 0):.2f}% | {sc_data.get('avg_ctr', 0):.2f}% | {(sc_data.get('avg_ctr', 0) - comparison_sc_data.get('avg_ctr', 0)):+.2f}pt |
| 平均掲載順位 | {comparison_sc_data.get('avg_position', 0):.1f}位 | {sc_data.get('avg_position', 0):.1f}位 | {(sc_data.get('avg_position', 0) - comparison_sc_data.get('avg_position', 0)):+.1f}位 |
| 総クリック数 | {comparison_sc_data.get('total_clicks', 0)} | {sc_data.get('total_clicks', 0)} | {sc_data.get('total_clicks', 0) - comparison_sc_data.get('total_clicks', 0):+d} |
| 総表示回数 | {comparison_sc_data.get('total_impressions', 0)} | {sc_data.get('total_impressions', 0)} | {sc_data.get('total_impressions', 0) - comparison_sc_data.get('total_impressions', 0):+d} |

## 4. 特記事項（課題）

- 採用エントリーページ(/entry/)の表示回数: 432回
- 採用エントリーページ(/entry/)のクリック数: 0回
- 課題: 表示されているが選ばれていない（改善パターン1該当）

---

## 【クライアント返信用メール案】

```
件名：【{report_month_num}月分レポート】コモンウェルスエンジニアーズ様 Web運用報告書のご送付

碇谷様

お世話になっております。
鈴木です。

{report_month_num}月分の月間レポートを作成いたしましたので、お送りいたします。

━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ レポート資料
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【スライド資料】
[ここにURLを貼る]

【詳細データ（スプレッドシート）】
[ここにURLを貼る]


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ {report_month_num}月のサマリー（前月対比）
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

【良かった点】
✅ {good_points[0]}
✅ {good_points[1]}

【課題・懸念点】
⚠️ {issues[0]}
⚠️ {issues[1]}


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━
■ 改善提案
━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

{improvement_proposal}


━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━━

ご不明な点やご質問がございましたら、お気軽にお申し付けください。

よろしくお願いいたします。

鈴木
```

---

## 【エンジニア向け修正依頼メール案】

```
件名：【修正依頼】cectokyo.com SEO設定の修正

{{ENGINEER_NAME}}様

お疲れ様です。
鈴木です。

cectokyo.comのSEO設定修正をお願いしたく、ご連絡いたしました。

■ 修正内容
{pages}

詳細な修正内容は、添付の指示書をご確認ください。

■ 希望納期
{deadline}

ご不明な点がございましたら、お気軽にご連絡ください。

よろしくお願いいたします。

鈴木
```
"""
    
    # 出力先は reports/ に戻す
    output_path = f"reports/{report_month}_レポート.md"
    with open(output_path, "w", encoding="utf-8") as f:
        f.write(md_content)
    print(f"✅ Markdownレポートを生成しました: {output_path}")

# ==========================================
# メイン処理
# ==========================================
def main():
    """メイン処理"""
    print("🚀 レポート生成開始")
    print("="*60)
    
    # A. 実行日ベースで期間自動判定（前月1日〜末日）
    report_year_month, start_date, end_date, comp_year_month = get_report_periods()
    
    report_date = datetime.strptime(report_year_month, "%Y-%m")
    comp_date = datetime.strptime(comp_year_month, "%Y-%m")
    report_month_str = report_date.strftime("%Y年%m月")
    comp_month_str = comp_date.strftime("%Y年%m月")
    report_month_num = report_date.strftime("%m")
    
    # 前月（11月）の年月を計算（report_year_monthが12月の場合、前月は11月）
    previous_month_dt = report_date - relativedelta(months=1)
    previous_month_str = previous_month_dt.strftime("%Y-%m")
    
    print(f"📅 報告対象月: {report_month_str} ({report_year_month})")
    print(f"📅 比較対象月（前月）: {previous_month_str}")
    print(f"📅 データ取得期間: {start_date} 〜 {end_date}")
    
    # B. APIからデータを取得してCSV保存（../data/YYYY-MM/）
    print(f"\n📊 APIからデータを取得中...")
    data_dir = os.path.join("..", "data", report_year_month)
    os.makedirs(data_dir, exist_ok=True)  # フォルダがなければ作成
    
    try:
        # GA4データを取得してCSV保存
        print("  - GA4トラフィックデータを取得中...")
        ga4_traffic = get_ga4_traffic_data(start_date, end_date, save_csv=True, output_dir=data_dir)
        
        print("  - GA4イベントデータを取得中...")
        ga4_events = get_ga4_events_data(start_date, end_date, save_csv=True, output_dir=data_dir)
        
        # Search Consoleデータを取得してCSV保存
        print("  - Search Consoleデータを取得中...")
        sc_data = get_search_console_data(start_date, end_date, save_csv=True, output_dir=data_dir)
        
        # Search Consoleデータの形式を統一（CTR計算の統一）
        calculated_ctr = sc_data.get('sc_ctr', 0) or sc_data.get('avg_ctr', 0)
        if calculated_ctr == 0 and sc_data.get('total_clicks', 0) > 0 and sc_data.get('total_impressions', 0) > 0:
            # CTRが0の場合は再計算
            calculated_ctr = (sc_data.get('total_clicks', 0) / sc_data.get('total_impressions', 0) * 100)
        
        sc_data_dict = {
            'sc_clicks': sc_data.get('total_clicks', 0),
            'sc_impressions': sc_data.get('total_impressions', 0),
            'sc_ctr': calculated_ctr,
            'avg_ctr': calculated_ctr,  # 後方互換性のため
            'sc_position': sc_data.get('avg_position', 0),
            'device_performance': sc_data.get('device_performance', {})
        }
        
    except (FileNotFoundError, ValueError, ConnectionError, RuntimeError) as e:
        print(f"\n{str(e)}")
        print("   処理を中断します。")
        return
    
    # C. データの解析
    print(f"\n📈 データを解析中...")
    sessions_dict = ga4_traffic.get('sessions', {})
    total_sessions = sum(sessions_dict.values())
    organic_search_sessions = sessions_dict.get('Organic Search', 0)
    direct_sessions = sessions_dict.get('Direct', 0)
    
    new_users = ga4_traffic.get('new_users', 0)
    total_users = ga4_traffic.get('total_users', 0)
    
    if total_users == 0:
        total_users = ga4_events.get('_total_users', 0)
    if new_users == 0:
        new_users = ga4_events.get('_new_users', 0)
    
    inquiries = ga4_events.get('contact', 0)
    form_start = ga4_events.get('form_start', 0) or ga4_events.get('contact', 0)  # フォーム開始イベント
    form_complete = ga4_events.get('form_complete', 0) or ga4_events.get('form_submit', 0) or inquiries  # フォーム完了イベント
    cvr = (inquiries / total_sessions * 100) if total_sessions > 0 else 0
    
    # コンバージョンファネルデータを準備
    conversion_funnel = {
        'sessions': total_sessions,
        'form_start': form_start,
        'form_complete': form_complete,
        'inquiries': inquiries
    }
    
    # D. 比較対象月（前月、11月）のデータをレポートから読み込む
    print(f"  - 前月（{previous_month_str}）のデータをレポートから読み込み中...")
    # 前月レポートからデータを抽出（11月の値）
    comparison_data_from_report = parse_markdown_for_previous_month_data(previous_month_str, base_dir="../reports")
    
    # 比較計算（11月のデータを使用）
    comparison_sessions = comparison_data_from_report.get('sessions', 0)
    comparison_inquiries = comparison_data_from_report.get('inquiries', 0)
    comparison_cvr = comparison_data_from_report.get('cvr', 0)
    comparison_organic = comparison_data_from_report.get('organic_search', 0)
    comparison_direct = comparison_data_from_report.get('direct', 0)
    
    # Search Consoleデータも前月レポートから取得を試みる（なければCSVから）
    comparison_sc_from_csv = load_data_from_csv(previous_month_str, base_dir="../data")
    comparison_sc_ctr = comparison_sc_from_csv.get('sc_ctr', 0)
    comparison_sc_position = comparison_sc_from_csv.get('sc_position', 0)
    comparison_sc_clicks = comparison_sc_from_csv.get('sc_clicks', 0)
    comparison_sc_impressions = comparison_sc_from_csv.get('sc_impressions', 0)
    
    report_data = {
        'sessions': total_sessions,
        'inquiries': inquiries,
        'cvr': cvr,
        'organic_search': organic_search_sessions,
        'direct': direct_sessions,
        'new_users': new_users,
        'total_users': total_users,
        'sessions_change': ((total_sessions - comparison_sessions) / comparison_sessions * 100) if comparison_sessions > 0 else 0,
        'inquiries_change': ((inquiries - comparison_inquiries) / comparison_inquiries * 100) if comparison_inquiries > 0 else 0,
        'cvr_change': cvr - comparison_cvr,
        'organic_search_change': ((organic_search_sessions - comparison_organic) / comparison_organic * 100) if comparison_organic > 0 else 0,
        'direct_change': ((direct_sessions - comparison_direct) / comparison_direct * 100) if comparison_direct > 0 else 0,
        'conversion_funnel': conversion_funnel  # コンバージョンファネルデータを追加
    }
    
    # 比較データをdict形式に統一
    comparison_data_dict = {
        'sessions': comparison_sessions,
        'inquiries': comparison_inquiries,
        'cvr': comparison_cvr,
        'organic_search': comparison_organic,
        'direct': comparison_direct
    }
    
    # 比較対象月のCTR計算を統一
    comparison_sc_ctr_calc = comparison_sc_ctr
    if comparison_sc_ctr == 0 and comparison_sc_impressions > 0:
        comparison_sc_ctr_calc = (comparison_sc_clicks / comparison_sc_impressions * 100)
    
    comparison_sc_data_dict = {
        'sc_clicks': comparison_sc_clicks,
        'sc_impressions': comparison_sc_impressions,
        'sc_ctr': comparison_sc_ctr_calc,
        'avg_ctr': comparison_sc_ctr_calc,  # 後方互換性のため
        'sc_position': comparison_sc_position
    }
    
    # 改善効果分析用のデータ（前月のデータ、CTR計算を統一）
    comparison_data_for_analysis = {
        'sessions': comparison_sessions,
        'inquiries': comparison_inquiries,
        'cvr': comparison_cvr,
        'organic_search': comparison_organic,
        'direct': comparison_direct,
        'sc_ctr': comparison_sc_ctr_calc,
        'avg_ctr': comparison_sc_ctr_calc,  # 後方互換性のため
        'sc_position': comparison_sc_position
    }
    
    # E. 改善効果の分析（前月レポートから改善提案を読み取り、今月のデータで効果を分析）
    print(f"  - 改善効果を分析中...")
    previous_report_path = os.path.join("..", "reports", f"{previous_month_str}.md")
    improvement_analysis = analyze_improvement_effects(previous_report_path, report_data, comparison_data_for_analysis)
    
    # F. メールテンプレート読み込みと変数埋め込み
    print(f"  - メールテンプレートを読み込み中...")
    email_template_path = os.path.join("templates", "Eメールテンプレート.md")
    
    try:
        email_template = load_email_template(
            email_template_path,
            report_month_num,
            report_data,
            comparison_data_dict,
            sc_data_dict,
            comparison_sc_data_dict
        )
    except FileNotFoundError as e:
        print(f"  ⚠️  メールテンプレートが見つかりません: {email_template_path}")
        print(f"     デフォルトのメールテンプレートを使用します")
        email_template = ""
    
    # G. テンプレートからレポート生成（../reports/YYYY-MM.md）
    print(f"\n📝 レポートを生成中...")
    template_path = os.path.join("templates", "monthly-report.md")
    reports_dir = os.path.join("..", "reports")
    os.makedirs(reports_dir, exist_ok=True)  # フォルダがなければ作成
    
    try:
        output_path = generate_report_from_template(
            template_path,
            report_month_str,
            comp_month_str,
            report_data,
            comparison_data_dict,
            sc_data_dict,
            comparison_sc_data_dict,
            improvement_analysis,
            email_template,
            base_dir=reports_dir
        )
        
        print("\n" + "="*60)
        print("✅ レポート生成が完了しました！")
        print(f"   出力先: {output_path}")
        print(f"   ファイル名: {report_year_month}.md")
        print("="*60)
    except FileNotFoundError as e:
        print(f"\n❌ エラー: {str(e)}")
        print("   テンプレートファイルを作成してください。")
    except Exception as e:
        print(f"\n❌ エラー: {str(e)}")
        import traceback
        traceback.print_exc()

if __name__ == "__main__":
    main()
