#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
GA4とSearch Consoleの接続テストプログラム

このプログラムは、設定ファイル（config.json）の情報を使って
GA4とSearch ConsoleのAPIに接続できるかテストします。
"""

import json
import os
import time
import requests
import jwt


def load_config():
    """設定ファイルを読み込む"""
    config_path = os.path.join(os.path.dirname(__file__), 'config.json')
    
    if not os.path.exists(config_path):
        print("❌ エラー: config.json が見つかりません")
        print(f"   {config_path} が存在するか確認してください")
        print("\n💡 ヒント: config.json.template をコピーして config.json を作成してください")
        print("   cp config.json.template config.json")
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


def test_ga4_connection(config):
    """GA4への接続をテストする"""
    print("\n" + "="*60)
    print("🔍 GA4接続テスト")
    print("="*60)
    
    ga4_config = config.get('ga4', {})
    property_id = ga4_config.get('property_id')
    credentials_file = ga4_config.get('credentials_file')
    
    # プロパティIDの確認
    if not property_id or property_id == "YOUR_GA4_PROPERTY_ID":
        print("❌ エラー: GA4のプロパティIDが設定されていません")
        print("   config.json の 'ga4.property_id' を設定してください")
        print("\n💡 プロパティIDの見つけ方:")
        print("   1. https://analytics.google.com/ にアクセス")
        print("   2. 左側メニューから「管理」をクリック")
        print("   3. 「プロパティ」列の「プロパティ設定」をクリック")
        print("   4. 「プロパティID」（数字のみ）をコピー")
        return False
    
    print(f"✅ プロパティID: {property_id}")
    
    # 認証情報ファイルの確認
    creds_path = os.path.join(os.path.dirname(__file__), credentials_file)
    if not os.path.exists(creds_path):
        print(f"❌ エラー: 認証情報ファイルが見つかりません: {credentials_file}")
        print(f"   {creds_path} が存在するか確認してください")
        return False
    
    print(f"✅ 認証情報ファイル: {credentials_file}")
    
    try:
        # アクセストークンを取得
        scopes = ['https://www.googleapis.com/auth/analytics.readonly']
        print("\n📡 認証中...")
        access_token = get_access_token(credentials_file, scopes)
        
        # GA4 APIに接続してテスト
        print("📡 APIに接続中...")
        api_url = f"https://analyticsdata.googleapis.com/v1beta/properties/{property_id}:runReport"
        
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        # テストリクエスト
        request_body = {
            "dateRanges": [{
                "startDate": "2024-01-01",
                "endDate": "2024-01-31"
            }],
            "metrics": [{"name": "sessions"}],
            "dimensions": [{"name": "date"}]
        }
        
        response = requests.post(api_url, headers=headers, json=request_body)
        response.raise_for_status()
        
        data = response.json()
        row_count = len(data.get('rows', []))
        
        print("✅ GA4への接続に成功しました！")
        print(f"   テストデータ取得: {row_count} 行")
        return True
        
    except requests.exceptions.HTTPError as e:
        error_detail = ""
        try:
            error_json = e.response.json()
            error_detail = error_json.get('error', {}).get('message', str(e))
        except:
            error_detail = str(e)
        print(f"❌ GA4への接続に失敗しました: {error_detail}")
        print("\n💡 トラブルシューティング:")
        print("   1. credentials.json が正しいか確認")
        print("   2. GA4でサービスアカウントにアクセス権限があるか確認")
        print("   3. プロパティIDが正しいか確認")
        return False
    except Exception as e:
        print(f"❌ GA4への接続に失敗しました: {str(e)}")
        print("\n💡 トラブルシューティング:")
        print("   1. credentials.json が正しいか確認")
        print("   2. GA4でサービスアカウントにアクセス権限があるか確認")
        print("   3. プロパティIDが正しいか確認")
        return False


def test_search_console_connection(config):
    """Search Consoleへの接続をテストする"""
    print("\n" + "="*60)
    print("🔍 Search Console接続テスト")
    print("="*60)
    
    sc_config = config.get('search_console', {})
    site_url = sc_config.get('site_url')
    credentials_file = sc_config.get('credentials_file')
    
    # サイトURLの確認
    if not site_url or site_url == "sc-domain:YOUR_SITE":
        print("❌ エラー: Search ConsoleのサイトURLが設定されていません")
        print("   config.json の 'search_console.site_url' を設定してください")
        print("\n💡 サイトURLの形式:")
        print("   - sc-domain:cectokyo.com (ドメインプロパティの場合)")
        print("   - https://cectokyo.com/ (URLプレフィックスプロパティの場合)")
        return False
    
    print(f"✅ サイトURL: {site_url}")
    
    # 認証情報ファイルの確認
    creds_path = os.path.join(os.path.dirname(__file__), credentials_file)
    if not os.path.exists(creds_path):
        print(f"❌ エラー: 認証情報ファイルが見つかりません: {credentials_file}")
        print(f"   {creds_path} が存在するか確認してください")
        return False
    
    print(f"✅ 認証情報ファイル: {credentials_file}")
    
    try:
        # アクセストークンを取得
        scopes = ['https://www.googleapis.com/auth/webmasters.readonly']
        print("\n📡 認証中...")
        access_token = get_access_token(credentials_file, scopes)
        
        # Search Console APIに接続してテスト
        print("📡 APIに接続中...")
        
        # サイト一覧を取得
        sites_url = "https://www.googleapis.com/webmasters/v3/sites"
        headers = {
            'Authorization': f'Bearer {access_token}',
            'Content-Type': 'application/json'
        }
        
        response = requests.get(sites_url, headers=headers)
        response.raise_for_status()
        
        sites_data = response.json()
        
        # 指定されたサイトが存在するか確認
        site_found = False
        if 'siteEntry' in sites_data:
            for site in sites_data['siteEntry']:
                if site['siteUrl'] == site_url:
                    site_found = True
                    print(f"✅ Search Consoleへの接続に成功しました！")
                    print(f"   サイト権限: {site.get('permissionLevel', '不明')}")
                    break
        
        if not site_found:
            print(f"⚠️  警告: {site_url} が見つかりませんでした")
            print("   利用可能なサイト:")
            if 'siteEntry' in sites_data:
                for site in sites_data['siteEntry']:
                    print(f"     - {site['siteUrl']}")
            print("\n💡 ヒント: config.json の site_url を確認してください")
            return False
        
        return True
        
    except requests.exceptions.HTTPError as e:
        error_detail = ""
        try:
            error_json = e.response.json()
            error_detail = error_json.get('error', {}).get('message', str(e))
        except:
            error_detail = str(e)
        print(f"❌ Search Consoleへの接続に失敗しました: {error_detail}")
        print("\n💡 トラブルシューティング:")
        print("   1. credentials.json が正しいか確認")
        print("   2. Search Consoleでサービスアカウントに所有者権限があるか確認")
        print("   3. サイトURLが正しいか確認")
        return False
    except Exception as e:
        print(f"❌ エラーが発生しました: {str(e)}")
        return False


def main():
    """メイン処理"""
    print("🚀 GA4 & Search Console 接続テスト開始")
    print("="*60)
    
    # 設定ファイルを読み込む
    config = load_config()
    if not config:
        return
    
    # GA4接続テスト
    ga4_ok = test_ga4_connection(config)
    
    # Search Console接続テスト
    sc_ok = test_search_console_connection(config)
    
    # 結果サマリー
    print("\n" + "="*60)
    print("📊 テスト結果サマリー")
    print("="*60)
    print(f"GA4接続: {'✅ 成功' if ga4_ok else '❌ 失敗'}")
    print(f"Search Console接続: {'✅ 成功' if sc_ok else '❌ 失敗'}")
    
    if ga4_ok and sc_ok:
        print("\n🎉 すべての接続テストに成功しました！")
        print("   次は fetch_data.py を実行してデータを取得できます")
    else:
        print("\n⚠️  一部の接続に失敗しました")
        print("   上記のエラーメッセージとトラブルシューティングを確認してください")


if __name__ == "__main__":
    main()
