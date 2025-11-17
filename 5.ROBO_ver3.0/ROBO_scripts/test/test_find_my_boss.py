"""
test_find_my_boss.py
指定ユーザーの上司（manager）情報を Microsoft Graph から取得します（Python 版）。

認証方式: デバイスコード（MSAL）
必要な環境変数:
  - MSAL_CLIENT_ID: Azure AD(Entra ID) 上の公開クライアントアプリのクライアントID
  - MSAL_TENANT_ID: テナントID（例: 00000000-0000-0000-0000-000000000000）

必要な Graph API の委任権限（Delegated Permissions）:
  - User.Read.All（管理者同意が必要）
  - Directory.Read.All（管理者同意が必要）

使い方:
  pip install msal requests
  set MSAL_CLIENT_ID=＜クライアントID＞
  set MSAL_TENANT_ID=＜テナントID＞
  python .\\test_find_my_boss.py --user okada.kazuhito@jp.panasonic.com

注: 実行時に表示される URL とコードでブラウザからサインインしてください。
    初回は管理者同意が必要になることがあります。
"""

from __future__ import annotations

import argparse
import os
import sys
from typing import Any, Dict, Optional

import requests
from msal import PublicClientApplication


GRAPH_RESOURCE = "https://graph.microsoft.com"
GRAPH_SCOPE_DELEGATED = ["User.Read.All", "Directory.Read.All"]


def get_env(name: str) -> str:
    v = os.getenv(name)
    if not v:
        print(f"環境変数 {name} が未設定です。", file=sys.stderr)
        sys.exit(2)
    return v


def acquire_token_interactive(client_id: str, tenant_id: str) -> str:
    authority = f"https://login.microsoftonline.com/{tenant_id}"
    app = PublicClientApplication(client_id=client_id, authority=authority)

    # まずキャッシュからトークンを探す
    accounts = app.get_accounts()
    if accounts:
        result = app.acquire_token_silent(GRAPH_SCOPE_DELEGATED, account=accounts[0])
        if result and "access_token" in result:
            return result["access_token"]

    # デバイスコードで取得
    flow = app.initiate_device_flow(scopes=[f"{GRAPH_RESOURCE}/.default", *GRAPH_SCOPE_DELEGATED])
    if "user_code" not in flow:
        raise RuntimeError("デバイスコードの初期化に失敗しました。")
    print(flow["message"])  # サインイン手順を表示
    result = app.acquire_token_by_device_flow(flow)
    if "access_token" not in result:
        raise RuntimeError(f"トークン取得に失敗: {result}")
    return result["access_token"]


def get_user_manager(token: str, user_id: str) -> Dict[str, Any]:
    url = f"{GRAPH_RESOURCE}/v1.0/users/{user_id}/manager"
    params = {"$select": "displayName,mail,userPrincipalName,id,jobTitle,businessPhones,mobilePhone"}
    resp = requests.get(url, headers={"Authorization": f"Bearer {token}"}, params=params, timeout=30)
    if resp.status_code == 404:
        raise RuntimeError("ユーザーまたは manager が見つかりませんでした (404)。")
    if resp.status_code == 403:
        raise RuntimeError("権限不足です (403)。管理者に 'User.Read.All' / 'Directory.Read.All' の同意を依頼してください。")
    if not resp.ok:
        raise RuntimeError(f"Graph 呼び出し失敗: {resp.status_code} {resp.text}")
    return resp.json()


def main(argv: Optional[list[str]] = None) -> int:
    parser = argparse.ArgumentParser(description="Lookup a user's manager via Microsoft Graph")
    parser.add_argument("--user", required=True, help="UserPrincipalName or object id")
    args = parser.parse_args(argv)

    client_id = get_env("MSAL_CLIENT_ID")
    tenant_id = get_env("MSAL_TENANT_ID")

    token = acquire_token_interactive(client_id, tenant_id)
    data = get_user_manager(token, args.user)

    # 整形して表示
    def prop(d: Dict[str, Any], k: str) -> str:
        v = d.get(k)
        if isinstance(v, list):
            return ", ".join(v)
        return v or ""

    print("Manager:")
    print(f"  DisplayName       : {prop(data, 'displayName')}")
    print(f"  Mail              : {prop(data, 'mail')}")
    print(f"  UserPrincipalName : {prop(data, 'userPrincipalName')}")
    print(f"  Id                : {prop(data, 'id')}")
    print(f"  JobTitle          : {prop(data, 'jobTitle')}")
    print(f"  BusinessPhones    : {prop(data, 'businessPhones')}")
    print(f"  MobilePhone       : {prop(data, 'mobilePhone')}")
    return 0


if __name__ == "__main__":
    raise SystemExit(main())

