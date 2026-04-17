"""
SharePoint client - Device Code Flow. Token cached after first login.
"""
from urllib.parse import urlparse
from typing import Dict, List
import os, msal, requests

class SharePointClient:
    GRAPH_BASE = "https://graph.microsoft.com/v1.0"
    SCOPES = ["https://graph.microsoft.com/Sites.Read.All","https://graph.microsoft.com/Files.Read.All","User.Read"]
    TOKEN_CACHE_FILE = "token_cache.json"

    def __init__(self, tenant_id, client_id, client_secret, site_url):
        self.tenant_id = tenant_id
        self.client_id = client_id
        self.site_url = site_url.rstrip("/")
        self._access_token = None
        self._site_id = None
        self._drive_id = None

    def authenticate(self):
        cache = msal.SerializableTokenCache()
        if os.path.exists(self.TOKEN_CACHE_FILE):
            with open(self.TOKEN_CACHE_FILE, "r") as f:
                cache.deserialize(f.read())
        app = msal.PublicClientApplication(self.client_id, authority=f"https://login.microsoftonline.com/{self.tenant_id}", token_cache=cache)
        result = None
        accounts = app.get_accounts()
        if accounts:
            print(f" Reusing cached login ...", end="", flush=True)
            result = app.acquire_token_silent(self.SCOPES, account=accounts[0])
        if not result:
            flow = app.initiate_device_flow(scopes=self.SCOPES)
            if "user_code" not in flow:
                raise RuntimeError(f"Device flow failed: {flow.get('error_description')}")
            print(f"\n\n{'='*60}")
            print("  ACTION REQUIRED - Login to Microsoft")
            print(f"{'='*60}")
            print(f"  1. Open browser: https://microsoft.com/devicelogin")
            print(f"  2. Enter code:   {flow['user_code']}")
            print(f"  3. Login with:   kareemulla@mseducation.academy")
            print(f"{'='*60}\n")
            print("Waiting for login ...", end="", flush=True)
            result = app.acquire_token_by_device_flow(flow)
        if "access_token" not in result:
            raise RuntimeError(f"Auth failed: {result.get('error_description', result.get('error'))}")
        if cache.has_state_changed:
            with open(self.TOKEN_CACHE_FILE, "w") as f:
                f.write(cache.serialize())
        self._access_token = result["access_token"]

    def _headers(self):
        return {"Authorization": f"Bearer {self._access_token}"}

    def _get(self, url, **kwargs):
        resp = requests.get(url, headers=self._headers(), **kwargs)
        resp.raise_for_status()
        return resp

    def connect(self, username=None, password=None):
        self.authenticate()
        parsed = urlparse(self.site_url)
        hostname = parsed.netloc
        site_path = parsed.path.rstrip("/")
        url = f"{self.GRAPH_BASE}/sites/{hostname}:{site_path}" if site_path else f"{self.GRAPH_BASE}/sites/{hostname}:/"
        self._site_id = self._get(url).json()["id"]
        self._drive_id = self._get(f"{self.GRAPH_BASE}/sites/{self._site_id}/drive").json()["id"]

    def list_files(self, folder_path=""):
        url = f"{self.GRAPH_BASE}/drives/{self._drive_id}/root:/{folder_path.strip('/')}:/children" if folder_path else f"{self.GRAPH_BASE}/drives/{self._drive_id}/root/children"
        items = []
        while url:
            data = self._get(url).json()
            items.extend(data.get("value", []))
            url = data.get("@odata.nextLink")
        return items

    def get_file_metadata(self, file_path):
        return self._get(f"{self.GRAPH_BASE}/drives/{self._drive_id}/root:/{file_path.strip('/')}").json()

    def download_file(self, item_id):
        return self._get(f"{self.GRAPH_BASE}/drives/{self._drive_id}/items/{item_id}/content", allow_redirects=True).content

    def download_file_by_path(self, file_path):
        return self.download_file(self.get_file_metadata(file_path)["id"])
