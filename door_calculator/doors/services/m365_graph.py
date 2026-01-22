import requests
import msal
from django.conf import settings
from urllib.parse import quote

GRAPH_BASE = "https://graph.microsoft.com/v1.0"

_app = None


def _get_app():
    global _app
    if _app is None:
        _app = msal.ConfidentialClientApplication(
            client_id=settings.M365_CLIENT_ID,
            authority=settings.M365_AUTHORITY,
            client_credential=settings.M365_CLIENT_SECRET,
        )
    return _app


def get_app_token() -> str:
    app = _get_app()

    # cache -> client credentials
    result = app.acquire_token_silent(settings.M365_SCOPE, account=None)
    if not result:
        result = app.acquire_token_for_client(scopes=settings.M365_SCOPE)

    if "access_token" not in result:
        raise RuntimeError(result)

    return result["access_token"]


def graph(method: str, url: str, **kwargs):
    token = get_app_token()
    headers = kwargs.pop("headers", {})
    headers["Authorization"] = f"Bearer {token}"

    r = requests.request(method, url, headers=headers, timeout=60, **kwargs)
    if not r.ok:
        raise RuntimeError(f"HTTP {r.status_code}: {r.text}")

    return r.json() if r.content else None


def graph_get(path: str, **kwargs):
    # path типу: "/sites?search=*"
    return graph("GET", f"{GRAPH_BASE}{path}", **kwargs)


def graph_get_all_pages(path: str) -> list:
    items = []
    url = f"{GRAPH_BASE}{path}"
    while url:
        data = graph("GET", url)
        items.extend(data.get("value", []))
        url = data.get("@odata.nextLink")
    return items


def find_site_by_display_name(display_name: str) -> dict:
    data = graph_get("/sites?search=*")
    for s in data.get("value", []):
        if s.get("displayName") == display_name:
            return s
    raise RuntimeError(f"Site '{display_name}' not found")


def pick_drive(site_id: str, drive_name: str) -> dict:
    drives = graph_get(f"/sites/{site_id}/drives").get("value", [])
    if not drives:
        raise RuntimeError("No drives found for this site")
    for d in drives:
        if d.get("name") == drive_name:
            return d
    return drives[0]


def list_children(drive_id: str, item_id: str) -> list:
    return graph_get_all_pages(f"/drives/{drive_id}/items/{item_id}/children")


def list_root_children(drive_id: str) -> list:
    return graph_get_all_pages(f"/drives/{drive_id}/root/children")


def search_in_folder(drive_id: str, folder_id: str, q: str) -> list:
    q_enc = quote(q)
    return graph_get_all_pages(f"/drives/{drive_id}/items/{folder_id}/search(q='{q_enc}')")


def graph_put(path: str, **kwargs):
    # path типу: "/drives/{drive_id}/items/{folder_id}:/{filename}:/content"
    return graph("PUT", f"{GRAPH_BASE}{path}", **kwargs)


def upload_bytes_to_folder(drive_id: str, folder_id: str, filename: str, content: bytes,
                           content_type="application/pdf"):
    # overwrite/upload
    safe_name = filename.replace("\\", "_").replace("/", "_")
    path = f"/drives/{drive_id}/items/{folder_id}:/{safe_name}:/content"
    headers = {"Content-Type": content_type}
    return graph_put(path, data=content, headers=headers)
