"""Migrate OneNote SOP pages into a SharePoint list via Microsoft Graph.

Usage:
  python src/onenote_to_sharepoint.py --section-id <section_id>

Authentication uses client credentials. Configure via env vars:
  TENANT_ID, CLIENT_ID, CLIENT_SECRET
  SITE_ID, LIST_ID

Optional env vars:
  PAGE_LIMIT (default: 50)
  FIELD_MAPPING_JSON (JSON mapping of SharePoint fields)
  GRAPH_ROOT (override Graph endpoint, e.g. https://microsoftgraph.chinacloudapi.cn/v1.0)
  AUTHORITY_HOST (override AAD authority host, e.g. https://login.chinacloudapi.cn)
"""
from __future__ import annotations

import argparse
import json
import os
from dataclasses import dataclass
from typing import Dict, List

import msal
import requests

GRAPH_ROOT_DEFAULT = "https://graph.microsoft.com/v1.0"
GRAPH_SCOPE_DEFAULT = "https://graph.microsoft.com/.default"
AUTHORITY_HOST_DEFAULT = "https://login.microsoftonline.com"


@dataclass
class Config:
    tenant_id: str
    client_id: str
    client_secret: str
    site_id: str
    list_id: str
    page_limit: int
    field_mapping: Dict[str, str]
    graph_root: str
    graph_scope: str
    authority_host: str


class GraphClient:
    def __init__(self, config: Config) -> None:
        self.config = config
        self._token = None

    def _acquire_token(self) -> str:
        if self._token:
            return self._token
        authority = f"{self.config.authority_host}/{self.config.tenant_id}"
        app = msal.ConfidentialClientApplication(
            self.config.client_id,
            authority=authority,
            client_credential=self.config.client_secret,
        )
        result = app.acquire_token_for_client(scopes=[self.config.graph_scope])
        if "access_token" not in result:
            raise RuntimeError(f"Failed to acquire token: {result.get('error_description')}")
        self._token = result["access_token"]
        return self._token

    def _headers(self) -> Dict[str, str]:
        return {"Authorization": f"Bearer {self._acquire_token()}"}

    def list_section_pages(self, section_id: str, limit: int) -> List[Dict[str, str]]:
        url = f"{self.config.graph_root}/me/onenote/sections/{section_id}/pages?$top={limit}"
        response = requests.get(url, headers=self._headers(), timeout=30)
        response.raise_for_status()
        return response.json().get("value", [])

    def get_page_content(self, page_id: str) -> str:
        url = f"{self.config.graph_root}/me/onenote/pages/{page_id}/content"
        response = requests.get(url, headers=self._headers(), timeout=30)
        response.raise_for_status()
        return response.text

    def create_list_item(self, fields: Dict[str, str]) -> Dict[str, str]:
        url = f"{self.config.graph_root}/sites/{self.config.site_id}/lists/{self.config.list_id}/items"
        payload = {"fields": fields}
        response = requests.post(url, headers={**self._headers(), "Content-Type": "application/json"}, json=payload, timeout=30)
        response.raise_for_status()
        return response.json()


def build_fields(page: Dict[str, str], html_content: str, mapping: Dict[str, str]) -> Dict[str, str]:
    source = {
        "title": page.get("title", "Untitled SOP"),
        "createdDateTime": page.get("createdDateTime"),
        "lastModifiedDateTime": page.get("lastModifiedDateTime"),
        "pageId": page.get("id"),
        "webUrl": page.get("links", {}).get("oneNoteWebUrl", {}).get("href"),
        "content": html_content,
    }
    fields = {}
    for sp_field, source_key in mapping.items():
        fields[sp_field] = source.get(source_key)
    return fields


def parse_args() -> argparse.Namespace:
    parser = argparse.ArgumentParser(description="Import OneNote pages into a SharePoint list.")
    parser.add_argument("--section-id", required=True, help="OneNote section ID containing SOP pages.")
    parser.add_argument(
        "--list-pages",
        action="store_true",
        help="List OneNote pages (id, title, created, lastModified, webUrl) and exit.",
    )
    parser.add_argument(
        "--list-output",
        help="Optional JSON output path for --list-pages. Defaults to stdout.",
    )
    return parser.parse_args()


def format_page_listing(pages: List[Dict[str, str]]) -> List[Dict[str, str]]:
    return [
        {
            "id": page.get("id"),
            "title": page.get("title"),
            "createdDateTime": page.get("createdDateTime"),
            "lastModifiedDateTime": page.get("lastModifiedDateTime"),
            "webUrl": page.get("links", {}).get("oneNoteWebUrl", {}).get("href"),
        }
        for page in pages
    ]


def load_config() -> Config:
    tenant_id = os.getenv("TENANT_ID")
    client_id = os.getenv("CLIENT_ID")
    client_secret = os.getenv("CLIENT_SECRET")
    site_id = os.getenv("SITE_ID")
    list_id = os.getenv("LIST_ID")

    if not all([tenant_id, client_id, client_secret, site_id, list_id]):
        missing = [
            name
            for name, value in [
                ("TENANT_ID", tenant_id),
                ("CLIENT_ID", client_id),
                ("CLIENT_SECRET", client_secret),
                ("SITE_ID", site_id),
                ("LIST_ID", list_id),
            ]
            if not value
        ]
        raise RuntimeError(f"Missing required environment variables: {', '.join(missing)}")

    page_limit = int(os.getenv("PAGE_LIMIT", "50"))
    graph_root = os.getenv("GRAPH_ROOT", GRAPH_ROOT_DEFAULT)
    graph_scope = os.getenv("GRAPH_SCOPE", GRAPH_SCOPE_DEFAULT)
    authority_host = os.getenv("AUTHORITY_HOST", AUTHORITY_HOST_DEFAULT)
    mapping_json = os.getenv(
        "FIELD_MAPPING_JSON",
        json.dumps(
            {
                "Title": "title",
                "SOPId": "pageId",
                "SOPUrl": "webUrl",
                "CreatedDate": "createdDateTime",
                "LastUpdated": "lastModifiedDateTime",
                "ContentHtml": "content",
            }
        ),
    )
    field_mapping = json.loads(mapping_json)
    return Config(
        tenant_id=tenant_id,
        client_id=client_id,
        client_secret=client_secret,
        site_id=site_id,
        list_id=list_id,
        page_limit=page_limit,
        field_mapping=field_mapping,
        graph_root=graph_root,
        graph_scope=graph_scope,
        authority_host=authority_host,
    )


def main() -> None:
    args = parse_args()
    config = load_config()
    client = GraphClient(config)

    pages = client.list_section_pages(args.section_id, config.page_limit)
    if not pages:
        print("No pages found in the specified section.")
        return

    if args.list_pages:
        listing = format_page_listing(pages)
        output = json.dumps(listing, indent=2, ensure_ascii=False)
        if args.list_output:
            with open(args.list_output, "w", encoding="utf-8") as handle:
                handle.write(output)
            print(f"Wrote page listing to {args.list_output}")
        else:
            print(output)
        return

    for page in pages:
        page_id = page.get("id")
        if not page_id:
            print("Skipping page with missing id.")
            continue
        content = client.get_page_content(page_id)
        fields = build_fields(page, content, config.field_mapping)
        created = client.create_list_item(fields)
        print(f"Created list item {created.get('id')} for page {page.get('title')}")


if __name__ == "__main__":
    main()
