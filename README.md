# OneNote SOPs to SharePoint List

This repository contains a small migration utility that reads SOP pages from a OneNote section and creates matching items in a SharePoint list using Microsoft Graph.

## What it does
- Reads OneNote pages from a specific section.
- Pulls the HTML content of each page.
- Maps OneNote metadata + content into your SharePoint list fields.

## Prerequisites
- Azure AD app registration with Microsoft Graph permissions:
  - `Notes.Read.All`
  - `Sites.ReadWrite.All`
- A SharePoint list with columns that match your mapping.

## Setup
1. Install dependencies:
   ```bash
   python -m venv .venv
   source .venv/bin/activate
   pip install -r requirements.txt
   ```

2. Set required environment variables:
   ```bash
   export TENANT_ID="your-tenant-id"
   export CLIENT_ID="your-client-id"
   export CLIENT_SECRET="your-client-secret"
   export SITE_ID="your-site-id"
   export LIST_ID="your-list-id"
   ```

3. Optional configuration:
   ```bash
   export PAGE_LIMIT=50
   export FIELD_MAPPING_JSON='{"Title":"title","SOPId":"pageId","SOPUrl":"webUrl","CreatedDate":"createdDateTime","LastUpdated":"lastModifiedDateTime","ContentHtml":"content"}'
   ```

4. China (21Vianet) cloud configuration:
   ```bash
   export GRAPH_ROOT="https://microsoftgraph.chinacloudapi.cn/v1.0"
   export GRAPH_SCOPE="https://microsoftgraph.chinacloudapi.cn/.default"
   export AUTHORITY_HOST="https://login.chinacloudapi.cn"
   ```

## Run
```bash
python src/onenote_to_sharepoint.py --section-id <onenote-section-id>
```

## List OneNote pages
```bash
python src/onenote_to_sharepoint.py --section-id <onenote-section-id> --list-pages
```

To save the listing as JSON:
```bash
python src/onenote_to_sharepoint.py --section-id <onenote-section-id> --list-pages --list-output pages.json
```

## Notes
- The mapping is a JSON object where keys are SharePoint column internal names and values are OneNote-derived keys.
- The script uses the `/me/onenote` endpoints. If you need tenant-wide access, switch to `/users/{user-id}/onenote` or `/groups/{group-id}/onenote`.
