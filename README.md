# BC Data Connector — Excel Add-in for Business Central

A task-pane Excel add-in that:
- Adds a **Business Central** group to the Excel ribbon with **Add Table** and **Refresh Data** buttons.
- **Add Table** opens a 3-step wizard: sign in to Business Central → pick table & fields → pick a placeholder cell → writes the data as a native Excel Table.
- **Refresh Data** lets you re-pull one or more existing tables. If a refresh would grow into another registered table's cells, it stops and asks you to relocate the other table or cancel.
- Supports 8 tables via the Microsoft **Standard API v2.0**: Customer, Vendor, Sales Header (Sales Order), Sales Line (Sales Order Lines), Customer Ledger Entry, Detailed Customer Ledger Entry, Sales Inv. Header, Sales Inv. Line.
- No client secrets anywhere — sign-in uses delegated OAuth2 (MSAL.js, Authorization Code + PKCE), the same safe pattern Microsoft's own add-ins use.
- Can be added or removed from Excel at any time without touching your workbook data.

## Project files
```
bc-excel-addin/
├── manifest.xml              ← add-in manifest (sideload this file)
├── assets/                   ← ribbon icons (generated)
└── src/taskpane/
    ├── taskpane.html          UI shell
    ├── taskpane.css           styling
    ├── taskpane.js            UI wiring (wizard, refresh, manage)
    ├── config.js              the 8-table entity map
    ├── auth.js                MSAL sign-in wrapper
    ├── bcapi.js                Business Central REST calls
    ├── excelops.js            Excel JS API: create/resize/move tables, overlap checks
    └── store.js               persists connections/tables inside the workbook
```

## Important note on table mapping
Business Central's Standard API does **not** have one combined "Sales Header/Sales Line" entity across all document types — each document type (Quote/Order/Invoice/Credit Memo) is its own API entity. I mapped:
- "Sales Header" / "Sales Line" → `salesOrders` / `salesOrderLines` (closest table-level equivalent)
- "Sales Inv. Header" / "Sales Inv. Line" → `salesInvoices` / `salesInvoiceLines` (exact match)

`customerLedgerEntries` and `detailedCustomerLedgerEntries` are listed by Microsoft as **Beta** endpoints under `/api/v2.0` — functional today, but Microsoft could still change them.

---

## Step A — One-time setup: register an Azure AD app (needed for sign-in)

1. Go to [portal.azure.com](https://portal.azure.com) → **Microsoft Entra ID** → **App registrations** → **New registration**.
2. Name it (e.g. `BC Excel Connector`), choose the supported account type for your tenant, leave Redirect URI blank for now → **Register**.
3. Copy the **Application (client) ID** and **Directory (tenant) ID** from the Overview page — you'll type these into the wizard later.
4. Go to **Authentication** → **Add a platform** → **Single-page application** → Redirect URI: the exact URL where `taskpane.html` will be hosted (see Step B), e.g. `https://localhost:3000/src/taskpane/taskpane.html` → **Configure**.
5. Go to **API permissions** → **Add a permission** → **APIs my organization uses** → search **Dynamics 365 Business Central** → **Delegated permissions** → check `user_impersonation` → **Add permissions**. If your tenant requires it, click **Grant admin consent**.

## Step B — Host the files over HTTPS (Excel requires HTTPS)

On your own PC (with Node.js installed):

```bash
cd bc-excel-addin
npx office-addin-dev-certs install
npx http-server -S -C "%USERPROFILE%\.office-addin-dev-certs\localhost.crt" -K "%USERPROFILE%\.office-addin-dev-certs\localhost.key" -p 3000
```
(On Mac/Linux replace `%USERPROFILE%` with `~`.)

This serves the folder at `https://localhost:3000`.

> Alternative: upload the folder to any static HTTPS host you control (Azure Static Web Apps, GitHub Pages, SharePoint, etc.) and use that URL instead of localhost everywhere below.

## Step C — Point the manifest at your host

Open `manifest.xml` and replace every `https://REPLACE_WITH_YOUR_HOST` with your actual host, e.g. `https://localhost:3000`.

## Step D — Sideload the add-in into Excel

**Excel desktop (Windows/Mac):**
Insert tab → **My Add-ins** → **...** (or the dropdown arrow) → **Upload My Add-in** → browse to `manifest.xml` → **Upload**.

**Excel on the web:**
Insert tab → **Add-ins** → **Upload My Add-in** → choose `manifest.xml`.

You'll see a new **Business Central** group on the Home ribbon with **Add Table** and **Refresh Data**.

## Step E — Use it
1. Click **Add Table** → enter Tenant ID, Client ID (from Step A), Environment Type, Environment Name → **Sign in & load companies** → pick the company → **Next**.
2. Pick one of the 8 tables → uncheck any fields you don't want → **Next**.
3. Click a cell on the sheet → **Use active cell** → name the table → **Create table**.
4. Repeat with **Add another table**, or go back to Home.
5. Click **Refresh Data** any time → check the tables to refresh → **Refresh selected**. If a refresh would overlap another placeholder table, a dialog lets you click a free cell to relocate the other table, or cancel that refresh.
6. **Manage connections & tables** (from Home) lets you remove a placeholder table (deletes it from the sheet) or forget a saved connection.

---

## Uninstalling (if anything goes wrong)

**Excel desktop:** Insert tab → **My Add-ins** → right-click (or the **...** menu on) **BC Data Connector** → **Remove**.

**Excel on the web:** Insert tab → **Add-ins** → **My Add-ins** → hover the tile → trash/remove icon.

Removing the add-in only removes the ribbon buttons and task pane — **any Excel Tables already created stay on your sheet as normal data/tables**; nothing is deleted from your workbook.

If Excel still shows stale UI after removing: close Excel completely and delete the Office add-in cache folder (Windows: `%LOCALAPPDATA%\Microsoft\Office\16.0\Wef\`), then reopen Excel.

To stop hosting: just stop the `http-server` process (Ctrl+C) — the add-in will show a connection error next time it's opened, which is expected once you no longer need it.

---

## Known limitations (read before relying on this for production)
- **Field discovery** reads one sample record to infer column names — if a table/company has zero records, no fields will list for it yet.
- **Overlap resolution** handles conflicts one pair at a time per refresh; very densely packed sheets with many placeholders may occasionally need a second refresh pass.
- **Row cap**: fetches stop at 20,000 rows per table per refresh (edit `maxRows` in `bcapi.js` to change).
- MSAL.js is loaded from Microsoft's legacy CDN (`alcdn.msauth.net`, v2.x). Microsoft has deprecated CDN hosting for MSAL v3+; v2.x CDN builds still work today but if Microsoft ever retires them, download `msal-browser.min.js` from npm and host it alongside your own files instead.
- This uses delegated (user) sign-in, not app-only/client-credentials — intentional, since a client secret should never live in browser JavaScript.
