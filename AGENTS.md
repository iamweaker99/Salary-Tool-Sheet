# Google Sheet Project

## Standing update rules

1. **Phase gates:** Work phase by phase; stop, report completion, and wait for approval before the next phase.
2. **Very wrong scenarios:** Stop and discuss if no solution is available or a solution would materially broaden the requested change (about 20%+).
3. **Tool split:** Discover and use installed tools first: `gws` for live Sheet data, `clasp` for Apps Script deployment, and signed-in Edge computer use for UI buttons/verification when service-account access is insufficient.
4. **Access reality:** The service account has limited live-Sheet access; do not assume it can perform every action. Use the signed-in browser as the fallback.
5. **Live tests:** Record original values, use small controlled test rows, restore every temporary change, and ask the user only when a safe row or manual result genuinely requires input. 
6. **Handoff:** Use `/handoff` or `$handoff` only when explicitly requested, and save the handoff outside the workspace.

## The Two CLI Tools

This project uses **two complementary CLI tools** for Google Sheets automation:

### 1. `clasp` — Push/pull Apps Script code
- **What it does**: Uploads/downloads the Apps Script project code bound to the Sheet (functions, triggers, custom menus)
- **Key commands**:
  - `clasp push` — push local code to the remote script project
  - `clasp pull` — pull the latest remote code to local
  - `clasp status` — see tracked/untracked files
  - `clasp open` — open the script in browser
  - `clasp run <function>` — execute a function remotely
- **Config**: `.clasp.json` links to Script ID `1MZY04eJUPUY8B2kOzIA0BXgt2vn_lw08Y_pRhDQqlg-VEhGc64G-_bl1`
- **Script ID** (from URL): `1MZY04eJUPUY8B2kOzIA0BXgt2vn_lw08Y_pRhDQqlg-VEhGc64G-_bl1`

### 2. `gws` — Live Sheet data read/write
- **What it does**: Reads/writes cell values, ranges, and appends rows directly to the Sheet from the CLI — no browser needed
- **Auth**: Uses a service account via `GOOGLE_WORKSPACE_CLI_CREDENTIALS_FILE`
- **Credentials file**: `C:\Github Repo\Salary-Tool-Sheet\gws_credentials.json`
- **Sheet ID**: `19paXVwMO2vtKxL0LQw64jLRPXPmYAS2_qoJQ5yM6ek8`
- **Key commands**:
  - `gws sheets spreadsheets values get` — read a range
  - `gws sheets spreadsheets values update` — write data
  - `gws sheets +append --spreadsheet <ID> --values "a,b,c"` — append a row
  - `gws drive files list` — list/find files

### Deployment workflow
- **Code changes** → edit local `.js`/`.gs` files → `clasp push`
- **Data operations** → use `gws` commands directly (no deployment needed — it's live)

## Sheet Details
- **URL**: `https://docs.google.com/spreadsheets/d/19paXVwMO2vtKxL0LQw64jLRPXPmYAS2_qoJQ5yM6ek8/edit`
- **Service account email**: `salary@salary-tool-503801.iam.gserviceaccount.com` (shared as Editor)
- **GCP Project**: `salary-tool-503801` (project number: `922207163069`)
- **Timezone**: `Asia/Singapore`
- **Runtime**: V8

## Enabling APIs (if needed)
- Drive API: `https://console.developers.google.com/apis/api/drive.googleapis.com/overview?project=922207163069`
- Sheets API: Already enabled
- App Script API: `AIzaSyAFOppy_cXZ2TtDNWzbn9RFNr2JSb4Q720`