# PDM & PM tracker → live dashboard

How the pieces fit together:

```
Google Sheet ("PDM & PM")
   │  edit a row, click Dashboard ▸ Publish now
   ▼
Apps Script (apps-script/Code.gs)
   │  reads Products + Engineering tabs, builds data.json,
   │  commits it via the GitHub Contents API
   ▼
GitHub repo (radhikap-hash/PDM-and-PM-tracker-)
   │  dashboard.html + data.json, served by GitHub Pages
   ▼
Browser: dashboard.html fetches ./data.json on load
```

Nothing lives outside those three places — no server to run, no separate
backend. The Sheet is the input, the repo is the store, GitHub Pages is the
host.

## Files in this folder

- `dashboard.html` — your dashboard, modified to fetch `data.json` at load
  instead of using a hardcoded array. If `data.json` can't be reached (e.g.
  opened as a local file, or before the pipeline is wired up) it falls back
  to a baked-in snapshot and shows an amber "showing last saved snapshot"
  note, so it never breaks.
- `data.json` — the current snapshot (same numbers as your original file).
  This is what gets overwritten each time the sheet is published.
- `apps-script/Code.gs` — bound to the Google Sheet, does the publishing.

## One-time setup

**1. Push these files to the repo** (I can't write to your GitHub directly —
I don't have credentials for it. Either drag `dashboard.html` and `data.json`
into the repo on github.com, or from a terminal you're logged into:
```
git clone https://github.com/radhikap-hash/PDM-and-PM-tracker-.git
cp dashboard.html data.json <repo>/
cd <repo> && git add . && git commit -m "Add live dashboard" && git push
```

**2. Enable GitHub Pages** — repo Settings ▸ Pages ▸ Source: `main` branch,
root folder. You'll get a URL like
`https://radhikap-hash.github.io/PDM-and-PM-tracker-/dashboard.html`.
That's the link to bookmark/share — opening `dashboard.html` as a local file
won't be able to fetch `data.json` due to browser restrictions.

**3. Set up the Sheet.** Open the published sheet, Extensions ▸ Apps Script,
paste in `apps-script/Code.gs`, save. Run `setup()` once from the editor
toolbar (approve the permissions prompt) — this creates two tabs if they
don't already exist:

- **Products** — one row per product variant:
  `Category | Product Name | ID | PDM | PM/TL | Stage (1-11) | Days In Stage | Days Total | Current Doc | Blocker | Tentative Date`
- **Engineering** — one row per weekly update (latest row per Product ID
  is what gets published):
  `Product ID | Period | Owners | Platform | Approach | Completed | Next Steps | Risk | Status | ETA`
  — for Completed/Next Steps, put one bullet per line inside the cell
  (Alt+Enter in Sheets), or separate them with ` | `.

Reload the sheet — a **Dashboard** menu appears.

**4. Create a GitHub token.** github.com ▸ Settings ▸ Developer settings ▸
Fine-grained tokens ▸ generate one scoped to just this repo, with
**Contents: Read and write** permission. Copy it.

**5. Wire the token into the script.** In the Apps Script editor: Project
Settings ▸ Script Properties ▸ add:
| Key | Value |
|---|---|
| `GITHUB_TOKEN` | the token from step 4 |
| `GITHUB_OWNER` | `radhikap-hash` |
| `GITHUB_REPO` | `PDM-and-PM-tracker-` |
| `GITHUB_BRANCH` | `main` |
| `GITHUB_PATH` | `data.json` |

**6. Publish.** Back in the Sheet: Dashboard ▸ Publish now. Check the repo —
`data.json` should update, and the live dashboard URL should reflect it
within a few seconds (GitHub Pages caches briefly; hit the "↻ Refresh data"
button on the page if it looks stale).

Optional: Dashboard ▸ "Enable auto-publish" makes every edit publish itself
~2 minutes later (debounced, so a flurry of edits = one commit, not one per
keystroke), so you never have to remember to click Publish.

## Notes

- `STAGE_DOCS` (the per-stage document checklist) and `STAGES` stay
  hardcoded in `dashboard.html` — they're structural and don't change
  week to week. Only the two data tabs are wired up to the pipeline.
- Several `CAM_ENGINEERING` rows in the original file were noted as
  "folded from" a differently-named source note (Renesas Platform,
  Bellycam, VCT SoM) — worth confirming those mappings are correct before
  they become the seed data other people edit from.
