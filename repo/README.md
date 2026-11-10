# Monster Brewing Field Survey Dashboard — automated build

Rebuilds https://mbrewfielddashboard.netlify.app/ every morning straight from Power BI
and deploys it to Netlify. No exports, no uploads. The site URL never changes.

```
Power BI (KARMA model) ──┐
                         ├─► build_from_powerbi.py ─► dist/index.html ─► Netlify CLI ─► live site
Power BI (Depletions) ───┘
```

## What's in this repo

| File | Purpose |
|---|---|
| `build_from_powerbi.py` | Pulls surveys + buyer commitments (KARMA model) and distributors, chain flags, case equivalents (Depletions model); renders and encrypts the page |
| `template/dashboard.html` | The dashboard itself (HTML/CSS/JS, logos, map) with `__ROWS__`, `__BUY__`, `__EMPMAP__` placeholders |
| `template/wrapper.html` | Login page + decryption wrapper; `__PIV__` / `__PCT__` are filled at build time |
| `empmap.json` | Employee → `[title, manager]`. **Edit this when reps join or move.** Manager drives the Manager filter |
| `dist_codes.json` | Distributor code → name fallback for codes not in the Depletions Bottler table |
| `.github/workflows/build.yml` | Daily 12:00 PM EST schedule + manual "Run workflow" button |
| `tests/fixture.json` | Small sample for a local dry run: `python3 build_from_powerbi.py --fixture tests/fixture.json` |

## One-time setup

### 1. IT: Power BI service principal (the only blocker)
Ask IT for a Microsoft Entra ID **app registration** and give you:
- Tenant ID (`f01d65ba-e57e-4f8f-8143-a43ed634f3c2`), Application (client) ID, and a client secret
- In the Power BI admin portal → Tenant settings → **Developer settings** → *Allow service principals to use Power BI APIs* enabled for a security group containing the app
- The app added as **Member** (or Contributor) of the **Monster Brewing** workspace, so it can run queries against
  `MBrew KARMA Preview` (07f05c80-…) and `Monster Brewing Depletions Dashboard` (53e6e702-…)

No API permissions need to be granted on the app registration itself — workspace membership is what authorizes `executeQueries`.

### 2. Netlify
- Site settings → General → **Site ID** (copy it)
- User settings → Applications → **Personal access tokens** → New access token (copy it)

### 3. GitHub
Create a private repo, push these files, then **Settings → Secrets and variables → Actions → New repository secret** for each:

| Secret | Value |
|---|---|
| `PBI_TENANT_ID` | from IT |
| `PBI_CLIENT_ID` | from IT |
| `PBI_CLIENT_SECRET` | from IT |
| `NETLIFY_AUTH_TOKEN` | Netlify personal access token |
| `NETLIFY_SITE_ID` | Netlify site ID |

Then **Actions → Build dashboard from Power BI → Run workflow** to test. The run log shows survey and
commitment counts, and the finished `index.html` is attached to the run as an artifact if you want to inspect it.

## Weekly operation
Nothing. The workflow runs every day at 12:00 PM EST (1 PM during daylight time). To refresh sooner, press *Run workflow*.
If the log shows a `NEW EMPLOYEES` warning, add them to `empmap.json` with their title and manager and commit —
the next run picks it up.

## How the data is shaped
- **Surveys**: forms `MBREW B.T.B. 2026` and `MBREW B.T.B. 2026 - NY`, from `SURVEY_START` (2026-07-01).
  One dashboard row per (outlet, rep, timestamp). Account = store name + store number; chain = store name when
  VIP flags the outlet as *Chain*, otherwise *Independents*; distributor = the outlet's VIP distributor, named via the
  Depletions Bottler table; soft-drink bottlers are dropped.
- **Commitments**: forms `Buyer Commitments` and `Buyer Commitments - UT`. One block per (outlet, distributor, rep);
  commitments accumulate across visits; survey dates are the real KARMA submission dates.
- **Status**: `FULFILLED` when the outlet has case-equivalent depletions from that distributor in the trailing
  3-month window (first of the month two months before the latest invoice, through the latest invoice); otherwise `OPEN`.
- **Data lag**: the KARMA model usually trails KARMA by a few days and Depletions by about a week; the run log prints both dates.

## Changing the dashboard
Edit `template/dashboard.html` (the UI) — never the generated `dist/index.html`. The three placeholders must stay.
Run the fixture build locally to check it still renders, then commit.
