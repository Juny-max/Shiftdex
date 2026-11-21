# Shift Desk

Electron desktop roster editor, mobile worker portal, and Cloudflare D1 API for the 24/7 12-hour roster.

## Workbook analysis and domain model

The supplied `NOV-DEC 2026 ROOSTER.xlsx` has two sheets, `NOV 2026 ROOSTER` and `DEC 2026 ROOSTER`. Both use the same spreadsheet layout: weekday/date headers across columns E onward, staff names in column A, four pod blocks at rows 4–15, then Fueling, Power NOC, Manager, and leave/support rows at 19–23. The four source pod fills are Orange `#FFC000`, Green `#92D050`, Blue `#00B0F0`, and Red `#FF0000`; morning cells are yellow and night cells are pale blue.

The anchors below are inferred from the workbook's date cells and the uninterrupted 2M/2N/4OFF sequence in both months:

| Pod | First M date in this continuous sequence | Color |
|---|---|---|
| Green | 2026-11-01 | `#92D050` |
| Orange | 2026-11-03 | `#FFC000` |
| Blue | 2026-11-05 | `#00B0F0` |
| Red | 2026-10-30 | `#FF0000` |

This matches the workbook across the November/December boundary. Pod membership and role assignments can vary between months, so import reads pod color, staff identity, and role tag from each sheet row. Blank roster cells import as OFF and unmatched staff rows are reported in the import summary. The workbook contains staff information and should remain private; it is not included in this repository.

## Run the desktop app

1. Install Node.js 20 or newer and pnpm 12.15.0.
2. Run `pnpm install` in this folder.
3. Run `pnpm dev` to open the Electron admin app. `pnpm web` serves the worker portal locally; `pnpm worker:dev` starts the Cloudflare Worker API.

## Build and release the Windows app

- Run `pnpm dist` to build a single per-user installer at `release/ShiftDex-Setup-1.0.1.exe` (the version follows `package.json`). The installed program and shortcuts are named **ShiftDex**; the installed executable is `ShiftDex.exe`.
- Uninstall it from Windows Settings or the Start menu. The installer is configured to remove ShiftDex's local app data when uninstalling.
- Every push to the `main` branch runs the Windows release workflow. It builds the installer, generates `latest.yml`, and publishes both files as a GitHub release. The release patch version increments automatically from the package version, so installed copies can discover and install newer releases from Settings. Keep the major and minor package version current when starting a new release line.

This project uses pnpm and its `pnpm-lock.yaml`. Avoid running `npm install` over pnpm's `node_modules`: npm can encounter pnpm-specific `link:` entries in already installed package metadata. The `allowBuilds` entries in `pnpm-workspace.yaml` permit the Electron binary download, esbuild setup, and Cloudflare workerd runtime setup required by these scripts.

The roster works locally in the desktop app and saves in browser local storage. Use the workbook buttons to import or export `.xlsx`; export creates a month sheet in the roster's tabular layout. The current export covers the pod matrix and support names. PNG export uses the system print dialog; choose Save as PDF or a PNG printer/extension if installed.

## Cloudflare D1 API

1. Copy `worker/wrangler.example.toml` to `worker/wrangler.toml` and put your D1 ID and public configuration in the local file. The local config is ignored by Git. Store API keys with `wrangler secret put`, never in TOML, source code, or the desktop app.
2. From the project root, run `pnpm exec wrangler d1 execute shift-roster --remote --file=worker/schema.sql --config worker/wrangler.toml` to create tables and seed the generic pods. Add staff by importing the workbook in the admin app, then publish the roster. Rerun this command after updating an existing database so it creates the monthly pod assignment table used by drag-and-drop team changes.
3. Deploy the API with `pnpm exec wrangler deploy --config worker/wrangler.toml` from the project root.
4. Create the first administrator with `pnpm admin:create`. The script requires an authenticated Wrangler session, prompts for the password without echoing it, and only creates an account while `admin_accounts` is empty. Admin passwords are salted and hashed before they reach D1. The script contains no Cloudflare credentials and does not deploy or modify Worker configuration.
5. Workers create accounts with email and password; their registrations wait for an administrator to link each account to the correct staff record in the Requests page. Worker passwords are salted and hashed in D1.
6. The API URL is configured in the app build. In ShiftDex, use **Publish schedule**, then sign in with the administrator email and password. Admin sessions expire after 12 hours and are stored hashed in D1; workers never enter API URLs or admin credentials.

## Worker portal

The static worker portal lives in `worker/`. Deploy its build output to Cloudflare Pages. Staff register with their name, email, and password; an administrator links the account to its roster record before sign-in is enabled. The portal includes a schedule view, a cover request form, unavailability submission, and downloadable `.ics` shifts.

## API routes

- `GET /api/health`
- `GET /api/roster?month=YYYY-MM` (authenticated worker/admin read)
- `POST /api/admin/auth/login` and `POST /api/admin/auth/logout` (admin email/password)
- `POST /api/admin/publish` (admin session)
- `GET /api/admin/requests` and `POST /api/admin/requests/:id` (admin session)
- `POST /api/covers` and `POST /api/unavailability` (worker token; identity is checked against payload)

The D1 schema includes pods, staff, shift assignments, worker and administrator accounts/sessions, shift covers/payback links, pre-roster unavailability flags, and audit logs.

## Current integration boundary

The Electron app and worker portal are scaffolded as independent clients. Worker requests are stored in D1 and the admin API can list and approve/reject them. The admin drawer currently displays a local placeholder until its request sync UI is connected. Cover approval records request status; automatic cover overrides, payback settlement, calendar subscription feeds, and natural-language AI command execution still need to be wired before operational rollout. The provided workbook's staff-name changes and support/leave entries should be reviewed during import.
