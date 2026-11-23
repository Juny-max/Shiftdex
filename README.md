# ShiftDex

ShiftDex is an Electron desktop roster editor, a mobile worker portal, and a Cloudflare Worker API backed by D1. It models the 24/7 roster as 12-hour shifts on a continuous 2M → 2N → 4 OFF cycle.

## Run locally

Requirements: Node.js 20 or newer and pnpm 12.15.0.

```powershell
pnpm install
pnpm dev
```

`pnpm dev` opens the Electron administrator app. `pnpm web` starts the worker portal, and `pnpm worker:dev` starts the local API.

Use the roster’s **Import** control to load the workbook. The source workbook contains staff details; keep it in a private location and do not commit it. The checked-in initial roster is intentionally empty of staff and assignments.

## Cloudflare setup

1. Sign in to the Cloudflare account that owns the D1 database:

   ```powershell
   pnpm exec wrangler login
   ```

2. Copy `worker/wrangler.example.toml` to `worker/wrangler.toml`. Put the D1 database ID and public configuration in the local file. Git ignores `worker/wrangler.toml`.
3. Create the tables and generic pod definitions:

   ```powershell
   pnpm exec wrangler d1 execute shift-roster --remote --file=worker/schema.sql --config worker/wrangler.toml
   ```

   The schema does not insert staff names or shift assignments. Import the private workbook in ShiftDex, then publish the roster.

4. Add the Brevo API key as a Cloudflare secret. Wrangler prompts for the value; do not put the key in a file or command argument:

   ```powershell
   pnpm set-brevo-key
   ```

   Configure the verified sender address in the local Wrangler file. Sender settings are public configuration; the API key is a secret.

5. Deploy the API:

   ```powershell
   pnpm worker:deploy
   ```

6. Deploy the worker portal to the configured Cloudflare Pages project. Build it, then upload `dist/worker` with Wrangler Pages.

## Administrator account commands

Both commands require an interactive terminal, a local `worker/wrangler.toml`, an authenticated Wrangler session, and Cloudflare D1 write permission for the configured database. No Cloudflare token or reset password is built into the public repository.

### Create the first administrator

Run this once after applying the schema:

```powershell
pnpm admin:create
```

The command only creates an account when `admin_accounts` is empty. It asks for the email, display name, and a password without echoing the password. It refuses to add or alter accounts when an administrator already exists.

### Reset an existing administrator password

```powershell
pnpm admin:reset-password
```

This command only resets the active administrator matching the email entered at the prompt. It cannot create an account, change an email, or reset a worker password. It confirms the target database and email, requires a 16-character minimum password entered without echo, and revokes that administrator’s existing sessions. The reset is recorded in the audit log. Its temporary SQL file is removed when the command finishes.

The script is not itself a secret or an access-control mechanism. Anyone can read or run the public script against a database they control. It cannot change your Cloudflare database unless that person is signed in to an account with D1 write access to your database and has the local database configuration. Do not share Cloudflare credentials or grant D1 write permission to untrusted people.

Cloudflare Workers currently rejects Web Crypto PBKDF2 counts above 100,000. ShiftDex uses that supported ceiling, salts each password, rate-limits sign-in attempts, requires long passwords for reset, and revokes sessions on reset. See [Cloudflare workerd issue 1346](https://github.com/cloudflare/workerd/issues/1346) for the platform limit.

## Publishing and updates

- In the desktop app, sign in with the administrator email and password, then use **Publish schedule** to sync roster changes to D1.
- Pushes to `main` run the Windows packaging workflow. It publishes the installer and `latest.yml` to a GitHub release.
- Installed apps check for updates at startup and about every six hours. **Settings → Check for updates** also checks manually.
- Release downloads use the public GitHub repository. Never add a GitHub token to the app or repository to enable updates.

## Data and security notes

- The live D1 database is separate from the public source repository. Workers must authenticate to read the roster. Worker accounts can submit their own cover and unavailability requests; administrator actions require an admin session.
- Passwords are stored as salted PBKDF2 hashes, never plaintext. D1 stores session-token hashes. Admin sessions expire after 12 hours; worker sessions expire after 30 days.
- `worker/wrangler.toml`, `.env` files, local build output, and release artifacts are ignored by Git. Use `worker/wrangler.example.toml` as the public template.
- A public repository makes its source, commit history, workflow, and releases visible to everyone. Do not commit real names, roster exports, credentials, private configuration, or production data. The app cannot hide information intentionally sent to a client from a person who controls that client device.
- Keep the workbook, local roster cache, Cloudflare account, D1 permissions, and Brevo account access limited to people who need them. Back up D1 before destructive database changes.

## Windows installer

Build a per-user installer with:

```powershell
pnpm dist
```

The installer is named `ShiftDex-Setup-<version>.exe`. Windows installs `ShiftDex.exe` and creates Start menu and desktop shortcuts. Uninstall from Windows Settings or the Start menu; the NSIS configuration removes ShiftDex app data during uninstall.

## Operations and telecommunications model

ShiftDex is purpose-built for mission-critical 24/7/365 telecommunications and cell tower infrastructure management. Shared cell sites hosting multiple carriers (e.g. MTN, Telecel, AT/AirtelTigo) require continuous, real-time power and facility oversight:

- **Grid and generator monitoring**: Monitoring ECG grid power continuity, automated and manual generator start triggers, battery bank backups, and rectifier stability. If a plant generator fails to start during a power outage, on-duty field teams must be alerted immediately to prevent site drops and subscriber outage.
- **Continuous 24/7/365 mission**: Because cell tower communications must remain operational without interruption, the operations hub operates continuous 12-hour shifts every day of the year without company-wide holidays.
- **Continuous rotation model**: Core shift teams operate on an 8-day rotation: 4 days on (2 Day shifts 06:00–18:00, 2 Night shifts 18:00–06:00), followed by 4 days off (2M → 2N → 4 OFF), providing 100% continuous desk coverage.

## Collective shift rules

Administrators can apply scheduling policies across multiple staff simultaneously via **Settings & Shift Rules**:

- **Batch Shift**: Assigns a specific shift code (`M`, `N`, `OFF`, `LEAVE`) across an entire target group (All Personnel, Team Leads, NSP Personnel, a specific Team, or Day Support) for any date range in one action. This eliminates repetitive cell-by-cell editing during group training, maintenance drills, or temporary reassignment.
- **Swap Teams**: Exchanges the 8-day rotation cycles between two teams across a chosen date range. Useful when two teams trade schedules due to field logistics, emergency maintenance, or operational swaps.
- **Adjust Rotation**: Advances or delays an entire team's 8-day rotation anchor by +1, +2, +4, -1, or -2 days, realigning their rotational phase without disrupting individual shift history.
- **Instant reversibility**: All collective shift rules push state to the undo stack, allowing full rollback via the **Undo** button.

## Shift teams and anchors customization

Administrators can customize teams directly in **Settings & Shift Rules**:

- **Team names**: Customize team labels (e.g., *Team 1*, *Alpha Pod*, *Power & Tower Ops*).
- **Dynamic color palette**: Select from curated telecom presets or choose any custom hex color. Custom team colors automatically apply soft light-wash background gradients to calendar headers (`color-mix(in srgb, var(--pod-color) 12%, #f8fafc)`) and avatar chips while preserving clean UI contrast.
- **Rotation anchors**: Adjust the baseline anchor date to control which phase of the 8-day rotation each team begins on.
- **Dynamic team expansion**: Add new shift teams or remove unassigned teams as operations scale.

## Workbook and roster model

The supplied workbook uses orange `#FFC000`, green `#92D050`, blue `#00B0F0`, and red `#FF0000` pod colors. The continuous pod anchors are Orange `2026-11-03`, Green `2026-11-01`, Blue `2026-11-05`, and Red `2026-10-30`. Import reads the sheet’s staff names, role tags, pod colors, and shift cells; blank shift cells become rest days. These values are not included as staff seed data in the public source.

## Contributors

- **Shift Desk** (Lead & Product Operations)
- **Antigravity** (Google DeepMind) - System architecture, Cloudflare Workers & D1 API, AI scheduling copilot, security hardening, and Electron desktop development.
