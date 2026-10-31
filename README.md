<img src="assets/icon.svg" width="64" height="64" alt="">

# Assignments+ — Material UI for Teams Assignments

A **student-side, read-only** alternative frontend for Microsoft Teams Assignments, built as a browser extension (Manifest V3) with a React + Vite Material 3 dashboard.

## Research: what's actually going on in Teams Web

The endpoints in the original spec (`edu.microsoft.com`, `ts.accessToken`) don't reflect the current client. Verified against Microsoft's docs:

### Real API surface (Microsoft Graph, delegated student scopes only)

| Need | Endpoint | Notes |
|---|---|---|
| All your assignments, all classes | `GET /v1.0/education/me/assignments` | In **v1.0** since Aug 2024. Least-privilege scope `EduAssignments.ReadBasic`. **Lean payload** — since the 2021 breaking change it omits `instructions`, `resources`, `submissions`. |
| Your classes | `GET /v1.0/education/me/classes` | `EduRoster.ReadBasic`. Fallback: `GET /v1.0/me/joinedTeams` (class team id == education class id). |
| Your submission + grade | `GET /v1.0/education/classes/{c}/assignments/{a}/submissions?$expand=outcomes` | Students only ever see their own submission; `outcomes` carries `educationPointsOutcome` (score) and `educationFeedbackOutcome` (teacher feedback). |
| Full detail (instructions + attachments) | `GET /v1.0/education/classes/{c}/assignments/{a}?$expand=resources` | `educationAssignment.instructions` (itemBody HTML) + `resources`. |
| Deep link back into Teams | `educationAssignment.webUrl` | Official property — no deep-link reverse engineering needed. |

Relevant enums — assignment `status`: `draft, scheduled, published, assigned, inactive`; submission `status`: `excused, reassigned, returned, submitted, working`.

### Auth reality check

- Teams Web uses **MSAL Browser**; tokens live in `sessionStorage`/`localStorage` under client-id-prefixed keys — and since recent MSAL builds those cache entries are **encrypted** (HKDF key in an `msal.cache.encryption` cookie). Scraping storage is a dead end.
- What still works: the **`Authorization: Bearer` header** the Teams tab itself sends. This extension hooks `fetch`/`XMLHttpRequest` in the page's MAIN world, decodes each JWT, and keeps only tokens whose audience is Microsoft Graph (`aud = 00000003-0000-0000-c000-000000000000`). The token is used exactly as the Teams tab uses it — read-only, on-device, never uploaded anywhere.
- If a 403 occurs (token lacks education scopes), open the Assignments tab inside Teams Web once so the client requests the right scope, then Refresh.

## Architecture

```
teams.microsoft.com tab
  ├─ capture.js   MAIN-world content script: hooks fetch/XHR,
  │               extracts Graph tokens from Authorization headers
  │               + passively snapshots education JSON payloads
  └─ bridge.js    isolated-world script: relays captures via window.postMessage,
                  injects the "A+" FAB that opens the dashboard
        ↓ chrome.runtime messages
assets/background.js  service worker: stores token in chrome.storage.session
                      (memory-only), runs Graph calls ($batch for submissions),
                      caches derived data in chrome.storage.local
        ↓
index.html        React dashboard (Vite) — Material 3 UI
preview.html      same dashboard on a built-in demo dataset (no account needed)
```

Data normalization is shared (`src/shared/normalize.js`), so the demo and live paths run identical code.

## Build & install

```bash
npm install
npm run icons   # one-off: regenerate extension icons (PIL)
npm run build   # → dist/  (esbuild for extension glue + Vite for the React pages)
```

Then load `dist/` as an unpacked extension (`chrome://extensions` → Developer mode → *Load unpacked*).

## Usage

1. Open **teams.microsoft.com** (or `*.cloud.microsoft`) and sign in — browse a little so a Graph token is issued.
2. Click the extension toolbar icon (or the floating **A+** button in Teams) to open the dashboard.
3. Views: **To do / Done / All**, per-class filter chips, sections Overdue → Due today → Next few days → Later.
4. Click a card for the detail sheet: instructions, resources, your grade + teacher feedback, *Open in Teams*.
5. No token / no scopes / expired session each get a dedicated recovery screen.
6. `preview.html` in `dist/` shows the full UI on demo data without any account.

## Security notes

- The token never leaves the browser: stored in `chrome.storage.session` (memory), used only for `graph.microsoft.com`.
- All reads are exactly what the signed-in student can already do in Teams. Nothing is sent to third parties.
- Teacher-authored instructions are whitelist-sanitized before rendering.
- Passive snapshots of assignment JSON are size-capped and only retained as an activity hint in settings.

## Limits

- If the Teams session expires, refresh the Teams tab (a new token is minted) and hit Refresh.
- Passive capture requires the Teams tab to be open at least once per browser session.
- v1.0 `me/assignments` is lean by design; instructions/attachments are fetched per-assignment on demand.
