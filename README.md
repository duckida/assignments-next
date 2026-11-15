<img src="assets/icon.svg" width="64" height="64" alt="">

# Assignments+ — Material UI for Teams Assignments

A **student-side** alternative frontend for Microsoft Teams Assignments, built as a browser extension (Manifest V3) with a React + Vite Material 3 dashboard. It reads your assignments and can turn them in (or undo a turn-in) using the same Teams Web session you already have.

## Research: what's actually going on in Teams Web

The endpoints in the original spec (`edu.microsoft.com`, `ts.accessToken`) don't reflect the current client. Verified against Microsoft's docs:

### Real API surface (Microsoft Graph, delegated student scopes only)

| Need | Endpoint | Notes |
|---|---|---|
| All your assignments, all classes | `GET /v1.0/education/me/assignments` | In **v1.0** since Aug 2024. Least-privilege scope `EduAssignments.ReadBasic`. **Lean payload** — since the 2021 breaking change it omits `instructions`, `resources`, `submissions`. |
| Your classes | `GET /v1.0/education/me/classes` | `EduRoster.ReadBasic`. Fallback: `GET /v1.0/me/joinedTeams` (class team id == education class id). |
| Your submission + grade | `GET /v1.0/education/classes/{c}/assignments/{a}/submissions?$expand=outcomes` | Students only ever see their own submission; `outcomes` carries `educationPointsOutcome` (score) and `educationFeedbackOutcome` (teacher feedback). |
| Full detail (instructions + attachments) | `GET /v1.0/education/classes/{c}/assignments/{a}?$expand=resources` | `educationAssignment.instructions` (itemBody HTML) + `resources`. |
| Turn in / undo | `POST /v1.0/education/classes/{c}/assignments/{a}/submissions/{s}/submit` (and `…/unsubmit`) | Needs `EduAssignments.ReadWrite*`. Same action Teams Web performs; `working → submitted → working`. |
| Attach work (links / files) | `POST .../submissions/{s}/resources` (`educationLinkResource` / `educationFileResource`) | Needs `EduAssignments.ReadWrite*`. File upload: `setUpResourcesFolder` → `PUT {resourcesFolderUrl}:/{name}:/content` → POST `educationFileResource` with the returned item URL. |
| Your attached work | `GET .../submissions/{s}/resources` + `…/submittedResources` | Working set vs. what was handed in. |
| Class display name for a classId | `GET /v1.0/education/classes/{id}?$select=id,displayName` | Used as a batched backfill when the roster list is empty or its ids don't line up with `classId`. |
| Deep link back into Teams | `educationAssignment.webUrl` | Official property — no deep-link reverse engineering needed. |

Relevant enums — assignment `status`: `draft, scheduled, published, assigned, inactive`; submission `status`: `excused, reassigned, returned, submitted, working`.

### Auth reality check

- Teams Web uses **MSAL Browser**; tokens live in `sessionStorage`/`localStorage` under client-id-prefixed keys — and since recent MSAL builds those cache entries are **encrypted** (HKDF key in an `msal.cache.encryption` cookie). Scraping storage is a dead end.
- What still works: the **`Authorization: Bearer` header** the Teams tab itself sends. This extension hooks `fetch`/`XMLHttpRequest` in the page's MAIN world, decodes each JWT, and keeps only tokens whose audience is Microsoft Graph (`aud = 00000003-0000-0000-c000-000000000000`). The token is used exactly as the Teams tab uses it — read-only, on-device, never uploaded anywhere.
- If a 403 occurs (token lacks education scopes), open the Assignments tab inside Teams Web once so the client requests the right scope, then Refresh.

### Signing in without going through Teams

The education scopes (`EduAssignments.*`, `EduRoster.*`) are **admin-consent required** in Entra, and in a school tenant the consent already granted is the one for Teams. So a student cannot mint an equivalent token themselves: no app registration, device-code flow, or Graph Explorer session will be allowed to request those scopes. The token has to come from the Assignments client.

What the extension can do is stop making you drive it:

1. The first time an education-scoped token is seen, the service worker records the URL of the tab **and** of the frame that emitted it (`mad3.eduTrigger.v1`) — that is the deep link that makes the Assignments client run.
2. On any user-initiated sync without a usable token (or after a 403), it opens that URL in a **background tab** (`active: false`), waits up to ~18 s for an education token, then closes the tab and retries — no Teams UI involved.
3. If nothing arrives, the Microsoft session itself has expired: the tab is surfaced once so the sign-in can be completed, which is the one step that cannot be automated.

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
npm test        # service-worker smoke test: mock chrome API, no browser needed
```

Then load `dist/` as an unpacked extension (`chrome://extensions` → Developer mode → *Load unpacked*).

## Usage

1. Open the extension (toolbar icon or the floating **A+** button in Teams). It pulls a session by itself: if there is none, a Teams tab opens in the background on the Assignments deep link. The only manual step ever is signing in to Microsoft when that session has expired.
2. Click the extension toolbar icon (or the floating **A+** button in Teams) to open the dashboard.
3. Views: **Overdue / Upcoming / Done**, per-class filter chips, sections Due today → Next few days → Later. Done is ordered by the latest submission, return, or excuse.
4. Click a card for the detail sheet: instructions (with images rendered inline), resources, your grade + teacher feedback, *Open in Teams*. Graph file links are resolved with the captured session token, so they open as normal SharePoint links instead of failing with an empty Authorization header.
5. **Attach links or files** to your submission right in the detail sheet (add/remove before turning in), then **Turn in** with a two-step confirm. **Undo turn-in** is available while a submission awaits grading.
6. No token / no scopes / expired session each get a dedicated recovery screen.
7. `preview.html` in `dist/` shows the full UI on demo data without any account.

## Performance notes

- Assignments are fetched first; submission lookups and class-name resolution then run concurrently.
- Missing class display names are backfilled from every captured Graph token (education roster, joined teams, group, or team metadata) using case-insensitive IDs. A literal fallback name such as `Class` never masks a real one.
- The dashboard paints from the local cache instantly, and assignment details are cached in the service worker for 15 minutes.
- Detail images lazy-load as you scroll. Graph-hosted read-only images are fetched in the service worker and passed to the page as data URLs.

## Security notes

- The Graph token never leaves the browser: stored in `chrome.storage.session` (memory) and used only for Microsoft Graph. Other service tokens are replayed only to the exact origin that Teams itself sent them to (used for inline images), never to Graph or a third party.
- Graph file links are resolved with that token inside the service worker. The dashboard never puts a bearer token in a URL, DOM attribute, log, or local storage. Short-lived Graph download URLs are used without forwarding the bearer token, and image bytes reach the page as data URLs.
- All calls are exactly what the signed-in student can already do in Teams. Nothing is sent to third parties.
- Turn-in only happens when you press the button — read paths are untouched until then.
- Teacher-authored instructions are whitelist-sanitized before rendering (images allowed from `http(s)` and inline `data:image` only).
- Passive snapshots of assignment JSON are size-capped and only retained as an activity hint in settings.

## Limits

- Access tokens live ~1 hour and refresh tokens sit in MSAL's encrypted cache, so a background tab is spun up roughly once an hour (or whenever Teams' token is gone). That costs a few seconds, not a click.
- The very first sign-in, and any time the Microsoft session expires, still needs the surfaced Teams tab once.
- Passive capture requires a Teams tab to be open at least once per browser session; after that the background refresh replaces it.
- v1.0 `me/assignments` is lean by design; instructions/attachments are fetched per-assignment on demand.
- Attachments upload through the extension: files are capped at 20 MB, and the assignment must allow student-added resources (Teams enforces the same rule).
