/**
 * Service worker — owns captured session tokens (storage.session, memory-only)
 * and performs read-only reads on behalf of the dashboard.
 *
 * Two strategies, picked from whatever audience tokens the Teams tab actually sends:
 *   graph — Microsoft Graph education endpoints (preferred)
 *   edu   — edu.microsoft.com internal assignments API (fallback, the one visible
 *           in Teams Web's own network tab)
 * Every captured token's audience is logged (sans token) so the settings panel can
 * show exactly what the page emitted.
 */
import { MSG } from '../shared/protocol.js'
import { GraphError, graphBatch, graphFetch, graphPaged } from './graph.js'

const CANDIDATES_KEY = 'mad3.session.tokens.v2'
const SEEN_KEY = 'mad3.session.seen.v1'
const CACHE_KEY = 'mad3.cache.v2' // v2: rows now carry className; v1 lacked it
const SNAPSHOT_KEY = 'mad3.snapshot.v1'
const LEGACY_LOCAL_KEYS = ['mad3.cache.v1']
const DETAILS_KEY = 'mad3.details.v2' // per-assignment detail cache (fast re-open)
const DETAIL_TTL = 15 * 60_000
const DETAILS_CAP = 80
const VISIBLE = new Set(['assigned', 'published', 'inactive', 'unknownfuturevalue'])

// Background sign-in: the edu-scoped Graph token only exists once the Assignments
// client has run inside Teams, so we remember which URL produced one and replay it
// in a hidden tab instead of making the user go through Teams.
const TRIGGER_KEY = 'mad3.eduTrigger.v1' // { candidates: [{ url }], ts }
const REFRESH_KEY = 'mad3.session.refreshAt.v1'
const REFRESH_COOLDOWN = 45_000
const REFRESH_BUDGET = [18_000, 12_000] // per candidate URL, ms
const TEAMS_URL_RE = /^https:\/\/(?:[a-z0-9-]+\.)*(?:teams\.microsoft\.com|teams\.live\.com|cloud\.microsoft|edu\.microsoft\.com)\//i
const TEAMS_HOME = 'https://teams.microsoft.com/'

const MATCH_URLS = [
  'https://teams.microsoft.com/*',
  'https://*.teams.microsoft.com/*',
  'https://*.cloud.microsoft/*',
  'https://teams.live.com/*',
  'https://edu.microsoft.com/*',
]

const isGraphAud = (aud) =>
  aud === '00000003-0000-0000-c000-000000000000' || /graph\.microsoft\.(com|us)/i.test(String(aud))
const isEduAud = (aud) => /edu\.microsoft\.com/i.test(String(aud))

/* ---------------- lifecycle: inject into already-open tabs ---------------- */

chrome.runtime.onInstalled.addListener(() => {
  try {
    chrome.storage?.session?.setAccessLevel?.({ accessLevel: 'TRUSTED_CONTEXTS' })?.catch?.(() => {})
  } catch {
    /* not supported */
  }
  reinject()
})
chrome.runtime.onStartup?.addListener?.(reinject)

// Content scripts only run on page load; inject into tabs that were open before install.
async function reinject() {
  try {
    if (!chrome.scripting?.executeScript) return
    const tabs = await chrome.tabs.query({ url: MATCH_URLS })
    for (const tab of tabs) {
      const target = { tabId: tab.id, allFrames: true }
      await chrome.scripting.executeScript({ target, files: ['assets/capture.js'], world: 'MAIN' }).catch(() => {})
      await chrome.scripting.executeScript({ target, files: ['assets/bridge.js'] }).catch(() => {})
    }
  } catch {
    /* ignore */
  }
}

chrome.action.onClicked.addListener(() => openDashboard())

function openDashboard() {
  const url = chrome.runtime.getURL('index.html')
  chrome.tabs.query({ url: url + '*' }).then((tabs) => {
    if (tabs && tabs.length) {
      chrome.tabs.update(tabs[0].id, { active: true })
      chrome.windows?.update?.(tabs[0].windowId, { focused: true })
    } else {
      chrome.tabs.create({ url })
    }
  })
}

/* ---------------- captured token store ---------------- */

async function saveToken(p) {
  if (!p?.token || !p.aud) return
  const store = await chrome.storage.session.get(CANDIDATES_KEY)
  const list = (store?.[CANDIDATES_KEY] || []).filter((c) => !(c.aud === p.aud && c.expMs === p.expMs))
  list.unshift({ token: p.token, aud: p.aud, scp: p.scp || '', upn: p.upn || '', host: p.host || '', expMs: p.expMs || 0, capturedAt: p.capturedAt || Date.now() })
  // graph-audience first, then newest expiry
  list.sort((a, b) => (isGraphAud(b.aud) ? 1 : 0) - (isGraphAud(a.aud) ? 1 : 0) || (b.expMs || 0) - (a.expMs || 0))
  await chrome.storage.session.set({ [CANDIDATES_KEY]: list.slice(0, 10) })
  // aud-only diagnostics (no tokens)
  const seenStore = await chrome.storage.session.get(SEEN_KEY)
  const seen = seenStore?.[SEEN_KEY] || []
  if (!seen.some((s) => s.aud === p.aud)) {
    seen.push({ aud: p.aud, host: p.host || '', expMs: p.expMs || 0, capturedAt: Date.now() })
    await chrome.storage.session.set({ [SEEN_KEY]: seen.slice(0, 20) })
  }
}

async function getCandidates() {
  const store = await chrome.storage.session.get([CANDIDATES_KEY, SEEN_KEY])
  const now = Date.now()
  const list = (store?.[CANDIDATES_KEY] || []).filter((c) => !c.expMs || c.expMs > now)
  return { candidates: list, seen: store?.[SEEN_KEY] || [] }
}

function makePicks(candidates) {
  const picks = []
  const add = (strategy, c) => {
    if (!c) return
    const key = `${strategy}|${c.aud}|${c.expMs}|${c.token.length}`
    if (!picks.some((p) => p.key === key)) picks.push({ key, token: c.token, strategy, meta: c })
  }

  // Graph-audience tokens are ideal when they include EduAssignments/EduRoster.
  for (const c of candidates.filter((x) => isGraphAud(x.aud))) add('graph', c)
  // Teams EDU also commonly calls edu.microsoft.com internal endpoints with an edu-audience token.
  for (const c of candidates.filter((x) => isEduAud(x.aud))) add('edu', c)
  // Rare case: a non-standard audience token still carries education scopes.
  for (const c of candidates.filter((x) => /EduAssignments|EduRoster/i.test(x.scp || ''))) add('graph', c)

  return picks
}

/* ---------------- background sign-in (no Teams UI) ---------------- */

/** Tokens that can actually read education data: edu scopes, or the edu audience. */
const hasEduScopes = (p) => /EduAssignments|EduRoster/i.test(p?.scp || '')
const resolvesRefresh = (p) => hasEduScopes(p) || isEduAud(p?.aud)
const hasUsableTokenSource = (candidates) => candidates.some((c) => hasEduScopes(c) || isEduAud(c.aud))

const isTeamsUrl = (url) => typeof url === 'string' && url.length <= 2048 && TEAMS_URL_RE.test(url)

/** Prefer URLs that look like they open the Assignments experience. */
function triggerScore(url) {
  let score = 0
  if (/\/school\/assignments/i.test(url)) score += 6
  if (/assignments/i.test(url)) score += 3
  if (/\/l\/entity\//i.test(url)) score += 2
  if (/#\/[^?#]*assign/i.test(url)) score += 2
  if (/\/v2\/?(?:$|[?#])/i.test(url)) score -= 2
  return score
}

/** Remember the URL(s) that emitted an education-scoped token. */
async function rememberTrigger(sender, payload) {
  let tabUrl = ''
  try {
    if (sender?.tab?.id != null) tabUrl = (await chrome.tabs.get(sender.tab.id))?.url || sender.tab.url || ''
  } catch {
    /* tab closed already */
  }
  const frameUrl = typeof payload?.pageUrl === 'string' ? payload.pageUrl : ''
  const fresh = [tabUrl, frameUrl].filter(isTeamsUrl)
  if (!fresh.length) return
  const store = await chrome.storage.local.get(TRIGGER_KEY)
  const prev = store?.[TRIGGER_KEY]?.candidates || []
  const all = []
  const add = (url) => {
    if (!isTeamsUrl(url) || all.some((c) => c.url === url)) return
    all.push({ url })
  }
  fresh.forEach(add)
  prev.forEach((c) => add(c.url))
  all.sort((a, b) => triggerScore(b.url) - triggerScore(a.url))
  const candidates = all.slice(0, 3)
  if (candidates.map((c) => c.url).join('\n') === prev.map((c) => c.url).join('\n')) return
  await chrome.storage.local.set({ [TRIGGER_KEY]: { candidates, ts: Date.now() } })
}

/** Trigger URLs to replay, best first: learned, then a cached assignment deep link. */
async function getTriggerUrls() {
  const urls = []
  let learned = false
  const add = (u) => {
    if (isTeamsUrl(u) && !urls.includes(u)) urls.push(u)
  }
  try {
    const store = await chrome.storage.local.get([TRIGGER_KEY, CACHE_KEY])
    const remembered = store?.[TRIGGER_KEY]?.candidates || []
    learned = remembered.length > 0
    for (const c of remembered) add(c.url)
    for (const a of store?.[CACHE_KEY]?.assignments || []) {
      if (urls.length > 1) break
      if (!isTeamsUrl(a?.webUrl)) continue
      learned = true
      add(a.webUrl) // educationAssignment.webUrl deep-links straight into the app
    }
  } catch {
    /* fall through to Teams home */
  }
  if (!urls.length) urls.push(TEAMS_HOME)
  return { urls: urls.slice(0, REFRESH_BUDGET.length), learned }
}

/* Waiters resolved when the Assignments client mints a usable token. */
const eduWaiters = new Set()

function notifyEduToken() {
  for (const w of [...eduWaiters]) w()
}

function waitForCapturedEduToken(ms) {
  if (!(ms > 0)) return Promise.resolve(false)
  return new Promise((resolve) => {
    let timer = 0
    const onToken = () => {
      clearTimeout(timer)
      eduWaiters.delete(onToken)
      resolve(true)
    }
    timer = setTimeout(() => {
      eduWaiters.delete(onToken)
      resolve(false)
    }, ms)
    eduWaiters.add(onToken)
  })
}

let refreshInFlight = null

function backgroundSignIn(opts = {}) {
  if (refreshInFlight) return refreshInFlight
  const p = runBackgroundSignIn(opts).finally(() => {
    if (refreshInFlight === p) refreshInFlight = null
  })
  refreshInFlight = p
  return p
}

/**
 * Open the Assignments deep link in a background tab, wait for the education
 * token, then close it. If nothing arrives (expired Microsoft session) the tab is
 * surfaced so the sign-in can be completed once — that part cannot be automated.
 */
async function runBackgroundSignIn({ force = false } = {}) {
  const { candidates } = await getCandidates()
  if (!force && hasUsableTokenSource(candidates)) return { attempted: false, ok: true, reused: true }

  const { urls, learned } = await getTriggerUrls()
  if (!learned) {
    // Cold start: the deep link is unknown, so a hidden tab could only show the
    // Teams home without ever requesting education scopes. Surface it instead.
    await openTeams()
    return { attempted: true, ok: false, code: 'no-trigger', url: urls[0] }
  }

  const last = await chrome.storage.session.get(REFRESH_KEY).catch(() => null)
  const lastAt = last?.[REFRESH_KEY] || 0
  if (!force && Date.now() - lastAt < REFRESH_COOLDOWN) return { attempted: false, ok: false, code: 'cooldown' }
  await chrome.storage.session.set({ [REFRESH_KEY]: Date.now() }).catch(() => {})

  let tab = null
  for (let i = 0; i < urls.length; i++) {
    try {
      tab = tab ? await chrome.tabs.update(tab.id, { url: urls[i] }) : await chrome.tabs.create({ url: urls[i], active: false })
    } catch {
      continue
    }
    if (!tab?.id) continue
    const budget = REFRESH_BUDGET[i] ?? REFRESH_BUDGET[REFRESH_BUDGET.length - 1]
    if (await waitForCapturedEduToken(budget)) {
      await chrome.tabs.remove(tab.id).catch(() => {})
      return { attempted: true, ok: true, url: urls[i] }
    }
  }
  if (!tab?.id) return { attempted: true, ok: false, code: 'error' }
  // Surface the tab only when we genuinely could not get a token.
  await chrome.tabs.update(tab.id, { active: true }).catch(() => {})
  try {
    await chrome.windows?.update?.(tab.windowId, { focused: true })
  } catch {
    /* ignore */
  }
  return { attempted: true, ok: false, code: 'needs-login', tabId: tab.id, url: urls[0] }
}

/** Open (or focus) Teams at the learned Assignments URL — used by recovery buttons. */
async function openTeams() {
  const { urls } = await getTriggerUrls()
  const url = urls[0] || TEAMS_HOME
  const tabs = await chrome.tabs.query({ url: MATCH_URLS }).catch(() => [])
  const hit = (tabs || []).find((t) => t.url === url) || (tabs || [])[0]
  if (hit?.id != null) {
    await chrome.tabs.update(hit.id, { active: true }).catch(() => {})
    try {
      await chrome.windows?.update?.(hit.windowId, { focused: true })
    } catch {}
    return { ok: true, url, reused: true }
  }
  await chrome.tabs.create({ url, active: true }).catch(() => {})
  return { ok: true, url }
}

/* ---------------- graph strategy ---------------- */

async function getClasses(token) {
  // Both roster reads are independent — run them together and prefer education classes.
  const [classesRes, teamsRes] = await Promise.allSettled([
    graphPaged(token, '/education/me/classes?$select=id,displayName,description&$top=100', {}),
    graphPaged(token, '/me/joinedTeams?$select=id,displayName,description,webUrl&$top=200', {}),
  ])
  const classes = classesRes.status === 'fulfilled' ? classesRes.value : []
  if (classes.length) return classes.map((c) => ({ id: c.id, name: c.displayName || 'Class', description: c.description, webUrl: null }))
  const teams = teamsRes.status === 'fulfilled' ? teamsRes.value : []
  if (teams.length) return teams.map((t) => ({ id: t.id, name: t.displayName || 'Team', description: t.description, webUrl: t.webUrl || null }))
  return null
}

async function graphSync(pick, meta) {
  const token = pick.token
  // Roster + assignment list are independent reads: fetch them at the same time.
  let classes, rawAssignments
  try {
    ;[classes, rawAssignments] = await Promise.all([
      getClasses(token).catch(() => null),
      graphPaged(token, '/education/me/assignments?$top=200', { version: 'v1.0' }),
    ])
  } catch (e) {
    return graphFail(e, meta)
  }
  classes = (classes || []).slice()
  const assignments = rawAssignments.filter((a) => VISIBLE.has(String(a.status || '').toLowerCase()))
  const classNames = new Map(classes.map((c) => [c.id, c.name]))

  // The v1.0 assignment payload carries no class display name, and the roster call
  // can come back empty (EduRoster scope missing) or use ids that differ from the
  // classIds on assignments. Resolve any missing names directly, one batch call.
  const missing = [...new Set(assignments.map((a) => a.classId).filter((cid) => !classNames.get(cid)))]
  if (missing.length) {
    const lookups = await safeBatch(token, missing.map((cid) => `/education/classes/${cid}?$select=id,displayName,description`))
    for (const body of lookups.values()) {
      if (!body?.id || !body.displayName || classNames.has(body.id)) continue
      classNames.set(body.id, body.displayName)
      classes.push({ id: body.id, name: body.displayName, description: body.description || '', webUrl: null })
    }
  }

  const named = assignments.map((a) => ({ ...a, className: classNames.get(a.classId) || a.classDisplayName || null }))
  const subPaths = named.map((a) => `/education/classes/${a.classId}/assignments/${a.id}/submissions?$expand=outcomes`)
  const subs = await safeBatch(token, subPaths)
  const submissions = {}
  for (const [path, body] of subs) {
    const m = /\/assignments\/([a-z0-9-]+)\/submissions/i.exec(path)
    const val = body?.value
    if (m && Array.isArray(val) && val.length) submissions[m[1]] = val[0]
  }
  return { ok: true, data: { classes, assignments: named, submissions, ts: Date.now() }, meta: { ...meta, strategy: 'graph' } }
}

async function safeBatch(token, paths) {
  try {
    return await graphBatch(token, paths)
  } catch (e) {
    if (e instanceof GraphError && e.status === 0) throw e
    return new Map()
  }
}

function graphFail(e, meta) {
  return {
    ok: false,
    code: e.status === 401 ? 'expired' : e.status === 403 ? 'forbidden' : 'error',
    meta: { ...meta, status: e.status || 0, graphCode: e.code || '', message: e.message, hint: e.hint || '' },
  }
}

function attemptLabel(pick) {
  return `${pick.strategy}:${String(pick.meta.aud || '').replace(/^https?:\/\//, '').slice(0, 72)}`
}

function attemptFromResult(pick, res) {
  return {
    strategy: pick.strategy,
    aud: String(pick.meta.aud || ''),
    status: res?.meta?.status || 0,
    code: res?.code || 'error',
    graphCode: res?.meta?.graphCode || '',
    message: res?.meta?.message || res?.meta?.hint || 'Failed',
  }
}

function allFailed(picks, seen, attempts) {
  const forbidden = attempts.some((a) => a.code === 'forbidden' || a.status === 403)
  const expired = attempts.length && attempts.every((a) => a.code === 'expired' || a.status === 401)
  const summary = attempts
    .map((a) => `${a.strategy} ${a.status || a.code}${a.graphCode ? ` ${a.graphCode}` : ''}`)
    .join('; ')
  return {
    ok: false,
    code: expired ? 'expired' : forbidden ? 'forbidden' : 'error',
    meta: {
      seen: seen.map((s) => s.aud),
      attempts,
      strategy: attempts.map((a) => a.strategy).join('→'),
      message: summary || 'No usable education API token succeeded.',
      hint:
        'Captured token(s) were present but none could read assignments. In Teams Web, open Assignments, wait for the assignment list to load, then Refresh here. If it still fails, open Settings and copy the attempts/seen values.',
    },
  }
}

/* ---------------- edu.microsoft.com strategy (fallback) ---------------- */

const EDU_BASE = 'https://edu.microsoft.com/api/v1.0'

async function eduGet(token, path) {
  const res = await fetch(EDU_BASE + path, { headers: { Authorization: `Bearer ${token}` } })
  if (!res.ok) throw new GraphError(String(res.status), res.status, res.statusText)
  return res.json().catch(() => null)
}

const arrOf = (b) => (Array.isArray(b) ? b : Array.isArray(b?.value) ? b.value : Array.isArray(b?.teams) ? b.teams : [])

function eduMapAssignment(a, teamId) {
  const id = a?.id || a?.assignmentId
  if (!id) return null
  const maxPts = a?.maxPoints ?? a?.points
  return {
    id,
    classId: teamId,
    displayName: a.displayName || a.name || a.title || 'Assignment',
    dueDateTime: a.dueDateTime || a.dueDate || null,
    closeDateTime: a.closeDateTime || a.closeDate || null,
    assignedDateTime: a.assignedDateTime || null,
    status: 'assigned',
    allowLateSubmissions: a.allowLateSubmissions ?? null,
    grading: maxPts != null ? { '@odata.type': '#microsoft.graph.educationAssignmentPointsGradeType', maxPoints: maxPts } : undefined,
    webUrl: a.webUrl || a.teamsWebUrl || null,
    lastModifiedDateTime: a.lastModifiedDateTime || null,
  }
}

function eduMapSubmission(s, assignmentId) {
  if (!s) return null
  const outcomes = []
  const grade = s.grade ?? s.pointsGrade ?? s.points
  if (grade != null) {
    outcomes.push({ '@odata.type': '#microsoft.graph.educationPointsOutcome', points: { points: grade, publishedPoints: grade } })
  }
  const fb = s.feedback?.text || s.feedback || null
  if (fb) outcomes.push({ '@odata.type': '#microsoft.graph.educationFeedbackOutcome', feedback: { text: fb } })
  return {
    id: s.id || '',
    status: s.status || 'working',
    submittedDateTime: s.submittedDateTime || null,
    returnedDateTime: s.returnedDateTime || null,
    outcomes,
  }
}

async function pool(items, size, fn) {
  const queue = [...items]
  await Promise.all(
    Array.from({ length: size }, async () => {
      while (queue.length) await fn(queue.shift())
    })
  )
}

async function eduSync(pick, meta) {
  const token = pick.token
  let teamsBody
  try {
    teamsBody = await eduGet(token, '/teams')
  } catch (e) {
    return graphFail(e, { ...meta, hint: 'edu.microsoft.com rejected the captured token. Reload the Teams tab and retry.' })
  }
  const teams = arrOf(teamsBody)
    .map((t) => ({ id: t.id || t.teamId, name: t.displayName || t.name || 'Class', description: t.description || '', webUrl: t.webUrl || null }))
    .filter((t) => t.id)

  const assignments = []
  const submissions = {}
  await pool(teams.slice(0, 12), 5, async (t) => {
    try {
      const body = await eduGet(token, `/teams/${t.id}/assignments`)
      for (const a of arrOf(body)) {
        const raw = eduMapAssignment(a, t.id)
        if (!raw) continue
        raw.className = t.name
        assignments.push(raw)
        try {
          const sBody = await eduGet(token, `/teams/${t.id}/assignments/${raw.id}/submissions/me`)
          const one = Array.isArray(sBody?.value) ? sBody.value[0] : sBody
          const mapped = eduMapSubmission(one, raw.id)
          if (mapped) submissions[raw.id] = mapped
        } catch {
          /* per-assignment submission may not exist yet */
        }
      }
    } catch {
      /* team may be inaccessible */
    }
  })
  return { ok: true, data: { classes: teams, assignments, submissions, ts: Date.now() }, meta: { ...meta, strategy: 'edu' } }
}

/* ---------------- link titles ---------------- */

const TITLES_KEY = 'mad3.titles.v1'
const TITLE_MAX_AGE = 24 * 3_600_000

function decodeEntities(s) {
  return s
    .replace(/&#(\d+);/g, (_, n) => String.fromCodePoint(Number(n)))
    .replace(/&#x([0-9a-f]+);/gi, (_, n) => String.fromCodePoint(parseInt(n, 16)))
    .replace(/&(amp|lt|gt|quot|apos|nbsp);/g, (_, e) => ({ amp: '&', lt: '<', gt: '>', quot: '"', apos: "'", nbsp: ' ' }[e]))
}

async function fetchTitle(url) {
  if (!/^https?:\/\//i.test(String(url))) return { ok: false, title: null }
  const store = await chrome.storage.local.get(TITLES_KEY)
  const cache = store?.[TITLES_KEY] || {}
  const hit = cache[url]
  if (hit && Date.now() - hit.ts < TITLE_MAX_AGE) return { ok: !!hit.title, title: hit.title }
  let title = null
  try {
    const res = await fetch(url, { redirect: 'follow', signal: AbortSignal.timeout(9000) })
    if (res.ok && (res.headers.get('content-type') || '').includes('text/html')) {
      const reader = res.body.getReader()
      let text = ''
      const dec = new TextDecoder()
      for (let bytes = 0; bytes < 262_144; ) {
        const { done, value } = await reader.read()
        if (done) break
        bytes += value.length
        text += dec.decode(value, { stream: true })
        if (/<\/title>/i.test(text)) break
      }
      try {
        reader.cancel()
      } catch {}
      const m = /<title[^>]*>([\s\S]*?)<\/title>/i.exec(text)
      if (m) title = decodeEntities(m[1].replace(/\s+/g, ' ')).trim().slice(0, 180) || null
    }
  } catch {
    /* page unreachable / blocked — keep displayName fallback */
  }
  cache[url] = { title, ts: Date.now() }
  const keys = Object.keys(cache)
  if (keys.length > 400) delete cache[keys[0]]
  chrome.storage.local.set({ [TITLES_KEY]: cache }).catch(() => {})
  return { ok: !!title, title }
}

/* ---------------- detail cache (instant re-open) ---------------- */

async function readDetailCache() {
  try {
    const store = await chrome.storage.local.get(DETAILS_KEY)
    const map = store?.[DETAILS_KEY] || {}
    const now = Date.now()
    for (const k of Object.keys(map)) {
      if (!map[k] || now - (map[k].ts || 0) > DETAIL_TTL) delete map[k]
    }
    return { map }
  } catch {
    return { map: {} }
  }
}

async function writeDetailCache(map, key, assignment, strategy) {
  map[key] = { ts: Date.now(), strategy, assignment }
  const keys = Object.keys(map).sort((a, b) => (map[a].ts || 0) - (map[b].ts || 0))
  while (keys.length > DETAILS_CAP) delete map[keys.shift()]
  try {
    await chrome.storage.local.set({ [DETAILS_KEY]: map })
  } catch {
    /* best effort */
  }
}

/* ---------------- sync / detail / submit / clear ---------------- */

let syncInFlight = null

async function runPicks({ candidates, seen }) {
  const picks = makePicks(candidates)
  if (!picks.length) {
    return {
      ok: false,
      code: 'no-token',
      meta: {
        seen: seen.map((s) => s.aud),
        hint: 'No Teams session is available yet. Teams is opened in a background tab for you; if it asks you to sign in, finish that once and hit Refresh.',
      },
    }
  }
  const attempts = []
  for (const pick of picks) {
    const meta = {
      capturedAt: pick.meta.capturedAt,
      expMs: pick.meta.expMs,
      scp: pick.meta.scp,
      upn: pick.meta.upn,
      seen: seen.map((s) => s.aud),
      attempt: attemptLabel(pick),
    }
    const res = pick.strategy === 'graph' ? await graphSync(pick, meta) : await eduSync(pick, meta)
    if (res?.ok) {
      // Treat zero assignments as a successful API read; some users may genuinely have none.
      return { ...res, meta: { ...(res.meta || {}), attempts: [...attempts, { strategy: pick.strategy, aud: pick.meta.aud, status: 200, code: 'ok' }] } }
    }
    attempts.push(attemptFromResult(pick, res))
  }
  return allFailed(picks, seen, attempts)
}

/** 401/403 mean the session exists but lacks what we need → a refresh can fix it. */
function isAuthFailure(res) {
  if (!res || res.ok) return false
  if (res.code === 'forbidden' || res.code === 'expired') return true
  return (res.meta?.attempts || []).some((a) => a.status === 401 || a.status === 403)
}

async function doSync({ auto = false } = {}) {
  let auth = null
  let state = await getCandidates()
  if (auto && !hasUsableTokenSource(state.candidates)) {
    auth = await backgroundSignIn()
    if (auth?.ok) state = await getCandidates()
  }
  let res = await runPicks(state)
  // An existing token can still turn out to lack the education scopes (403) — that
  // is the other case the Assignments client fixes, so retry once behind a tab.
  if (!auth && auto && isAuthFailure(res)) {
    auth = await backgroundSignIn({ force: true })
    if (auth?.ok) res = await runPicks(await getCandidates())
  }
  if (!res.ok) res.meta = { ...(res.meta || {}), autoRefresh: auth || { attempted: false } }
  return res
}

/** Coalesce concurrent syncs (initial paint + manual refresh) onto one request. */
function sharedSync(opts = {}) {
  if (!syncInFlight) {
    syncInFlight = doSync(opts)
      .then(async (res) => {
        if (res?.ok) {
          const cache = { ts: res.data.ts, classes: res.data.classes, assignments: res.data.assignments, submissions: res.data.submissions }
          await chrome.storage.local.set({ [CACHE_KEY]: cache }).catch(() => {})
        }
        return res
      })
      .catch(() => ({ ok: false, code: 'error' }))
      .finally(() => {
        syncInFlight = null
      })
  }
  return syncInFlight
}

async function doDetail({ classId, id }) {
  const key = `${classId}|${id}`
  const { map } = await readDetailCache()
  if (map[key]?.assignment) {
    return { ok: true, assignment: map[key].assignment, meta: { strategy: map[key].strategy || 'graph', cached: true } }
  }
  const { candidates, seen } = await getCandidates()
  const picks = makePicks(candidates)
  if (!picks.length) return { ok: false, code: 'no-token', meta: { seen: seen.map((s) => s.aud) } }
  const attempts = []
  const sameTenant = (p) => p.strategy === 'graph' || String(p.meta.aud || '').includes('edu.microsoft.com')
  for (const pick of picks.filter(sameTenant)) {
    const meta = { upn: pick.meta.upn, strategy: pick.strategy, seen: seen.map((s) => s.aud), attempt: attemptLabel(pick) }
    try {
      if (pick.strategy === 'graph') {
        const base = `/education/classes/${classId}/assignments/${id}`
        let a
        try {
          a = await graphFetch(pick.token, `${base}?$expand=resources`)
        } catch (e) {
          if (e.status === 400) a = await graphFetch(pick.token, base)
          else throw e
        }
        writeDetailCache(map, key, a, 'graph').catch(() => {})
        return { ok: true, assignment: a, meta }
      }
      const a = await eduGet(pick.token, `/teams/${classId}/assignments/${id}`)
      const raw = eduMapAssignment(a, classId) || { id, classId }
      raw.instructions = { content: a?.instructions || a?.description || '' }
      raw.resources = Array.isArray(a?.resources)
        ? a.resources.map((r) => ({ id: r.id || r.name, displayName: r.displayName || r.name || 'Resource', resource: { link: r.webUrl || r.link || null } }))
        : []
      writeDetailCache(map, key, raw, 'edu').catch(() => {})
      return { ok: true, assignment: raw, meta }
    } catch (e) {
      attempts.push(attemptFromResult(pick, graphFail(e, meta)))
    }
  }
  return allFailed(picks, seen, attempts)
}

/** Mark an assignment turned in (submit) or reclaim it (unsubmit).
 *  Needs a Graph-audience token with EduAssignments.ReadWrite*; the Teams session
 *  normally has it because Teams Web performs the same action.
 */
async function doSubmit({ classId, id, submissionId, action = 'submit' }) {
  const act = action === 'unsubmit' ? 'unsubmit' : 'submit'
  const { candidates, seen } = await getCandidates()
  const picks = makePicks(candidates).filter((p) => p.strategy === 'graph')
  if (!picks.length) {
    return {
      ok: false,
      code: 'no-token',
      error: 'Turning in needs a Graph session token. Open Teams Web once, then retry.',
      meta: { seen: seen.map((s) => s.aud) },
    }
  }
  const attempts = []
  let hint = ''
  for (const pick of picks) {
    const meta = { upn: pick.meta.upn, strategy: 'graph', seen: seen.map((s) => s.aud), attempt: attemptLabel(pick) }
    try {
      let sid = submissionId || null
      if (!sid) {
        const body = await graphFetch(pick.token, `/education/classes/${classId}/assignments/${id}/submissions`)
        sid = Array.isArray(body?.value) && body.value.length ? body.value[0].id : null
      }
      if (!sid) throw new GraphError('no-submission', 404, 'No submission exists for this assignment yet.')
      const submission = await graphFetch(pick.token, `/education/classes/${classId}/assignments/${id}/submissions/${sid}/${act}`, { method: 'POST' })
      return { ok: true, submission: submission || null, meta: { ...meta, action: act } }
    } catch (e) {
      hint =
        e.status === 403
          ? 'Your Teams session token lacks EduAssignments.ReadWrite. Open the Assignments tab in Teams Web once (so it requests write access), then retry.'
          : e.hint || hint
      attempts.push(attemptFromResult(pick, graphFail(e, meta)))
    }
  }
  const first = attempts[0] || {}
  const base = first.message && first.message !== first.graphCode ? first.message : ''
  return {
    ok: false,
    error: [base, hint].filter(Boolean).join(' — ') || `Couldn’t ${act === 'submit' ? 'turn in' : 'undo the turn-in'}. Check your Teams session and try again.`,
    meta: { attempts },
  }
}

/** Student work on a submission: list / add link / upload file / remove.
 *  File uploads follow Graph's documented flow: setUpResourcesFolder →
 *  PUT {resourcesFolderUrl}:/{name}:/content → POST educationFileResource.
 */
const MAX_UPLOAD_B64 = 28 * 1024 * 1024 // ≈20 MB of file content

async function submissionIdFor(token, classId, assignmentId, submissionId) {
  if (submissionId) return submissionId
  const body = await graphFetch(token, `/education/classes/${classId}/assignments/${assignmentId}/submissions`)
  return Array.isArray(body?.value) && body.value.length ? body.value[0].id : null
}

function mapResourceList(json) {
  const arr = Array.isArray(json?.value) ? json.value : []
  return arr.map((r) => ({
    id: r.id,
    displayName: r.resource?.displayName || 'Resource',
    type: String(r.resource?.['@odata.type'] || ''),
    link: r.resource?.link || null,
    fileUrl: r.resource?.fileUrl || null,
    studentAdded: !r.assignmentResourceUrl,
    createdTs: r.resource?.createdDateTime ? Date.parse(r.resource.createdDateTime) : null,
  }))
}

const hostOf = (url) => {
  try {
    return new URL(url).hostname.replace(/^www\./, '')
  } catch {
    return 'Link'
  }
}

async function addFileResource(token, base, sid, payload) {
  const name = String(payload.name || '').trim().slice(0, 180)
  const b64 = String(payload.dataBase64 || '')
  if (!name) return { ok: false, error: 'Missing file name.' }
  if (!b64) return { ok: false, error: 'The selected file is empty.' }
  if (b64.length > MAX_UPLOAD_B64) return { ok: false, error: 'File is too large — the limit is 20 MB.' }
  let bytes
  try {
    bytes = Uint8Array.from(atob(b64), (c) => c.charCodeAt(0))
  } catch {
    return { ok: false, error: 'Could not read the file data.' }
  }

  // 1) the submission's own SharePoint folder (created on first upload)
  let folderUrl = null
  try {
    const sub = await graphFetch(token, `${base}/${sid}`)
    folderUrl = sub?.resourcesFolderUrl || null
  } catch {
    /* fall through to setUpResourcesFolder */
  }
  if (!folderUrl) {
    const created = await graphFetch(token, `${base}/${sid}/setUpResourcesFolder`, { method: 'POST' })
    folderUrl = created?.resourcesFolderUrl || null
    if (!folderUrl) {
      const sub = await graphFetch(token, `${base}/${sid}`)
      folderUrl = sub?.resourcesFolderUrl || null
    }
  }
  if (!folderUrl) throw new GraphError('no-folder', 500, 'Could not create a submission folder for this assignment.')
  const m = /\/drives\/([^/]+)\/items\/([^/?#]+)/i.exec(folderUrl)
  if (!m) throw new GraphError('bad-folder', 500, 'Unexpected resources folder URL from Graph.')

  // 2) upload the bytes, then register the file as a submission resource
  const item = await graphFetch(token, `${folderUrl}:/${encodeURIComponent(name)}:/content`, {
    method: 'PUT',
    rawBody: bytes,
    contentType: payload.mime || 'application/octet-stream',
  })
  if (!item?.id) throw new GraphError('upload-failed', 500, 'The file upload did not complete.')
  const fileUrl = `https://graph.microsoft.com/v1.0/drives/${m[1]}/items/${item.id}`

  // 3) re-upload case: the PUT above replaced the bytes in the folder, so drop the
  //    older resource entry for the same name before registering the new one —
  //    Graph rejects a second file resource for the same file.
  try {
    const list = await graphFetch(token, `${base}/${sid}/resources`)
    for (const r of mapResourceList(list)) {
      if (r.studentAdded && r.displayName === name) {
        await graphFetch(token, `${base}/${sid}/resources/${r.id}`, { method: 'DELETE' })
      }
    }
  } catch {
    /* best effort */
  }
  const created = await graphFetch(token, `${base}/${sid}/resources`, {
    method: 'POST',
    body: { resource: { '@odata.type': '#microsoft.graph.educationFileResource', displayName: name, fileUrl } },
  })
  return { ok: true, resource: mapResourceList({ value: [created] })[0] || null }
}

async function doSubmission(payload) {
  const { op = 'list', classId, assignmentId } = payload
  if (!classId || !assignmentId) return { ok: false, error: 'Missing assignment reference.' }
  const { candidates, seen } = await getCandidates()
  const picks = makePicks(candidates).filter((p) => p.strategy === 'graph')
  if (!picks.length) {
    return { ok: false, error: 'Attaching work needs a Graph session token. Open Teams Web once, then retry.', meta: { seen: seen.map((s) => s.aud) } }
  }
  const base = `/education/classes/${classId}/assignments/${assignmentId}/submissions`
  const attempts = []
  let hint = ''
  for (const pick of picks) {
    const meta = { upn: pick.meta.upn, strategy: 'graph', seen: seen.map((s) => s.aud), attempt: attemptLabel(pick) }
    try {
      // Opening a file doesn't need a submission (teacher resources have none yet).
      if (op === 'openFile') {
        // Submission/assignment file resources point at Graph's drive API, which a
        // browser tab can't open without an auth header. Resolve the SharePoint webUrl.
        const fileUrl = String(payload.fileUrl || '')
        if (/^https:\/\/graph\.microsoft\.com\/v1\.0\/drives\//i.test(fileUrl)) {
          const item = await graphFetch(pick.token, `${fileUrl}?$select=webUrl,name`)
          if (!item?.webUrl) return { ok: false, error: 'This file has no web link — open it in Teams instead.' }
          await chrome.tabs.create({ url: item.webUrl })
          return { ok: true }
        }
        if (/^https?:\/\//i.test(fileUrl)) {
          await chrome.tabs.create({ url: fileUrl })
          return { ok: true }
        }
        return { ok: false, error: 'Unsupported file link.' }
      }

      const sid = await submissionIdFor(pick.token, classId, assignmentId, payload.submissionId)
      if (!sid) throw new GraphError('no-submission', 404, 'No submission exists for this assignment yet.')

      if (op === 'list') {
        // One of the two lists can fail on older tenants; show whichever loads.
        const [working, submitted] = await Promise.allSettled([
          graphFetch(pick.token, `${base}/${sid}/resources`),
          graphFetch(pick.token, `${base}/${sid}/submittedResources`),
        ])
        if (working.status === 'rejected' && submitted.status === 'rejected') throw working.reason
        return {
          ok: true,
          resources: mapResourceList(working.status === 'fulfilled' ? working.value : null),
          submittedResources: mapResourceList(submitted.status === 'fulfilled' ? submitted.value : null),
        }
      }

      if (op === 'addLink') {
        const url = String(payload.url || '').trim()
        if (!/^https?:\/\//i.test(url)) return { ok: false, error: 'Enter a full http(s) link.' }
        const name = String(payload.name || '').trim().slice(0, 180) || hostOf(url)
        const created = await graphFetch(pick.token, `${base}/${sid}/resources`, {
          method: 'POST',
          body: { resource: { '@odata.type': '#microsoft.graph.educationLinkResource', displayName: name, link: url } },
        })
        return { ok: true, resource: mapResourceList({ value: [created] })[0] || null }
      }

      if (op === 'addFile') return await addFileResource(pick.token, base, sid, payload)

      if (op === 'remove') {
        if (!payload.resourceId) return { ok: false, error: 'Missing resource id.' }
        await graphFetch(pick.token, `${base}/${sid}/resources/${payload.resourceId}`, { method: 'DELETE' })
        return { ok: true }
      }

      return { ok: false, error: 'Unknown submission operation.' }
    } catch (e) {
      hint =
        e.status === 403
          ? 'Your Teams session token lacks EduAssignments.ReadWrite. Open the Assignments tab in Teams Web once, then retry.'
          : e.hint || hint
      attempts.push(attemptFromResult(pick, graphFail(e, meta)))
    }
  }
  const first = attempts[0] || {}
  const baseMsg = first.message && first.message !== first.graphCode ? first.message : ''
  return {
    ok: false,
    error: [baseMsg, hint].filter(Boolean).join(' — ') || 'Couldn’t update your work. Check your Teams session and try again.',
    meta: { attempts },
  }
}

async function clearData() {
  await chrome.storage.session.remove([CANDIDATES_KEY, SEEN_KEY, REFRESH_KEY])
  await chrome.storage.local.remove([CACHE_KEY, SNAPSHOT_KEY, DETAILS_KEY, TRIGGER_KEY, ...LEGACY_LOCAL_KEYS])
  return { ok: true }
}

/* ---------------- message router ---------------- */

chrome.runtime.onMessage.addListener((req, sender, respond) => {
  const type = req?.type
  if (type === MSG.TOKEN) {
    const p = req.payload
    if (p?.token)
      saveToken(p)
        .then(() => {
          if (resolvesRefresh(p)) {
            notifyEduToken()
            rememberTrigger(sender, p).catch(() => {})
          }
          respond?.({ ok: true })
        })
        .catch(() => {})
    else respond?.({ ok: false })
    return true
  }
  if (type === MSG.SNAPSHOT) {
    if (req.payload?.body) {
      chrome.storage.local
        .set({ [SNAPSHOT_KEY]: { url: req.payload.url, takenAt: req.payload.takenAt, size: String(req.payload.body).length } })
        .catch(() => {})
      respond?.({ ok: true })
    } else respond?.({ ok: false })
    return true
  }
  if (type === MSG.OPEN_DASHBOARD) {
    openDashboard()
    respond?.({ ok: true })
    return
  }
  if (type === MSG.OPEN_TEAMS) {
    openTeams()
      .then((res) => respond?.(res))
      .catch(() => respond?.({ ok: false }))
    return true
  }
  if (type === MSG.SYNC) {
    sharedSync(req.payload || {})
      .then((res) => respond?.(res))
      .catch(() => respond?.({ ok: false, code: 'error' }))
    return true
  }
  if (type === MSG.DETAIL) {
    doDetail(req.payload || {})
      .then((res) => respond?.(res))
      .catch(() => respond?.({ ok: false, code: 'error' }))
    return true
  }
  if (type === MSG.SUBMIT) {
    doSubmit(req.payload || {})
      .then((res) => respond?.(res))
      .catch(() => respond?.({ ok: false, error: 'Turn-in failed.' }))
    return true
  }
  if (type === MSG.SUBMISSION) {
    doSubmission(req.payload || {})
      .then((res) => respond?.(res))
      .catch(() => respond?.({ ok: false, error: 'Couldn’t update your work.' }))
    return true
  }
  if (type === MSG.TITLE) {
    fetchTitle(req.payload?.url || '')
      .then((res) => respond?.(res))
      .catch(() => respond?.({ ok: false, title: null }))
    return true
  }
  if (type === MSG.CLEAR) {
    clearData()
      .then((res) => respond?.(res))
      .catch(() => respond?.({ ok: false }))
    return true
  }
  return false
})
