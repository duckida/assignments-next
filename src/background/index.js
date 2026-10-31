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
const CACHE_KEY = 'mad3.cache.v1'
const SNAPSHOT_KEY = 'mad3.snapshot.v1'
const VISIBLE = new Set(['assigned', 'published', 'inactive', 'unknownfuturevalue'])

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

function pickToken(candidates) {
  const graph = candidates.find((c) => isGraphAud(c.aud))
  if (graph) return { token: graph.token, strategy: 'graph', meta: graph }
  const edu = candidates.find((c) => isEduAud(c.aud))
  if (edu) return { token: edu.token, strategy: 'edu', meta: edu }
  // any token carrying education scopes is usable against Graph
  const scoped = candidates.find((c) => /EduAssignments|EduRoster/i.test(c.scp || ''))
  if (scoped) return { token: scoped.token, strategy: 'graph', meta: scoped }
  return null
}

/* ---------------- graph strategy ---------------- */

async function getClasses(token) {
  try {
    const items = await graphPaged(token, '/education/me/classes?$select=id,displayName,description&$top=100', {})
    if (items.length) return items.map((c) => ({ id: c.id, name: c.displayName || 'Class', description: c.description, webUrl: null }))
  } catch {}
  try {
    const items = await graphPaged(token, '/me/joinedTeams?$select=id,displayName,description,webUrl&$top=200', {})
    return items.map((t) => ({ id: t.id, name: t.displayName || 'Team', description: t.description, webUrl: t.webUrl || null }))
  } catch {}
  return null
}

async function graphSync(pick, meta) {
  const token = pick.token
  const classes = (await getClasses(token)) || []
  let rawAssignments
  try {
    rawAssignments = await graphPaged(token, '/education/me/assignments?$top=200', { version: 'v1.0' })
  } catch (e) {
    return graphFail(e, meta)
  }
  const assignments = rawAssignments.filter((a) => VISIBLE.has(String(a.status || '').toLowerCase()))
  const subPaths = assignments.map((a) => `/education/classes/${a.classId}/assignments/${a.id}/submissions?$expand=outcomes`)
  const subs = await safeBatch(token, subPaths)
  const submissions = {}
  for (const [path, body] of subs) {
    const m = /\/assignments\/([a-z0-9-]+)\/submissions/i.exec(path)
    const val = body?.value
    if (m && Array.isArray(val) && val.length) submissions[m[1]] = val[0]
  }
  return { ok: true, data: { classes, assignments, submissions, ts: Date.now() }, meta: { ...meta, strategy: 'graph' } }
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
    meta: { ...meta, message: e.message, hint: e.hint || '' },
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

/* ---------------- sync / detail / clear ---------------- */

let syncing = false

async function doSync() {
  const { candidates, seen } = await getCandidates()
  const pick = pickToken(candidates)
  if (!pick) {
    return {
      ok: false,
      code: 'no-token',
      meta: {
        seen: seen.map((s) => s.aud),
        hint:
          'Reload the Teams tab (extensions inject into pages only on load), then click around in Teams Web so a token is issued, and hit Refresh here.',
      },
    }
  }
  const meta = { capturedAt: pick.meta.capturedAt, expMs: pick.meta.expMs, scp: pick.meta.scp, upn: pick.meta.upn, seen: seen.map((s) => s.aud) }
  const res = pick.strategy === 'graph' ? await graphSync(pick, meta) : await eduSync(pick, meta)
  return res
}

async function doDetail({ classId, id }) {
  const { candidates, seen } = await getCandidates()
  const pick = pickToken(candidates)
  if (!pick) return { ok: false, code: 'no-token', meta: { seen: seen.map((s) => s.aud) } }
  const meta = { upn: pick.meta.upn, strategy: pick.strategy }
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
      return { ok: true, assignment: a, meta }
    }
    const a = await eduGet(pick.token, `/teams/${classId}/assignments/${id}`)
    const raw = eduMapAssignment(a, classId) || { id, classId }
    raw.instructions = { content: a?.instructions || a?.description || '' }
    raw.resources = Array.isArray(a?.resources)
      ? a.resources.map((r) => ({ id: r.id || r.name, displayName: r.displayName || r.name || 'Resource', resource: { link: r.webUrl || r.link || null } }))
      : []
    return { ok: true, assignment: raw, meta }
  } catch (e) {
    return graphFail(e, meta)
  }
}

async function clearData() {
  await chrome.storage.session.remove([CANDIDATES_KEY, SEEN_KEY])
  await chrome.storage.local.remove([CACHE_KEY, SNAPSHOT_KEY])
  return { ok: true }
}

/* ---------------- message router ---------------- */

chrome.runtime.onMessage.addListener((req, sender, respond) => {
  const type = req?.type
  if (type === MSG.TOKEN) {
    if (req.payload?.token)
      saveToken(req.payload)
        .then(() => respond?.({ ok: true }))
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
  if (type === MSG.SYNC) {
    if (syncing) {
      respond?.({ ok: false, code: 'busy' })
      return
    }
    syncing = true
    doSync()
      .then(async (res) => {
        if (res?.ok) {
          const cache = { ts: res.data.ts, classes: res.data.classes, assignments: res.data.assignments, submissions: res.data.submissions }
          await chrome.storage.local.set({ [CACHE_KEY]: cache }).catch(() => {})
        }
        respond?.(res)
      })
      .catch(() => respond?.({ ok: false, code: 'error' }))
      .finally(() => {
        syncing = false
      })
    return true
  }
  if (type === MSG.DETAIL) {
    doDetail(req.payload || {})
      .then((res) => respond?.(res))
      .catch(() => respond?.({ ok: false, code: 'error' }))
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
