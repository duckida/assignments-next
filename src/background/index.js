/**
 * Service worker — owns the captured session token (storage.session, memory-only)
 * and performs read-only Graph calls on behalf of the dashboard.
 * Token never leaves this device; it is used exactly as the Teams Web client uses
 * it, for the same read-only endpoints the user's session can already access.
 */
import { MSG } from '../shared/protocol.js'
import { GraphError, graphBatch, graphFetch, graphPaged } from './graph.js'

const TOKEN_KEY = 'mad3.session.token.v1'
const CACHE_KEY = 'mad3.cache.v1'
const SNAPSHOT_KEY = 'mad3.snapshot.v1'
const VISIBLE = new Set(['assigned', 'published', 'inactive', 'unknownFutureValue'])

chrome.runtime.onInstalled.addListener(() => {
  // storage.session defaults to TRUSTED_CONTEXTS already; this just pins it.
  // (property is `accessLevel` — throws synchronously when malformed, so wrap in try)
  try {
    chrome.storage?.session?.setAccessLevel?.({ accessLevel: 'TRUSTED_CONTEXTS' })?.catch?.(() => {})
  } catch {
    /* not supported */
  }
})

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

/* ---------------- session token ---------------- */

async function saveToken(p) {
  const prev = await getSessionToken()
  // prefer newest expiry
  if (!prev || !prev.expMs || (p.expMs && p.expMs >= prev.expMs)) {
    await chrome.storage.session.set({ [TOKEN_KEY]: { token: p.token, expMs: p.expMs, scp: p.scp, upn: p.upn, capturedAt: p.capturedAt || Date.now() } })
  }
}

async function getSessionToken() {
  const store = await chrome.storage.session.get(TOKEN_KEY)
  const t = store?.[TOKEN_KEY]
  if (!t?.token) return null
  if (t.expMs && t.expMs < Date.now()) return null // stale
  return t
}

/* ---------------- data assembly ---------------- */

const CLASS_PATH = '/education/classes'

async function getClasses(token) {
  // graph education classes include code (description) — preferred
  try {
    const items = await graphPaged(token, '/education/me/classes?$select=id,displayName,description&$top=100', {})
    if (items.length) return items.map((c) => ({ id: c.id, name: c.displayName || 'Class', description: c.description, webUrl: null }))
  } catch {}
  // Teams joinedTeams fallback: class team id == education class id; webUrl available
  try {
    const items = await graphPaged(token, '/me/joinedTeams?$select=id,displayName,description,webUrl&$top=200', {})
    return items.map((t) => ({ id: t.id, name: t.displayName || 'Team', description: t.description, webUrl: t.webUrl || null }))
  } catch {}
  // third fallback: me/assignments alone (classes unknown)
  return null
}

let syncing = 0

async function doSync() {
  const sess = await getSessionToken()
  if (!sess) return { ok: false, code: 'no-token', meta: { hint: openTeamsHint() } }
  const token = sess.token
  const meta = { capturedAt: sess.capturedAt, expMs: sess.expMs, scp: sess.scp, upn: sess.upn }

  const classes = (await getClasses(token)) || []
  const classMap = new Map(classes.map((c) => [c.id, c]))

  let rawAssignments
  try {
    rawAssignments = await graphPaged(token, '/education/me/assignments?$top=200', { version: 'v1.0' })
  } catch (e) {
    return graphFail(e, meta)
  }
  const assignments = rawAssignments.filter((a) => VISIBLE.has(String(a.status || '').toLowerCase()))

  // single batch of per-assignment "my submission" reads (students only see their own)
  const subPaths = assignments.map((a) => `${CLASS_PATH}/${a.classId}/assignments/${a.id}/submissions?$expand=outcomes`)
  const subs = await safeBatch(token, subPaths)
  const submissions = {}
  for (const [path, body] of subs) {
    const m = /\/assignments\/([a-z0-9-]+)\/submissions/i.exec(path)
    const val = body?.value
    if (m && Array.isArray(val) && val.length) submissions[m[1]] = val[0]
  }

  return {
    ok: true,
    data: { classes, assignments, submissions, ts: Date.now() },
    meta,
  }
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

function openTeamsHint() {
  return 'Sign into teams.microsoft.com (or teams.cloud.microsoft), browse once so a Graph token is issued, then press Refresh here.'
}

async function doDetail({ classId, id }) {
  const sess = await getSessionToken()
  if (!sess) return { ok: false, code: 'no-token', meta: { hint: openTeamsHint() } }
  const token = sess.token
  const base = `${CLASS_PATH}/${classId}/assignments/${id}`
  try {
    let a
    try {
      a = await graphFetch(token, `${base}?$expand=resources`)
    } catch (e) {
      if (e.status === 400) a = await graphFetch(token, base)
      else throw e
    }
    // include instructions content when present
    return { ok: true, assignment: a }
  } catch (e) {
    return graphFail(e, { upn: sess.upn })
  }
}

async function clearData() {
  await chrome.storage.session.remove(TOKEN_KEY)
  await chrome.storage.local.remove([CACHE_KEY, SNAPSHOT_KEY])
  return { ok: true }
}

/* ---------------- message router ---------------- */

chrome.runtime.onMessage.addListener((req, sender, respond) => {
  const type = req?.type
  if (type === MSG.TOKEN) {
    if (req.payload?.token) saveToken(req.payload).then(() => respond?.({ ok: true })).catch(() => {})
    else respond?.({ ok: false })
    return true
  }
  if (type === MSG.SNAPSHOT) {
    // passive record of edu-shaped traffic — kept small, for the settings view only
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
    if (syncing === 1) {
      respond?.({ ok: false, code: 'busy' })
      return
    }
    syncing = 1
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
        syncing = 0
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
    clearData().then((res) => respond?.(res)).catch(() => respond?.({ ok: false }))
    return true
  }
  return false
})
