/**
 * Data source adapters with one contract:
 *   { kind, initial(), sync(), detail({classId,id}), clear() }
 *  - extension: talks to the service worker (which owns the captured session token)
 *  - demo:      generates a live, now-relative mock dataset in-page
 * Both return raw-ish { classes, assignments, submissions } that App normalizes
 * through src/shared/normalize.js, so the UI code path is identical.
 */
import { MSG } from '../shared/protocol.js'
import { buildDemoData, demoDetail } from '../shared/demo-data.js'

const CACHE_KEY = 'mad3.cache.v2' // v2: rows carry className (v1 lacked it)
const LEGACY_KEYS = ['mad3.cache.v1']
const SNAP_KEY = 'mad3.snapshot.v1'

function hasRuntime() {
  try {
    return typeof chrome !== 'undefined' && !!chrome.runtime && !!chrome.runtime.id
  } catch {
    return false
  }
}

function send(type, payload) {
  return new Promise((resolve) => {
    try {
      chrome.runtime.sendMessage({ type, payload }, (res) => {
        if (chrome.runtime.lastError) resolve({ ok: false, code: 'error' })
        else resolve(res || { ok: false, code: 'error' })
      })
    } catch {
      resolve({ ok: false, code: 'error' })
    }
  })
}

function wrapData(d) {
  if (!d) return null
  return {
    classes: d.classes || [],
    assignments: d.assignments || [],
    submissions: d.submissions || {},
    ts: d.ts || null,
  }
}

function extSource() {
  return {
    kind: 'extension',
    async initial() {
      try {
        const store = await chrome.storage.local.get([CACHE_KEY, SNAP_KEY])
        const d = store?.[CACHE_KEY]
        if (!d?.assignments?.length) return null
        return { data: wrapData(d), snapshot: store[SNAP_KEY] || null }
      } catch {
        return null
      }
    },
    sync() {
      return send(MSG.SYNC)
    },
    detail({ classId, id }) {
      return send(MSG.DETAIL, { classId, id })
    },
    submit({ classId, id, submissionId, action }) {
      return send(MSG.SUBMIT, { classId, id, submissionId, action })
    },
    title(url) {
      return send(MSG.TITLE, { url })
    },
    async clear() {
      await send(MSG.CLEAR)
      try {
        await chrome.storage.local.remove([CACHE_KEY, SNAP_KEY, ...LEGACY_KEYS])
      } catch {}
      return true
    },
    snapshot() {
      return chrome.storage.local.get(SNAP_KEY).then((s) => s?.[SNAP_KEY] || null).catch(() => null)
    },
  }
}

function demoSource() {
  return {
    kind: 'demo',
    async initial() {
      const d = buildDemoData()
      return { data: wrapData(d), snapshot: null }
    },
    async sync() {
      await new Promise((r) => setTimeout(r, 420))
      return { ok: true, data: wrapData(buildDemoData()), meta: { demo: true, upn: 'student@demo.edu' } }
    },
    async detail({ id }) {
      await new Promise((r) => setTimeout(r, 180))
      const a = demoDetail(id)
      return a ? { ok: true, assignment: a } : { ok: false, code: 'error' }
    },
    async submit() {
      await new Promise((r) => setTimeout(r, 500))
      return { ok: true }
    },
    async clear() {
      return true
    },
    async snapshot() {
      return null
    },
    async title() {
      return { ok: false, title: null }
    },
  }
}

export function createSource(kind) {
  const k = kind || (hasRuntime() ? 'extension' : 'demo')
  return k === 'extension' ? extSource() : demoSource()
}
