/**
 * capture.js — runs in the MAIN world (manifest v3 content script,
 * "world": "MAIN"), document_start, on teams.microsoft.com / *.cloud.microsoft.
 *
 * Hooks fetch() + XMLHttpRequest to passively capture:
 *  1. Bearer tokens the Teams Web client itself sends, keeping only ones whose
 *     audience is Microsoft Graph
 *  2. Raw education/assignments JSON payloads flowing through (passive snapshot)
 *
 * Nothing the site sends is modified; this only observes. Captured material is
 * handed to the isolated content-script bridge via window.postMessage and then
 * to the extension service worker. Tokens stay on-device.
 */
import { isGraphAccessToken, tokenSummary } from '../shared/jwt.js'
import { WINDOW_MARKER, WT } from '../shared/protocol.js'

(() => {
  if (window.__MAD3_CAPTURE__) return
  window.__MAD3_CAPTURE__ = true

  let lastTokenKey = ''
  const MAX_SNAPSHOT_BODY = 1_500_000
  let sentSinceLoad = 0
  const MAX_MESSAGES_PER_PAGE = 80

  function post(type, payload) {
    if (sentSinceLoad >= MAX_MESSAGES_PER_PAGE) return
    sentSinceLoad++
    try {
      window.postMessage({ [WINDOW_MARKER]: true, type, payload }, window.location.origin)
    } catch {
      /* ignore */
    }
  }

  function considerToken(token) {
    if (!token || typeof token !== 'string' || token.length < 20) return
    if (!isGraphAccessToken(token)) return
    const s = tokenSummary(token)
    if (!s) return
    if (s.expMs && s.expMs < Date.now()) return
    const key = `${s.expMs}|${token.length}`
    if (key === lastTokenKey) return
    lastTokenKey = key
    post(WT.TOKEN, { token, expMs: s.expMs, scp: s.scp, upn: s.upn, capturedAt: Date.now() })
  }

  function createsTokenLeakRisk(url) {
    // Defensively ignore URLs that carry tokens in query strings; we never touch those.
    try {
      const u = new URL(url, window.location.origin)
      for (const k of u.searchParams.keys()) if (/token|access|code|auth/i.test(k)) return true
    } catch {}
    return false
  }

  /** Graph me/assignments-ish shape filter for snapshots */
  function shapedEdu(value, url) {
    const s = String(url || '')
    const isGraphEdu = /graph\.microsoft\./.test(s) && /\/education\/(me\/assignments|users\/[^/]+\/assignments|classes\/[^/]+\/assignments)/i.test(s)
    if (!isGraphEdu) return false
    const arr = Array.isArray(value?.value) ? value.value : Array.isArray(value) ? value : value?.id ? [value] : []
    return arr.some((x) => x && x.classId && (x.dueDateTime !== undefined || String(x['@odata.type'] || '').includes('educationAssignment')))
  }

  function sniffResponse(urlString, contentType /* string|null */, bodyProvider /* () => Promise<string> */) {
    if (!contentType || !contentType.includes('json')) return
    const url = String(urlString || '')
    const isTeamsEdu = /teams\.(microsoft\.com|cloud\.microsoft)/.test(url) && /assignment/i.test(url)
    if (!isTeamsEdu && !/graph\.microsoft\./.test(url)) return
    if (createsTokenLeakRisk(url)) return
    let body
    try {
      body = bodyProvider()
    } catch {
      return
    }
    Promise.resolve(body)
      .then((text) => {
        if (!text || text.length > MAX_SNAPSHOT_BODY) return
        if (!/^\s*[[{]/.test(text)) return
        let value = null
        try {
          value = JSON.parse(text)
        } catch {
          return
        }
        if (isTeamsEdu || shapedEdu(value, url)) {
          post(WT.SNAPSHOT, { url, takenAt: Date.now(), body: text })
        }
      })
      .catch(() => {})
  }

  /* ---------- fetch hook ---------- */
  const origFetch = window.fetch
  if (typeof origFetch === 'function') {
    window.fetch = function (input, init) {
      try {
        let url = ''
        let auth = ''
        if (typeof input === 'string') {
          url = input
        } else if (input && typeof input === 'object' && 'url' in input) {
          url = String(input.url)
          try {
            auth = input.headers?.get?.('authorization') ?? ''
          } catch {}
        }
        if (!auth && init?.headers) auth = headerGet(init.headers, 'authorization') ?? ''
        if (auth && /^bearer\s+/i.test(auth) && !createsTokenLeakRisk(url)) {
          considerToken(auth.replace(/^bearer\s+/i, '').trim())
        }
        const resP = origFetch.apply(this, arguments)
        try {
          resP.then((r) =>
            sniffResponse(url, r?.headers?.get?.('content-type'), () => (r.clone ? r.clone().text() : ''))
          )
        } catch {}
        return resP
      } catch {
        return origFetch.apply(this, arguments)
      }
    }
  }

  function headerGet(headers, name) {
    if (!headers) return undefined
    try {
      if (typeof Headers !== 'undefined' && headers instanceof Headers) return headers.get(name)
      if (Array.isArray(headers)) {
        for (const [k, v] of headers) if (String(k).toLowerCase() === name) return v
        return undefined
      }
      for (const k of Object.keys(headers)) if (k.toLowerCase() === name) return headers[k]
    } catch {}
    return undefined
  }

  /* ---------- XHR hook ---------- */
  const XHROpen = XMLHttpRequest.prototype.open
  const XHRSend = XMLHttpRequest.prototype.send
  const XHRSetHeader = XMLHttpRequest.prototype.setRequestHeader

  XMLHttpRequest.prototype.open = function (method, url, ...rest) {
    this.__mad3Url = String(url ?? '')
    this.__mad3Auth = ''
    return XHROpen.call(this, method, url, ...rest)
  }
  XMLHttpRequest.prototype.setRequestHeader = function (name, value) {
    if (String(name).toLowerCase() === 'authorization' && typeof value === 'string') this.__mad3Auth = value
    return XHRSetHeader.call(this, name, value)
  }
  XMLHttpRequest.prototype.send = function (body) {
    const self = this
    if (this.__mad3Auth && /^bearer\s+/i.test(this.__mad3Auth) && !createsTokenLeakRisk(this.__mad3Url)) {
      considerToken(this.__mad3Auth.replace(/^bearer\s+/i, '').trim())
    }
    this.addEventListener('load', () => {
      try {
        sniffResponse(
          self.__mad3Url,
          typeof self.getResponseHeader === 'function' ? self.getResponseHeader('content-type') : null,
          () => String(self.responseText ?? '')
        )
      } catch {}
    })
    return XHRSend.call(this, body)
  }
})()
