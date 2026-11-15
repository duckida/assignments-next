/** Thin Microsoft Graph client for the service worker (read-only education endpoints). */
import { isGraphUrl } from '../shared/resources.js'

export class GraphError extends Error {
  constructor(code, status, message, hint) {
    super(message || code)
    this.name = 'GraphError'
    this.code = code
    this.status = status
    this.hint = hint
  }
}

const GRAPH = 'https://graph.microsoft.com'

/** GET/POST/PUT/DELETE a Graph path; relPath must start with '/', or be an absolute Graph URL (from nextLink).
 *  Pass `rawBody` (Uint8Array) for binary uploads — used by submission file resources.
 */
export async function graphFetch(token, relPath, { version = 'v1.0', method = 'GET', body, rawBody, contentType } = {}) {
  const url = /^https?:/.test(relPath) ? relPath : `${GRAPH}/${version}${relPath}`
  if (!isGraphUrl(url)) throw new GraphError('invalid-url', 400, 'Expected a Microsoft Graph URL.')
  if (typeof token !== 'string' || !token.trim()) throw new GraphError('no-token', 401, 'No Graph session token is available.')
  let res
  try {
    res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': contentType || 'application/json',
      },
      body: rawBody || (body ? JSON.stringify(body) : undefined),
    })
  } catch (e) {
    throw new GraphError('network', 0, 'Network request to Graph failed', String(e))
  }
  let json = null
  const text = await res.text()
  if (text) {
    try {
      json = JSON.parse(text)
    } catch {}
  }
  if (!res.ok) {
    const code = json?.error?.code || String(res.status)
    const message = json?.error?.message || res.statusText
    let hint
    if (res.status === 401) hint = 'Captured token was rejected. Refresh the Teams tab so a fresh token is minted, then press Refresh.'
    if (res.status === 403 && /Authorization_RequestDenied|insufficient/i.test(message + code))
      hint = 'This Teams session token lacks an education-read scope. Open the Assignments tab inside Teams Web once (so it requests the right scope), then Refresh here.'
    throw new GraphError(code, res.status, message, hint)
  }
  return json
}

/** Walk $odata.nextLink pages. */
export async function graphPaged(token, relPath, { version = 'v1.0', maxPages = 10 } = {}) {
  const items = []
  let path = relPath
  for (let i = 0; i < maxPages && path; i++) {
    const page = await graphFetch(token, path, { version })
    if (Array.isArray(page.value)) items.push(...page.value)
    const next = page['@odata.nextLink']
    path = next || null
  }
  return items
}

/** $batch requests (max 20 per batch). subpaths: array of Graph relative URLs.
 *  Chunks are sent with limited concurrency (parallel requests instead of serial
 *  round-trips — the dominant sync cost for large assignment lists).
 */
export async function graphBatch(token, subpaths, { version = 'v1.0', concurrency = 4 } = {}) {
  const results = new Map() // subpath -> json|null
  const chunks = []
  for (let i = 0; i < subpaths.length; i += 20) chunks.push(subpaths.slice(i, i + 20))
  const queue = [...chunks]
  let networkErr = null
  const worker = async () => {
    while (queue.length) {
      const chunk = queue.shift()
      const body = {
        requests: chunk.map((p, j) => ({ id: String(j), method: 'GET', url: p })),
      }
      let json
      try {
        json = await graphFetch(token, '/$batch', { version, method: 'POST', body })
      } catch (e) {
        if (e && e.status === 0) networkErr = networkErr || e
        continue // batch endpoint hiccup: leave these subrequests unresolved
      }
      const responses = Array.isArray(json?.responses) ? json.responses : []
      for (const r of responses) {
        const path = chunk[Number(r.id)]
        if (!path) continue
        const okBody = r.status >= 200 && r.status < 300 ? r.body : null
        results.set(path, okBody)
      }
    }
  }
  await Promise.all(Array.from({ length: Math.min(concurrency, chunks.length) }, worker))
  if (networkErr) throw networkErr
  return results
}
