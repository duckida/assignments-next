/** Thin Microsoft Graph client for the service worker (read-only education endpoints). */

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

/** GET/POST a Graph path; relPath must start with '/', or be an absolute Graph URL (from nextLink). */
export async function graphFetch(token, relPath, { version = 'v1.0', method = 'GET', body } = {}) {
  const url = /^https?:/.test(relPath) ? relPath : `${GRAPH}/${version}${relPath}`
  let res
  try {
    res = await fetch(url, {
      method,
      headers: {
        Authorization: `Bearer ${token}`,
        'Content-Type': 'application/json',
      },
      body: body ? JSON.stringify(body) : undefined,
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

/** $batch requests (max 20 per batch). subpaths: array of Graph relative URLs. */
export async function graphBatch(token, subpaths, { version = 'v1.0' } = {}) {
  const results = new Map() // subpath -> json|null
  for (let i = 0; i < subpaths.length; i += 20) {
    const chunk = subpaths.slice(i, i + 20)
    const body = {
      requests: chunk.map((p, j) => ({ id: String(j), method: 'GET', url: p })),
    }
    let json
    try {
      json = await graphFetch(token, '/$batch', { version, method: 'POST', body })
    } catch (e) {
      if (e && e.status === 0) throw e
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
  return results
}
