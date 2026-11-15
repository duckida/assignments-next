#!/usr/bin/env node
/**
 * Service-worker regressions: background sign-in, scoped roster lookup, resources
 * and authenticated inline images. Uses a mocked Chrome API and network.
 *
 *   npm test
 */
import { build } from 'esbuild'
import { parseHTML } from 'linkedom'
import { join, resolve } from 'node:path'
import vm from 'node:vm'
import { describeResource, driveItemUrl, isGraphUrl } from '../src/shared/resources.js'
import { mapAssignment, mapSubmission, sanitizeHtml, sortCompletedAssignments } from '../src/shared/normalize.js'
import { mergeAssignmentClasses } from '../src/shared/classes.js'

const root = resolve(import.meta.dirname, '..')
const bundled = await build({
  entryPoints: [join(root, 'src/background/index.js')],
  write: false,
  bundle: true,
  format: 'iife',
  target: ['chrome111'],
  logLevel: 'silent',
})
const code = bundled.outputFiles[0].text

const mkStore = () => {
  const m = new Map()
  const pick = (keys) => {
    const out = {}
    if (keys == null) {
      for (const [k, v] of m) out[k] = v
      return out
    }
    for (const k of Array.isArray(keys) ? keys : [keys]) if (m.has(k)) out[k] = m.get(k)
    return out
  }
  return {
    _m: m,
    get: async (keys) => pick(keys),
    set: async (obj) => {
      for (const [k, v] of Object.entries(obj)) m.set(k, v)
    },
    remove: async (keys) => {
      for (const k of Array.isArray(keys) ? keys : [keys]) m.delete(k)
    },
  }
}

const session = mkStore()
const local = mkStore()

const calls = { create: [], update: [], remove: [], query: [] }
let nextId = 100
const tabs = new Map()
let msgHandler = null
const sender = { tab: { id: 7, url: 'https://teams.microsoft.com/v2/' } }

const chrome = {
  runtime: {
    id: 'test',
    onInstalled: { addListener() {} },
    onStartup: { addListener() {} },
    onMessage: { addListener: (fn) => (msgHandler = fn) },
    getURL: (p) => 'chrome-extension://test/' + p,
    sendMessage: async () => ({ ok: true }),
  },
  storage: {
    session: { ...session, setAccessLevel: async () => {} },
    local,
  },
  action: { onClicked: { addListener() {} } },
  scripting: { executeScript: async () => {} },
  windows: { update: async () => {} },
  tabs: {
    create: async ({ url, active }) => {
      const tab = { id: nextId++, url, windowId: 1, active: !!active }
      tabs.set(tab.id, tab)
      calls.create.push({ url, active: !!active, id: tab.id })
      return tab
    },
    update: async (id, props) => {
      calls.update.push({ id, props })
      const tab = tabs.get(id) || { id, windowId: 1 }
      Object.assign(tab, props)
      return tab
    },
    remove: async (id) => {
      calls.remove.push(id)
      tabs.delete(id)
    },
    query: async (q) => {
      calls.query.push(q)
      return [...tabs.values()]
    },
    get: async (id) => tabs.get(id) || { id, url: sender.tab.url, windowId: 1 },
  },
}

const ASSIGNMENT = {
  id: 'a1',
  classId: 'c1',
  status: 'assigned',
  displayName: 'Lab report',
  dueDateTime: new Date(Date.now() + 86400000).toISOString(),
  webUrl: 'https://teams.microsoft.com/l/entity/app/assign?a=1',
}

let currentFetch = async (url, init) => {
  const u = String(url)
  let body = { value: [] }
  if (u.includes('/education/me/classes')) body = { value: [{ id: 'c1', displayName: 'Biology' }] }
  else if (u.includes('/education/me/assignments')) body = { value: [ASSIGNMENT] }
  else if (u.includes('/$batch')) {
    const reqs = JSON.parse(init.body).requests || []
    body = {
      responses: reqs.map((r) => ({
        id: r.id,
        status: 200,
        body: r.url.includes('/submissions') ? { value: [{ id: 's1', status: 'working' }] } : { id: 'c1', displayName: 'Biology' },
      })),
    }
  }
  return { ok: true, status: 200, statusText: 'OK', text: async () => JSON.stringify(body) }
}

vm.runInNewContext(code, { chrome, fetch: (u, i) => currentFetch(u, i), AbortSignal, setTimeout, clearTimeout, console, Date, URL, TextDecoder, atob, btoa })

const send = (type, payload) =>
  new Promise((resolve) => {
    const keep = msgHandler({ type, payload }, sender, resolve)
    if (keep !== true) resolve({ ok: false, code: 'no-handler' })
  })

const TOKEN_EDU = 'x'.repeat(40)
const MSG = {
  SYNC: 'mad3/sync',
  TOKEN: 'mad3/token',
  OPEN_TEAMS: 'mad3/open-teams',
  CLEAR: 'mad3/clear',
  RESOURCE: 'mad3/resource',
  DETAIL: 'mad3/detail',
  SUBMISSION: 'mad3/submission',
}

const out = []
const check = (name, cond, extra) => {
  out.push(`${cond ? 'PASS' : 'FAIL'}  ${name}${cond ? '' : '  ' + JSON.stringify(extra)}`)
}

// 1. cold start: no token, nothing learned -> visible Teams tab, no hidden wait
let res = await send(MSG.SYNC, { auto: true })
check('cold start returns no-token', res.ok === false && res.code === 'no-token', res)
check('cold start tags autoRefresh=no-trigger', res.meta?.autoRefresh?.code === 'no-trigger', res.meta?.autoRefresh)
check('cold start opens Teams visibly', calls.create.at(-1)?.active === true, calls.create.at(-1))

// 2. capture an edu-scoped token -> trigger learned from frame URL
const deepLink = 'https://teams.microsoft.com/l/entity/app/assignments?classId=c1'
await send(MSG.TOKEN, { token: TOKEN_EDU, aud: '00000003-0000-0000-c000-000000000000', scp: 'EduAssignments.ReadBasic EduRoster.ReadBasic', expMs: Date.now() + 3600e3, pageUrl: deepLink })
await new Promise((r) => setTimeout(r, 30))
const trigger = local._m.get('mad3.eduTrigger.v1')
check('trigger learned', trigger?.candidates?.[0]?.url === deepLink, trigger)

// 3. token was stored -> a plain sync is already ok and spawns no tab
const before = calls.create.length
res = await send(MSG.SYNC, { auto: true })
check('signed-in sync works', res.ok === true && res.data?.assignments?.length === 1, res.ok ? res.data?.assignments?.length : res)
check('no tab spawned when a token exists', calls.create.length === before, calls.create.length - before)

// 4. token gone (expired session) -> hidden tab at the learned link, token arrives, tab closed
session._m.delete('mad3.session.tokens.v2')
const createsBefore = calls.create.length
const syncP = send(MSG.SYNC, { auto: true })
await new Promise((r) => setTimeout(r, 50))
const spawned = calls.create.at(-1)
check('hidden tab created at learned deep link', calls.create.length === createsBefore + 1 && spawned?.url === deepLink, spawned)
check('hidden tab is not focused', spawned?.active === false, spawned)
await send(MSG.TOKEN, { token: 'y'.repeat(40), aud: '00000003-0000-0000-c000-000000000000', scp: 'EduAssignments.ReadWrite EduRoster.ReadBasic', expMs: Date.now() + 3600e3, pageUrl: deepLink })
res = await syncP
check('sync succeeds after background capture', res.ok === true, res)
check('hidden tab closed after capture', calls.remove.includes(spawned.id), calls.remove)

// 5. OPEN_TEAMS focuses an existing Teams tab instead of making a new one
const creates2 = calls.create.length
const open = await send(MSG.OPEN_TEAMS, {})
check('open-teams reuses a Teams tab', open.ok === true && calls.create.length === creates2, { open, created: calls.create.length - creates2 })

// 6. token present but rejected with 403 -> one forced refresh, then retry wins
{
  const base = currentFetch
  currentFetch = async (url, init) => {
    const auth = String(init?.headers?.Authorization || '')
    if (auth === 'Bearer ' + 'b'.repeat(40)) {
      return {
        ok: false,
        status: 403,
        statusText: 'Forbidden',
        text: async () => JSON.stringify({ error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges (test)' } }),
      }
    }
    return base(url, init)
  }
  session._m.set('mad3.session.tokens.v2', [
    { token: 'b'.repeat(40), aud: '00000003-0000-0000-c000-000000000000', scp: 'EduAssignments.ReadBasic', host: 'teams.microsoft.com', expMs: Date.now() + 1000e3, capturedAt: Date.now() },
  ])
  const creates3 = calls.create.length
  const syncP = send(MSG.SYNC, { auto: true })
  await new Promise((r) => setTimeout(r, 60))
  const spawned = calls.create.at(-1)
  check('403 triggers one hidden refresh tab', calls.create.length === creates3 + 1 && spawned?.active === false, { spawned, creates3, now: calls.create.length })
  await send(MSG.TOKEN, { token: 'g'.repeat(40), aud: '00000003-0000-0000-c000-000000000000', scp: 'EduAssignments.ReadWrite', expMs: Date.now() + 3600e3, pageUrl: deepLink })
  const res403 = await syncP
  const usedFresh = res403.ok === true && (res403.meta?.expMs || 0) > Date.now() + 3000e3
  check('403 retry succeeds with the fresh token', usedFresh, res403.ok ? res403.meta : res403)
  currentFetch = base
}

// Tokens with identical audience, expiry, and length can carry different scopes.
// Keep them all, including when captures arrive concurrently.
const graphAud = '00000003-0000-0000-c000-000000000000'
const expires = Date.now() + 3600e3
const assignmentToken = { token: 'a'.repeat(40), aud: graphAud, expMs: expires, scp: 'EduAssignments.ReadWrite', tid: 'school', oid: 'student', upn: 'student@school.edu' }
const rosterToken = { ...assignmentToken, token: 'r'.repeat(40), scp: 'Team.ReadBasic.All Group.Read.All' }
const filesToken = { ...assignmentToken, token: 'f'.repeat(40), scp: 'Files.Read.All' }
await send(MSG.CLEAR, {})
await Promise.all([assignmentToken, rosterToken, filesToken].map((p) => send(MSG.TOKEN, p)))
const keptTokens = session._m.get('mad3.session.tokens.v2') || []
check('concurrent captures retain distinct same-audience tokens', keptTokens.length === 3, keptTokens.map((p) => p.scp))

const json = (body, status = 200) => new Response(JSON.stringify(body), { status, headers: { 'content-type': 'application/json' } })
const denied = () => json({ error: { code: 'Authorization_RequestDenied', message: 'Insufficient privileges' } }, 403)
const apiCalls = []
const classId = 'a11a1111-1111-1111-1111-111111111111'
currentFetch = async (value, init = {}) => {
  const u = new URL(value)
  const auth = init.headers?.Authorization
  apiCalls.push({ url: u.href, auth })
  if (u.pathname.endsWith('/education/me/assignments')) {
    return auth === `Bearer ${assignmentToken.token}` ? json({ value: [{ ...ASSIGNMENT, classId, className: 'Class' }] }) : denied()
  }
  if (u.pathname.endsWith('/me/joinedTeams')) {
    if (u.search) return json({ error: { message: 'Query options are not supported.' } }, 400)
    return auth === `Bearer ${rosterToken.token}` ? json({ value: [{ id: classId.toUpperCase(), displayName: 'Biology 12' }] }) : denied()
  }
  if (u.pathname.endsWith('/education/me/classes')) return denied()
  if (u.pathname.endsWith('/$batch')) {
    const requests = JSON.parse(init.body).requests
    return json({ responses: requests.map((r) => ({ id: r.id, status: 200, body: { value: [{ id: 's1', status: 'working' }] } })) })
  }
  return denied()
}
res = await send(MSG.SYNC, { auto: false })
check('class names use a separate roster token and case-insensitive IDs', res.ok && res.data.assignments[0].className === 'Biology 12', res)
check('joinedTeams is requested without unsupported query options', apiCalls.some((r) => r.url.endsWith('/me/joinedTeams')) && !apiCalls.some((r) => r.url.includes('/me/joinedTeams?')), apiCalls)
check('placeholder class names do not mask assignment names', mergeAssignmentClasses([{ id: classId.toUpperCase(), name: 'Class' }], [{ classId, className: 'Chemistry' }])[0].name === 'Chemistry')

// If list membership fails, resolve the class via its group ID.
local._m.delete('mad3.cache.v2')
const rosterFetch = currentFetch
currentFetch = async (value, init = {}) => {
  if (String(value).endsWith('/me/joinedTeams')) return denied()
  if (String(value).endsWith('/$batch') && init.headers?.Authorization === `Bearer ${rosterToken.token}`) {
    return json({ responses: JSON.parse(init.body).requests.map((r) => ({
      id: r.id, status: r.url.startsWith('/groups/') ? 200 : 403,
      body: r.url.startsWith('/groups/') ? { displayName: 'Chemistry Lab' } : { error: {} },
    })) })
  }
  return rosterFetch(value, init)
}
res = await send(MSG.SYNC, { auto: false })
check('class-name backfill uses group metadata when roster is unavailable', res.ok && res.data.assignments[0].className === 'Chemistry Lab', res)

// File links may be supplied in link/webUrl/fileUrl and use any drive addressing.
const fileApi = 'https://graph.microsoft.com/v1.0/sites/school/drive/items/file/content?download=1'
const webUrl = 'https://school.sharepoint.com/sites/Biology/Shared%20Documents/diagram.png'
const downloadUrl = 'https://school.sharepoint.com/download/diagram?ticket=short-lived'
const imageBytes = Uint8Array.of(137, 80, 78, 71, 13, 10, 26, 10)
const resourceCalls = []
currentFetch = async (value, init = {}) => {
  const u = new URL(value)
  const auth = init.headers?.Authorization
  resourceCalls.push({ url: u.href, auth })
  if (u.origin === 'https://graph.microsoft.com') {
    if (auth !== `Bearer ${filesToken.token}`) return denied()
    if (u.pathname !== '/v1.0/sites/school/drive/items/file') return json({ error: { message: 'Expected item metadata' } }, 400)
    return json({ id: 'file', name: 'diagram.png', webUrl, file: { mimeType: 'image/png' }, '@microsoft.graph.downloadUrl': downloadUrl })
  }
  if (u.href === downloadUrl) return new Response(imageBytes, { headers: { 'content-type': 'image/png' } })
  return json({ error: {} }, 404)
}
const beforeOpen = calls.create.length
res = await send(MSG.RESOURCE, { op: 'open', url: fileApi })
check('file links resolve with a Files-scoped bearer token', res.ok && resourceCalls.some((r) => r.auth === `Bearer ${filesToken.token}`), { res, resourceCalls })
check('resource click opens SharePoint instead of a bare Graph URL', calls.create.length === beforeOpen + 1 && calls.create.at(-1).url === webUrl, calls.create.at(-1))
res = await send(MSG.RESOURCE, { op: 'preview', url: fileApi })
check('image attachments are returned as inline image bytes', res.ok && res.isImage && res.dataUrl === `data:image/png;base64,${Buffer.from(imageBytes).toString('base64')}`, res)
check('download URLs are fetched without forwarding Graph bearer tokens', resourceCalls.filter((r) => r.url === downloadUrl).every((r) => !r.auth), resourceCalls)
check('image data and download tickets are not persisted', !JSON.stringify([...local._m.values()]).includes('short-lived') && !JSON.stringify([...local._m.values()]).includes('data:image/png'))

const normalized = describeResource({ resource: { displayName: 'diagram.png', link: fileApi, '@odata.type': '#microsoft.graph.educationLinkResource' } })
check('Graph URLs are recognized even when supplied as link resources', normalized.name === 'diagram.png' && normalized.isImage && isGraphUrl(normalized.url), normalized)
const protectedImage = describeResource({
  resource: { displayName: 'diagram.png', webUrl, fileUrl: fileApi, '@odata.type': '#microsoft.graph.educationFileResource' },
})
check('inline images prefer the authenticated Graph file URL over a SharePoint link', protectedImage.url === webUrl && protectedImage.imageUrl === fileApi && protectedImage.apiUrl === fileApi, protectedImage)
const unnamedFile = describeResource({ resource: { webUrl, fileUrl: fileApi, '@odata.type': '#microsoft.graph.educationFileResource' } })
check('unnamed file resources keep their Graph URL for a metadata probe', unnamedFile.url === webUrl && unnamedFile.apiUrl === fileApi && !unnamedFile.hasName, unnamedFile)
const driveForms = [
  '/v1.0/drives/d/items/i/content?download=1',
  '/v1.0/groups/g/drive/items/i',
  '/beta/me/drive/items/i/content',
  '/v1.0/shares/u!abc/driveItem/content',
  '/v1.0/me/drive/root:/folder/image.png:/content',
]
check('all supported Graph drive URL forms resolve to item metadata', driveForms.every((u) => {
  const url = driveItemUrl(`https://graph.microsoft.com${u}`)
  return !!url && !url.endsWith('/content') && !url.includes('?')
}), driveForms)

const beforeInvalid = calls.create.length
res = await send(MSG.RESOURCE, { op: 'open', url: 'https://graph.microsoft.com/v1.0/not-a-file' })
check('unsupported Graph URLs report an error instead of opening without auth', !res.ok && calls.create.length === beforeInvalid, res)

// Protected instruction image uses the original image service's token, not Graph.
const imageOrigin = 'https://image-service.teams.microsoft.com'
await send(MSG.TOKEN, { token: 's'.repeat(40), aud: 'teams-image-audience', scp: '', expMs: expires, tid: 'school', oid: 'student', requestOrigin: imageOrigin })
let imageAuth
currentFetch = async (value, init) => {
  imageAuth = init.headers?.Authorization
  return new Response(imageBytes, { headers: { 'content-type': 'image/png' } })
}
res = await send(MSG.RESOURCE, { op: 'image', url: `${imageOrigin}/pictures/diagram` })
check('instruction images reuse the token observed on that image origin', res.ok && res.dataUrl?.startsWith('data:image/png') && imageAuth === `Bearer ${'s'.repeat(40)}`, res)
res = await send(MSG.RESOURCE, { op: 'image', url: 'https://unrelated.example/image.png' })
check('unrelated image hosts receive no captured bearer token', res.ok && !imageAuth, imageAuth)
currentFetch = async () => new Response('<html>Sign in</html>', { headers: { 'content-type': 'text/html' } })
res = await send(MSG.RESOURCE, { op: 'image', url: 'https://school.sharepoint.com/image.png' })
check('login HTML is not rendered as an inline image', !res.ok && !res.dataUrl, res)

// Rendering an internal EDU detail must retain file/media types and fileUrl.
await send(MSG.CLEAR, {})
await send(MSG.TOKEN, { token: 'e'.repeat(40), aud: 'https://edu.microsoft.com', expMs: expires })
currentFetch = async () => json({ id: 'a1', instructions: { content: '<p>Diagram</p>' }, resources: [{ id: 'r1', fileUrl: fileApi, displayName: 'diagram.png', '@odata.type': '#microsoft.graph.educationMediaResource' }] })
res = await send(MSG.DETAIL, { id: 'a1', classId: 'c1' })
check('internal EDU detail preserves instructions and authenticated file references', res.ok && res.assignment.instructions.content === '<p>Diagram</p>' && res.assignment.resources[0].resource.fileUrl === fileApi, res)

// Done follows completion activity, not oldest due date or assignment edits.
const done = [
  { id: 'old', due: '2026-11-14', submitted: '2026-11-01', modified: '2026-11-15' },
  { id: 'recent', due: '2026-10-01', submitted: '2026-11-13' },
  { id: 'graded', due: '2026-09-01', submitted: '2026-10-01', returned: '2026-11-14' },
  { id: 'fallback', due: '2026-11-02', submitted: 'not-a-date' },
].map((a) => ({
  ...mapAssignment({ id: a.id, classId: 'c1', displayName: a.id, dueDateTime: a.due, lastModifiedDateTime: a.modified }),
  submission: mapSubmission({ submittedDateTime: a.submitted, returnedDateTime: a.returned, status: a.returned ? 'returned' : 'submitted' }),
}))
check('Done is newest completed/returned first with date fallbacks', sortCompletedAssignments(done).map((a) => a.id).join(',') === 'graded,recent,fallback,old')

// Instruction HTML is rendered through a DOMParser sanitizer in the dashboard.
// linkedom supplies that DOM in this Node harness (the browser supplies the real one).
globalThis.DOMParser = class {
  parseFromString(html) {
    return parseHTML(`<!doctype html><html><body>${String(html || '')}</body></html>`).document
  }
}
const sanitized = sanitizeHtml(
  '<p>Read <a href="https://example.com/spec" onclick="steal()">the spec</a> ' +
  '<a href="javascript:alert(1)">or not</a>.</p>' +
  '<script>alert(1)</script>' +
  '<figure><img src="https://school.sharepoint.com/diagram.png" onerror="steal()"><figcaption>Diagram</figcaption></figure>' +
  '<img src="data:image/svg+xml,%3Csvg%3E%3C/svg%3E" alt="inline">' +
  '<img src="file:///etc/passwd">'
)
check('sanitizer keeps images and figures but drops scripts and event handlers', /<figure>/.test(sanitized) && /<figcaption>Diagram<\/figcaption>/.test(sanitized) && !/<script|onerror|onclick|javascript:/i.test(sanitized), sanitized)
check('sanitizer keeps remote and inline images and drops local-file sources', sanitized.includes('https://school.sharepoint.com/diagram.png') && sanitized.includes('data:image/svg+xml') && !sanitized.includes('file:///etc/passwd'), sanitized)
check('sanitizer makes outbound links safe and opaque', /href="https:\/\/example\.com\/spec"/.test(sanitized) && /rel="noopener noreferrer"/.test(sanitized) && !/href="javascript:/i.test(sanitized), sanitized)

// Clear wipes the learned trigger and detail cache, including the older format.
local._m.set('mad3.details.v2', { stale: true })
await send(MSG.CLEAR, {})
check('clear removes trigger', !local._m.has('mad3.eduTrigger.v1'), [...local._m.keys()])
check('clear removes old and current detail caches', !local._m.has('mad3.details.v2') && !local._m.has('mad3.details.v3'), [...local._m.keys()])

console.log(out.join('\n'))
process.exit(out.some((l) => l.startsWith('FAIL')) ? 1 : 0)
