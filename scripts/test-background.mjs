#!/usr/bin/env node
/**
 * Smoke test for the service worker's background sign-in logic (no browser needed).
 * Bundles src/background/index.js to a temp IIFE, runs it in a vm with a mocked
 * chrome API, and drives the SYNC / TOKEN / OPEN_TEAMS / CLEAR messages.
 *
 *   npm test
 */
import { build } from 'esbuild'
import { mkdtempSync, readFileSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join, resolve } from 'node:path'
import vm from 'node:vm'

const root = resolve(import.meta.dirname, '..')
const tmp = mkdtempSync(join(tmpdir(), 'mad3-test-'))
const outfile = join(tmp, 'background.js')
await build({
  entryPoints: [join(root, 'src/background/index.js')],
  outfile,
  bundle: true,
  format: 'iife',
  target: ['chrome111'],
  logLevel: 'silent',
})
const code = readFileSync(outfile, 'utf8')
process.on('exit', () => rmSync(tmp, { recursive: true, force: true }))

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

vm.runInNewContext(code, { chrome, fetch: (u, i) => currentFetch(u, i), AbortSignal, setTimeout, clearTimeout, console, Date, URL, TextDecoder })

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

// 7. clear wipes the learned trigger
await send(MSG.CLEAR, {})
check('clear removes trigger', !local._m.has('mad3.eduTrigger.v1'), [...local._m.keys()])

console.log(out.join('\n'))
process.exit(out.some((l) => l.startsWith('FAIL')) ? 1 : 0)
