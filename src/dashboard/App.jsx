import { useCallback, useEffect, useMemo, useRef, useState } from 'react'
import Icon from './components/Icon.jsx'
import { ClassAvatar, FilterChip, IconButton, Segment } from './components/Chips.jsx'
import AssignmentCard from './components/AssignmentCard.jsx'
import DetailSheet from './components/DetailSheet.jsx'
import SettingsSheet from './components/SettingsSheet.jsx'
import { Section, SkeletonList, StateAction } from './components/States.jsx'
import { createSource } from './data-source.js'
import { dueBucket, mapAssignment, mapSubmission, progress, sortAssignments } from '../shared/normalize.js'

const THEME_KEY = 'mad3.theme'
const QS = new URLSearchParams(location.search)

export default function App() {
  const source = useMemo(() => createSource(), [])
  const [theme, setTheme] = useState(() => QS.get('theme') || localStorage.getItem(THEME_KEY) || 'system')
  const [phase, setPhase] = useState('loading') // loading | ready | empty | no-token | forbidden | error
  const [data, setData] = useState(null)
  const [meta, setMeta] = useState(null)
  const [syncing, setSyncing] = useState(false)
  const [snack, setSnack] = useState(null)
  const [mode, setMode] = useState(QS.get('mode') || 'upcoming') // overdue | upcoming | done
  const [classFilter, setClassFilter] = useState(null)
  const [detail, setDetail] = useState(null) // {vm, cls, phase, raw, error}
  const [settingsOpen, setSettingsOpen] = useState(false)
  const snackTimer = useRef(0)

  /* ---------- theme ---------- */
  useEffect(() => {
    localStorage.setItem(THEME_KEY, theme)
    const apply = () => {
      const dark = theme === 'dark' || (theme === 'system' && matchMedia('(prefers-color-scheme: dark)').matches)
      document.documentElement.dataset.theme = dark ? 'dark' : 'light'
    }
    apply()
    matchMedia('(prefers-color-scheme: dark)').addEventListener('change', apply)
    return () => matchMedia('(prefers-color-scheme: dark)').removeEventListener('change', apply)
  }, [theme])

  const notify = useCallback((msg) => {
    setSnack(msg)
    clearTimeout(snackTimer.current)
    snackTimer.current = setTimeout(() => setSnack(null), 4200)
  }, [])

  /* ---------- data ---------- */
  const refresh = useCallback(
    async ({ silent } = {}) => {
      setSyncing(true)
      try {
        const res = await source.sync()
        if (res?.ok && res.data) {
          setData(res.data)
          setMeta(res.meta || null)
          setPhase(res.data.assignments.length ? 'ready' : 'empty')
        } else {
          const code = res?.code || 'error'
          setPhase(code === 'no-token' ? 'no-token' : code === 'forbidden' || code === 'expired' ? 'forbidden' : 'error')
          setMeta(res?.meta || null)
          if (!silent) notify(code === 'no-token' ? 'No Teams session token yet — open Teams Web once.' : 'Sync failed.')
        }
      } catch {
        setPhase('error')
        if (!silent) notify('Sync failed.')
      } finally {
        setSyncing(false)
      }
    },
    [source, notify]
  )

  useEffect(() => {
    let alive = true
    ;(async () => {
      const init = await source.initial()
      if (!alive) return
      if (init?.data) {
        setData(init.data)
        setPhase(init.data.assignments.length ? 'ready' : 'empty')
        refresh({ silent: true })
      } else {
        refresh({ silent: false })
      }
    })()
    return () => {
      alive = false
    }
  }, [source, refresh])

  /* ---------- view model ---------- */
  const classes = useMemo(() => {
    const list = (data?.classes || []).map((c) => ({ id: c.id, name: c.name || c.displayName || 'Class', description: c.description || '', webUrl: c.webUrl || null }))
    // Fill gaps from per-assignment class names (roster reads can be empty or use
    // different ids); keeps filter chips and avatars working when data.classes misses.
    const known = new Set(list.map((c) => c.id))
    for (const a of data?.assignments || []) {
      if (a.classId && a.className && !known.has(a.classId)) {
        known.add(a.classId)
        list.push({ id: a.classId, name: a.className, description: '', webUrl: null })
      }
    }
    return list
  }, [data])

  const classMap = useMemo(() => new Map(classes.map((c) => [c.id, c])), [classes])

  const rows = useMemo(() => {
    const now = Date.now()
    const subs = data?.submissions || {}
    const list = []
    for (const raw of data?.assignments || []) {
      const vm = mapAssignment(raw)
      if (!vm) continue
      vm.submission = mapSubmission(subs[raw.id], raw.id) || undefined
      vm.state = { bucket: dueBucket(vm, now), progress: progress(vm) }
      list.push(vm)
    }
    let filtered = list
    if (classFilter) filtered = filtered.filter((a) => a.classId === classFilter)
    filtered = filtered.filter((a) => {
      const p = a.state.progress
      const todoish = p === 'todo' || p === 'reassigned'
      if (mode === 'done') return !todoish
      if (mode === 'overdue') return todoish && a.state.bucket === 'overdue'
      return todoish && a.state.bucket !== 'overdue' // upcoming
    })
    return sortAssignments(filtered, now)
  }, [data, mode, classFilter])

  const sections = useMemo(() => {
    if (mode === 'done') return rows.length ? [{ key: 'done', title: 'Completed', rows }] : []
    if (mode === 'overdue') return rows.length ? [{ key: 'overdue', title: 'Overdue', rows }] : []
    const defs = [
      ['today', 'Due today'],
      ['soon', 'Next few days'],
      ['later', 'Later'],
      ['none', 'No due date'],
    ]
    const out = []
    for (const [key, title] of defs) {
      const r = rows.filter((a) => (key === 'none' ? a.state.bucket === 'none' : a.state.bucket === key))
      if (r.length) out.push({ key, title, rows: r })
    }
    return out
  }, [rows, mode])

  // class lookup with per-assignment name fallback (classes list can be empty
  // when the token lacks EduRoster scopes)
  const clsFor = useCallback((a) => classMap.get(a.classId) || (a.className ? { id: a.classId, name: a.className } : undefined), [classMap])

  /* ---------- detail ---------- */
  const openDetail = useCallback(
    async (vm) => {
      const cls = clsFor(vm)
      setDetail({ vm, cls, phase: 'loading', raw: null })
      const res = await source.detail({ classId: vm.classId, id: vm.id })
      setDetail((cur) => (cur && cur.vm.id === vm.id ? { ...cur, phase: res?.ok ? 'ready' : 'error', raw: res?.assignment || null } : cur))
    },
    [clsFor, source]
  )

  /* ---------- turn in / undo ---------- */
  const turnIn = useCallback(
    async (vm, action) => {
      const res = await source.submit({ classId: vm.classId, id: vm.id, submissionId: vm.submission?.id || null, action })
      if (!res?.ok) return { ok: false, error: res?.error || res?.meta?.message || '' }
      const status = action === 'submit' ? 'submitted' : 'working'
      const submittedDateTime = action === 'submit' ? new Date().toISOString() : null
      // update the open sheet immediately…
      setDetail((cur) => {
        if (!cur || cur.vm.id !== vm.id) return cur
        const nv = { ...cur.vm, submission: { ...(cur.vm.submission || {}), status, submittedDateTime } }
        nv.state = { bucket: dueBucket(nv), progress: progress(nv) }
        return { ...cur, vm: nv }
      })
      // …and the list rows
      setData((cur) => {
        if (!cur) return cur
        const submissions = { ...(cur.submissions || {}) }
        submissions[vm.id] = { ...(submissions[vm.id] || {}), status, submittedDateTime }
        return { ...cur, submissions }
      })
      notify(action === 'submit' ? 'Turned in.' : 'Turn-in undone.')
      if (source.kind === 'extension') refresh({ silent: true })
      return { ok: true }
    },
    [source, notify, refresh]
  )

  const clearData = useCallback(async () => {
    await source.clear()
    setData(null)
    setPhase('loading')
    setSettingsOpen(false)
    notify('Cleared. Re-opening Teams Web will re-capture your session.')
    refresh({ silent: true })
  }, [source, notify, refresh])

  const todoCount = useMemo(() => rows.filter((a) => a.state.progress === 'todo' || a.state.progress === 'reassigned').length, [rows])
  const failureBody = useMemo(() => {
    if (!meta) return 'Captured token(s) failed against every known assignments API. Open Settings for diagnostics.'
    const attempts = Array.isArray(meta.attempts)
      ? meta.attempts.map((a) => `${a.strategy || '?'} ${a.status || a.code || '?'}${a.graphCode ? ` ${a.graphCode}` : ''}`).join(' · ')
      : ''
    return meta.hint || [meta.message, attempts && `Attempts: ${attempts}`].filter(Boolean).join(' — ') || 'Sync failed.'
  }, [meta])

  /* ---------- render ---------- */
  const themeIcon = theme === 'dark' ? 'dark_mode' : theme === 'light' ? 'light_mode' : 'contrast'

  // demo deep-link: ?open=<assignment-id> opens that detail sheet once data is ready
  const openedViaQs = useRef(false)
  useEffect(() => {
    const id = QS.get('open')
    if (!id || openedViaQs.current || !data) return
    const raw = (data.assignments || []).find((a) => a.id === id)
    if (!raw) return
    openedViaQs.current = true
    const vm = mapAssignment(raw)
    vm.submission = mapSubmission((data.submissions || {})[raw.id], raw.id) || undefined
    vm.state = { bucket: dueBucket(vm), progress: progress(vm) }
    openDetail(vm)
  }, [data, openDetail])

  return (
    <div>
      <header className="topbar">
        <div className="topbar-title">
          <div className="t-headline-s" style={{ display: 'flex', alignItems: 'center', gap: 8 }}>
            Assignments
            {source.kind === 'demo' ? <span className="chip sel t-label-s" style={{ height: 22, pointerEvents: 'none' }}>demo</span> : null}
          </div>
          <div className="sub t-label-m">
            {phase === 'ready' || phase === 'empty'
              ? `${classes.length} classes · ${todoCount} to do`
              : phase === 'loading'
                ? 'Loading…'
                : 'Not synced'}
          </div>
        </div>
        <IconButton name={themeIcon} label={`Theme: ${theme}`} onClick={() => setTheme(theme === 'system' ? 'light' : theme === 'light' ? 'dark' : 'system')} />
        <IconButton name="settings" label="Settings" onClick={() => setSettingsOpen(true)} />
        <IconButton name="sync" label="Refresh" spin={syncing} onClick={() => refresh({ silent: false })} />
      </header>

      {phase === 'ready' || phase === 'empty' || phase === 'loading' ? (
        <>
          <div className="filterbar">
            <Segment
              value={mode}
              onChange={setMode}
              options={[
                { value: 'overdue', label: 'Overdue' },
                { value: 'upcoming', label: 'Upcoming' },
                { value: 'done', label: 'Done' },
              ]}
            />
            <span className="spacer" />
            {classes.length > 1 ? (
              <div className="chiprow" style={{ maxWidth: '100%' }}>
                <FilterChip selected={!classFilter} onClick={() => setClassFilter(null)}>
                  All classes
                </FilterChip>
                {classes.map((c) => (
                  <FilterChip
                    key={c.id}
                    selected={classFilter === c.id}
                    onClick={() => setClassFilter(classFilter === c.id ? null : c.id)}
                    leading={<ClassAvatar cls={c} small />}
                  >
                    {c.name}
                  </FilterChip>
                ))}
              </div>
            ) : null}
          </div>

          {phase === 'loading' && !data ? (
            <div className="list" style={{ paddingTop: 8 }}>
              <SkeletonList />
            </div>
          ) : (
            <main className="list">
              {sections.length ? (
                sections.map((s) => (
                  <Section key={s.key} title={s.title} count={s.rows.length}>
                    {s.rows.map((a) => (
                      <AssignmentCard key={a.id} a={a} cls={clsFor(a)} onOpen={openDetail} />
                    ))}
                  </Section>
                ))
              ) : (
                <StateAction
                  icon={mode === 'done' ? 'inbox' : 'celebration'}
                  headline={mode === 'overdue' ? 'Nothing overdue' : mode === 'upcoming' ? 'All clear!' : 'Nothing here yet'}
                  body={
                    mode === 'overdue'
                      ? 'You’re all caught up — check Upcoming for what’s next.'
                      : mode === 'upcoming'
                        ? 'Nothing left to do — check the Done tab for graded work.'
                        : 'Graded and turned-in work will appear here.'
                  }
                  actions={
                    <button type="button" className="btn btn-tonal btn-md" onClick={() => setMode(mode === 'done' ? 'upcoming' : mode === 'overdue' ? 'upcoming' : 'done')}>
                      {mode === 'done' ? 'View upcoming' : mode === 'overdue' ? 'View upcoming' : 'View completed'}
                    </button>
                  }
                />
              )}
            </main>
          )}
        </>
      ) : (
        <main className="list" style={{ paddingTop: 24 }}>
          {phase === 'no-token' ? (
            <StateAction
              icon="lock"
              headline="No session token yet"
              body="Reload your open Teams tab (extensions only inject on page load), then browse to Assignments in Teams Web so a token is issued. Then hit Refresh."
              actions={
                <>
                  <a className="btn btn-filled btn-md" href="https://teams.microsoft.com" target="_blank" rel="noreferrer">
                    <Icon name="open_in_new" size={18} />
                    Open Teams Web
                  </a>
                  <button type="button" className="btn btn-text btn-md" onClick={() => refresh({ silent: false })}>
                    Retry
                  </button>
                </>
              }
            />
          ) : phase === 'forbidden' ? (
            <StateAction
              icon="key"
              headline="Token lacks education scopes"
              body={meta?.hint || 'Open the Assignments tab inside Teams Web once so your session requests education scopes, then Refresh.'}
              actions={
                <>
                  <a className="btn btn-filled btn-md" href="https://teams.microsoft.com" target="_blank" rel="noreferrer">
                    <Icon name="open_in_new" size={18} />
                    Open Teams Web
                  </a>
                  <button type="button" className="btn btn-text btn-md" onClick={() => refresh({ silent: false })}>
                    Retry
                  </button>
                </>
              }
            />
          ) : (
            <StateAction
              icon="cloud_off"
              headline="Sync failed"
              body={failureBody}
              actions={
                <button type="button" className="btn btn-filled btn-md" onClick={() => refresh({ silent: false })}>
                  <Icon name="sync" size={18} />
                  Retry
                </button>
              }
            />
          )}
        </main>
      )}

      <DetailSheet
        open={!!detail}
        vm={detail?.vm}
        cls={detail?.cls}
        detail={detail ? { phase: detail.phase, raw: detail.raw } : null}
        source={source}
        onTurnIn={turnIn}
        onClose={() => setDetail(null)}
      />
      <SettingsSheet
        open={settingsOpen}
        onClose={() => setSettingsOpen(false)}
        theme={theme}
        setTheme={setTheme}
        source={source}
        meta={meta}
        syncedAt={data?.ts}
        onRefresh={() => {
          setSettingsOpen(false)
          refresh({ silent: false })
        }}
        onClear={clearData}
      />
      {snack ? <div className="snack t-body-m">{snack}</div> : null}
    </div>
  )
}
