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
  const [mode, setMode] = useState(QS.get('mode') || 'todo') // todo | done | all
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
    const list = data?.classes || []
    return list.map((c) => ({ id: c.id, name: c.name || c.displayName || 'Class', description: c.description || '', webUrl: c.webUrl || null }))
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
    if (mode !== 'all') {
      filtered = filtered.filter((a) => {
        const p = a.state.progress
        const todoish = p === 'todo' || p === 'reassigned'
        return mode === 'todo' ? todoish : !todoish
      })
    }
    return sortAssignments(filtered, now)
  }, [data, mode, classFilter])

  const sections = useMemo(() => {
    if (mode === 'done') return [{ key: 'done', title: 'Completed', rows }]
    const defs = [
      ['overdue', 'Overdue'],
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
    if (mode === 'all') {
      const done = rows.filter((a) => a.state.progress !== 'todo' && a.state.progress !== 'reassigned')
      if (done.length) out.push({ key: 'done', title: 'Completed', rows: done })
    }
    return out
  }, [rows, mode])

  /* ---------- detail ---------- */
  const openDetail = useCallback(
    async (vm) => {
      const cls = classMap.get(vm.classId)
      setDetail({ vm, cls, phase: 'loading', raw: null })
      const res = await source.detail({ classId: vm.classId, id: vm.id })
      setDetail((cur) => (cur && cur.vm.id === vm.id ? { ...cur, phase: res?.ok ? 'ready' : 'error', raw: res?.assignment || null } : cur))
    },
    [classMap, source]
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
                { value: 'todo', label: 'To do' },
                { value: 'done', label: 'Done' },
                { value: 'all', label: 'All' },
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
                      <AssignmentCard key={a.id} a={a} cls={classMap.get(a.classId)} onOpen={openDetail} />
                    ))}
                  </Section>
                ))
              ) : (
                <StateAction
                  icon={mode === 'todo' ? 'celebration' : 'inbox'}
                  headline={mode === 'todo' ? 'All clear!' : 'Nothing here'}
                  body={
                    mode === 'todo'
                      ? 'Nothing left to do — check the Done tab for graded work.'
                      : 'No assignments match this filter right now.'
                  }
                  actions={
                    <button type="button" className="btn btn-tonal btn-md" onClick={() => setMode('all')}>
                      Show all
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
              body={meta?.hint || meta?.message || 'Graph request failed — your Teams session may have expired. Refresh the Teams tab and retry.'}
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
