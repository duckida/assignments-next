import { useEffect, useMemo, useState } from 'react'
import Icon from './Icon.jsx'
import { sanitizeHtml, dueLabel } from '../../shared/normalize.js'

const fmtTime = (ts) => (ts ? new Date(ts).toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }) : '—')

function resourceLink(r) {
  const res = r?.resource || {}
  // educationLinkResource → link, educationFileResource/Media → fileUrl/webUrl
  return res.link || res.webUrl || res.fileUrl || null
}

function resourceIcon(r) {
  const t = String(r?.resource?.['@odata.type'] || '')
  if (t.includes('Link')) return 'link'
  if (t.includes('Word')) return 'description'
  if (t.includes('Excel')) return 'table_chart'
  if (t.includes('PowerPoint')) return 'co_present'
  if (t.includes('Pdf')) return 'picture_as_pdf'
  if (t.includes('Media')) return 'play_circle'
  if (t.includes('File')) return 'draft'
  return 'attach_file'
}

const hostnameOf = (url) => {
  try {
    return new URL(url).hostname.replace(/^www\./, '')
  } catch {
    return ''
  }
}

/** Modal popup (desktop): full assignment detail — instructions, resources, grade + feedback. */
export default function DetailSheet({ open, vm, cls, detail, source, onTurnIn, onClose }) {
  useEffect(() => {
    if (!open) return
    const onKey = (e) => e.key === 'Escape' && onClose?.()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [open, onClose])

  const d = detail?.raw || null
  const instructionsHtml = useMemo(() => sanitizeHtml(d?.instructions?.content || ''), [d])
  const resources = useMemo(() => (Array.isArray(d?.resources) ? d.resources : []), [d])
  const webUrl = d?.webUrl || vm?.webUrl || cls?.webUrl || 'https://teams.microsoft.com'

  // turn-in state: two-step confirm (turning in is visible to the teacher)
  const [confirming, setConfirming] = useState(false)
  const [turnBusy, setTurnBusy] = useState(false)
  const [turnErr, setTurnErr] = useState('')
  useEffect(() => {
    setConfirming(false)
    setTurnBusy(false)
    setTurnErr('')
  }, [vm?.id, open])

  const turnIn = async (action) => {
    if (turnBusy) return
    if (action === 'submit' && !confirming) {
      setConfirming(true)
      return
    }
    setTurnBusy(true)
    setTurnErr('')
    const res = await onTurnIn?.(vm, action)
    setTurnBusy(false)
    if (res?.ok) setConfirming(false)
    else setTurnErr(res?.error || 'Turn-in failed — check your Teams session and retry.')
  }

  // resolve <title> of resource links (server-fetched in the SW, cached)
  const [titles, setTitles] = useState({})
  useEffect(() => {
    let alive = true
    setTitles({})
    for (const r of resources) {
      const href = resourceLink(r)
      if (!href || !source?.title) continue
      source.title(href).then((res) => {
        if (alive && res?.ok && res.title) setTitles((cur) => (cur[href] ? cur : { ...cur, [href]: res.title }))
      })
    }
    return () => {
      alive = false
    }
  }, [resources, source])

  if (!vm) return null

  const sub = vm.submission
  const isOverdue = vm.state.bucket === 'overdue'
  const isGraded = vm.state.progress === 'graded'
  const isSubmitted = vm.state.progress === 'submitted'
  const isReassigned = vm.state.progress === 'reassigned'
  const isExcused = vm.state.progress === 'excused'
  const isTodo = vm.state.progress === 'todo'
  // the status banner already shows due info except for these states
  const bannerShowsDue = !isExcused && !isReassigned && !isSubmitted
  const className = cls?.name || vm.className || 'Class'
  // turn in: any unsubmitted work; undo: only plain "submitted" (not graded/returned/returned-for-changes)
  const canTurnIn = !!(sub?.id && (isTodo || isReassigned))
  const canUnsubmit = !!(sub?.id && isSubmitted)

  return (
    <>
      <div className={`scrim${open ? ' on' : ''}`} onClick={onClose} />
      <div className={`sheet-container center${open ? ' on' : ''}`} role="dialog" aria-modal="true" aria-label={vm.title}>
        <div className="sheet center detail">
          <div className="hdr">
            <div className="titles">
              <div className="over t-label-m">{className}</div>
              <div className="t-headline-s">{vm.title}</div>
            </div>
            <button type="button" className="iconbtn" aria-label="Close" onClick={onClose}>
              <Icon name="close" size={22} />
            </button>
          </div>
          <div className="sheet-body">
            {detail?.phase === 'loading' ? (
              <div style={{ padding: '18px 0' }}>
                <div className="skel" style={{ height: 14, width: '46%', marginBottom: 10 }} />
                <div className="skel" style={{ height: 14, width: '72%', marginBottom: 10 }} />
                <div className="skel" style={{ height: 14, width: '58%' }} />
              </div>
            ) : detail?.phase === 'error' ? (
              <div className="stateinfo late">
                <Icon name="error" size={20} />
                <span>Couldn’t load details — check that your Teams session is still active, then retry.</span>
              </div>
            ) : (
              <>
                {/* status banner */}
                {isOverdue && !isSubmitted && !isGraded ? (
                  <div className="stateinfo late">
                    <Icon name="event_busy" size={20} />
                    <span>
                      <strong>Overdue</strong> · {dueLabel(vm)}
                      {vm.allowLate ? ' — late submissions are allowed' : ''}
                    </span>
                  </div>
                ) : isExcused ? (
                  <div className="stateinfo info">
                    <Icon name="interests" size={20} />
                    <span>Excused — not counted toward your grade.</span>
                  </div>
                ) : isReassigned ? (
                  <div className="stateinfo warn">
                    <Icon name="assignment_return" size={20} />
                    <span>
                      <strong>Teacher returned this for changes</strong>
                      {sub?.returnedTs ? ` · ${fmtTime(sub.returnedTs)}` : ''}
                    </span>
                  </div>
                ) : isSubmitted ? (
                  <div className="stateinfo info">
                    <Icon name="task_alt" size={20} />
                    <span>
                      <strong>Turned in</strong>
                      {sub?.submittedTs ? ` · ${fmtTime(sub.submittedTs)}` : ''} — waiting for grading.
                    </span>
                  </div>
                ) : (
                  <div className={`stateinfo ${vm.state.bucket === 'today' ? 'warn' : 'info'}`}>
                    <Icon name="schedule" size={20} />
                    <span>{dueLabel(vm)}</span>
                  </div>
                )}

                {/* grade block */}
                {isGraded && sub ? (
                  <div className="gradeblock">
                    <div>
                      <span className="score">{sub.score ?? '—'}</span>
                      {vm.maxPoints ? <span className="maxof">/ {vm.maxPoints} points</span> : null}
                    </div>
                    {vm.maxPoints ? (
                      <div className="lin det" style={{ margin: '10px 0 4px' }} aria-hidden="true">
                        <i style={{ width: `${Math.max(4, Math.min(100, ((sub.score || 0) / vm.maxPoints) * 100))}%` }} />
                      </div>
                    ) : null}
                    {sub.feedback ? (
                      <div className="prose" style={{ marginTop: 10 }} dangerouslySetInnerHTML={{ __html: sanitizeHtml(sub.feedback) }} />
                    ) : null}
                  </div>
                ) : null}

                {/* key facts */}
                <div className="sectheader">Details</div>
                <div className="kv">
                  <div className="kvrow">
                    <Icon name="school" size={18} />
                    <span>{className}</span>
                  </div>
                  {!bannerShowsDue ? (
                    <div className="kvrow">
                      <Icon name="event" size={18} />
                      <span>{dueLabel(vm)}</span>
                    </div>
                  ) : null}
                  {vm.closeTs ? (
                    <div className="kvrow">
                      <Icon name="event_busy" size={18} />
                      <span>Closes {fmtTime(vm.closeTs)}</span>
                    </div>
                  ) : null}
                  {vm.maxPoints ? (
                    <div className="kvrow">
                      <Icon name="grade" size={18} />
                      <span>{vm.maxPoints} points possible</span>
                    </div>
                  ) : null}
                </div>

                {/* instructions */}
                {instructionsHtml ? (
                  <>
                    <div className="sectheader">Instructions</div>
                    <div className="prose" dangerouslySetInnerHTML={{ __html: instructionsHtml }} />
                  </>
                ) : null}

                {/* resources */}
                {resources.length ? (
                  <>
                    <div className="sectheader">Resources</div>
                    <div className="listmini">
                      {resources.map((r) => {
                        const href = resourceLink(r)
                        const fallback = r?.displayName && r.displayName !== 'Resource' ? r.displayName : hostnameOf(href) || r?.displayName || 'Resource'
                        const name = (href && titles[href]) || fallback
                        return href ? (
                          <a key={r.id || name} href={href} target="_blank" rel="noreferrer">
                            <Icon name={resourceIcon(r)} size={20} />
                            <span className="ellipsis">{name}</span>
                            <Icon name="open_in_new" size={16} style={{ opacity: 0.6 }} />
                          </a>
                        ) : (
                          <a key={r.id || name} href={webUrl} target="_blank" rel="noreferrer">
                            <Icon name={resourceIcon(r)} size={20} />
                            <span className="ellipsis">{name}</span>
                          </a>
                        )
                      })}
                    </div>
                  </>
                ) : null}

                <div className="gap-h" />
                <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap', alignItems: 'center' }}>
                  {canTurnIn ? (
                    <>
                      <button type="button" className="btn btn-filled btn-md" disabled={turnBusy} onClick={() => turnIn('submit')}>
                        <Icon name={turnBusy ? 'sync' : 'task_alt'} size={18} cls={turnBusy ? 'spin' : ''} />
                        {turnBusy ? 'Turning in…' : confirming ? 'Confirm turn-in' : 'Turn in'}
                      </button>
                      {confirming && !turnBusy ? (
                        <button type="button" className="btn btn-text btn-md" onClick={() => setConfirming(false)}>
                          Cancel
                        </button>
                      ) : null}
                    </>
                  ) : null}
                  {canUnsubmit ? (
                    <button type="button" className="btn btn-tonal btn-md" disabled={turnBusy} onClick={() => turnIn('unsubmit')}>
                      <Icon name={turnBusy ? 'sync' : 'undo'} size={18} cls={turnBusy ? 'spin' : ''} />
                      {turnBusy ? 'Undoing…' : 'Undo turn-in'}
                    </button>
                  ) : null}
                  <a className={canTurnIn || canUnsubmit ? 'btn btn-tonal btn-md' : 'btn btn-filled btn-md'} href={webUrl} target="_blank" rel="noreferrer">
                    <Icon name="open_in_new" size={18} />
                    Open in Teams
                  </a>
                </div>
                {turnErr ? (
                  <div className="stateinfo late" style={{ marginTop: 10 }}>
                    <Icon name="error" size={20} />
                    <span>{turnErr}</span>
                  </div>
                ) : null}
                {canTurnIn && confirming && !turnBusy ? (
                  <div className="stateinfo info" style={{ marginTop: 10 }}>
                    <Icon name="info" size={20} />
                    <span>Turn this in? Your teacher will see it as submitted and it moves to Done.</span>
                  </div>
                ) : null}
              </>
            )}
          </div>
        </div>
      </div>
    </>
  )
}
