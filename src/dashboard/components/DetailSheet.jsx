import { useEffect, useMemo, useRef, useState } from 'react'
import Icon from './Icon.jsx'
import { sanitizeHtml, dueLabel } from '../../shared/normalize.js'

const MAX_FILE = 20 * 1024 * 1024

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
      // Graph drive URLs require an auth header; their displayName is the useful label.
      if (!href || !source?.title || /^https:\/\/graph\.microsoft\.com\//i.test(href)) continue
      source.title(href).then((res) => {
        if (alive && res?.ok && res.title) setTitles((cur) => (cur[href] ? cur : { ...cur, [href]: res.title }))
      })
    }
    return () => {
      alive = false
    }
  }, [resources, source])

  /* ---------- student work (submission resources) ---------- */
  const subId = vm?.submission?.id || null
  const prog = vm?.state?.progress
  const workEditable = !!subId && (prog === 'todo' || prog === 'reassigned')
  const showSubmitted = !!subId && !workEditable && (prog === 'submitted' || prog === 'graded' || prog === 'excused')

  const [work, setWork] = useState({ loading: false, resources: [], submittedResources: [] })
  const [workBusy, setWorkBusy] = useState('')
  const [workMsg, setWorkMsg] = useState('')
  const [workErr, setWorkErr] = useState('')
  const [linkOpen, setLinkOpen] = useState(false)
  const [linkUrl, setLinkUrl] = useState('')
  const [linkName, setLinkName] = useState('')
  const [workTick, setWorkTick] = useState(0)
  const fileRef = useRef(null)

  useEffect(() => {
    setWorkMsg('')
    setWorkErr('')
    setLinkOpen(false)
    setLinkUrl('')
    setLinkName('')
    setWorkBusy('')
    setWork({ loading: false, resources: [], submittedResources: [] })
  }, [vm?.id, open])

  useEffect(() => {
    if (!open || !vm || !subId || (!workEditable && !showSubmitted)) return
    let alive = true
    setWork((cur) => ({ ...cur, loading: true }))
    const run = async () => {
      const res = await source.submission?.({ op: 'list', classId: vm.classId, assignmentId: vm.id, submissionId: subId })
      if (!alive) return
      if (res?.ok) {
        setWork({ loading: false, resources: res.resources || [], submittedResources: res.submittedResources || [] })
      } else {
        // keep whatever is already on screen; just stop the spinner
        setWork((cur) => ({ ...cur, loading: false }))
        if (res) setWorkErr(res.error || 'Could not load your attached work.')
      }
    }
    run()
    return () => {
      alive = false
    }
  }, [open, vm?.id, subId, workEditable, showSubmitted, source, workTick])

  const readFileBase64 = (file) =>
    new Promise((resolve, reject) => {
      const fr = new FileReader()
      fr.onload = () => resolve(String(fr.result || '').split(',')[1] || '')
      fr.onerror = () => reject(fr.error)
      fr.readAsDataURL(file)
    })

  const addLink = async () => {
    if (workBusy) return
    setWorkErr('')
    setWorkMsg('')
    setWorkBusy('link')
    const res = await source.submission?.({ op: 'addLink', classId: vm.classId, assignmentId: vm.id, submissionId: subId, url: linkUrl, name: linkName })
    setWorkBusy('')
    if (res?.ok) {
      setLinkOpen(false)
      setLinkUrl('')
      setLinkName('')
      setWorkMsg('Link added.')
      setWorkTick((t) => t + 1)
    } else {
      setWorkErr(res?.error || 'Could not add the link.')
    }
  }

  const onFile = async (e) => {
    const file = e.target.files?.[0]
    e.target.value = ''
    if (!file || workBusy) return
    if (file.size > MAX_FILE) {
      setWorkErr('File is too large — the limit is 20 MB.')
      return
    }
    setWorkErr('')
    setWorkMsg('')
    setWorkBusy('file')
    try {
      const dataBase64 = await readFileBase64(file)
      const res = await source.submission?.({
        op: 'addFile',
        classId: vm.classId,
        assignmentId: vm.id,
        submissionId: subId,
        name: file.name,
        mime: file.type,
        dataBase64,
      })
      if (res?.ok) {
        setWorkMsg(`${file.name} added.`)
        setWorkTick((t) => t + 1)
      } else {
        setWorkErr(res?.error || 'Upload failed.')
      }
    } catch {
      setWorkErr('Could not read that file.')
    } finally {
      setWorkBusy('')
    }
  }

  const removeWork = async (r) => {
    if (workBusy) return
    setWorkErr('')
    setWorkMsg('')
    setWorkBusy(r.id)
    const res = await source.submission?.({ op: 'remove', classId: vm.classId, assignmentId: vm.id, submissionId: subId, resourceId: r.id })
    setWorkBusy('')
    if (res?.ok) {
      setWorkMsg('Removed.')
      setWorkTick((t) => t + 1)
    } else {
      setWorkErr(res?.error || 'Could not remove that item.')
    }
  }

  const openWorkFile = async (fileUrl) => {
    if (!fileUrl) return
    setWorkErr('')
    const res = await source.submission?.({ op: 'openFile', classId: vm.classId, assignmentId: vm.id, submissionId: subId, fileUrl })
    if (!res?.ok) setWorkErr(res?.error || 'Could not open that file — open it in Teams instead.')
    else if (res?.url) window.open(res.url, '_blank', 'noreferrer')
  }

  const workName = (r) =>
    r.link ? (
      <a className="ellipsis" href={r.link} target="_blank" rel="noreferrer">
        {r.displayName}
      </a>
    ) : r.fileUrl ? (
      <button type="button" className="workname" onClick={() => openWorkFile(r.fileUrl)}>
        {r.displayName}
      </button>
    ) : (
      <span className="ellipsis">{r.displayName}</span>
    )

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
  // some assignments forbid student-added resources (Teams hides "Add work" too)
  const allowAttach = d?.allowStudentsToAddResources !== false

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
                        const res = r?.resource || {}
                        const isFile = !res.link && !res.webUrl && !!res.fileUrl
                        const fallback = r?.displayName && r.displayName !== 'Resource' ? r.displayName : hostnameOf(href) || r?.displayName || 'Resource'
                        const name = (href && titles[href]) || fallback
                        if (isFile) {
                          return (
                            <button key={r.id || name} type="button" className="reslink" onClick={() => openWorkFile(res.fileUrl)}>
                              <Icon name={resourceIcon(r)} size={20} />
                              <span className="ellipsis">{name}</span>
                              <Icon name="open_in_new" size={16} style={{ opacity: 0.6 }} />
                            </button>
                          )
                        }
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

                {/* student work: links + files attached to the submission */}
                {workEditable ? (
                  <>
                    <div className="sectheader">Your work</div>
                    {work.resources.length ? (
                      <div className="listmini">
                        {work.resources.map((r) => (
                          <div key={r.id} className="workrow">
                            <Icon name={resourceIcon({ resource: { '@odata.type': r.type } })} size={20} />
                            {workName(r)}
                            {r.studentAdded ? (
                              <button
                                type="button"
                                className="iconbtn workdel"
                                title="Remove"
                                aria-label={`Remove ${r.displayName}`}
                                disabled={!!workBusy}
                                onClick={() => removeWork(r)}
                              >
                                <Icon name={workBusy === r.id ? 'sync' : 'close'} size={18} cls={workBusy === r.id ? 'spin' : ''} />
                              </button>
                            ) : null}
                          </div>
                        ))}
                      </div>
                    ) : (
                      <div className="note" style={{ margin: '2px 0 10px' }}>
                        {work.loading ? 'Loading…' : 'Nothing attached yet — add a link or upload a file.'}
                      </div>
                    )}

                    {linkOpen ? (
                      <div className="linkform">
                        <label className="field">
                          <span>Link URL</span>
                          <input value={linkUrl} onChange={(e) => setLinkUrl(e.target.value)} placeholder="https://…" inputMode="url" autoFocus />
                        </label>
                        <label className="field">
                          <span>Name (optional)</span>
                          <input value={linkName} onChange={(e) => setLinkName(e.target.value)} placeholder="e.g. Research sources" />
                        </label>
                        <div style={{ display: 'flex', gap: 8 }}>
                          <button type="button" className="btn btn-filled btn-md" disabled={!!workBusy || !linkUrl.trim()} onClick={addLink}>
                            <Icon name={workBusy === 'link' ? 'sync' : 'add'} size={18} cls={workBusy === 'link' ? 'spin' : ''} />
                            Add link
                          </button>
                          <button type="button" className="btn btn-text btn-md" disabled={!!workBusy} onClick={() => setLinkOpen(false)}>
                            Cancel
                          </button>
                        </div>
                      </div>
                    ) : (
                      <div style={{ display: 'flex', gap: 10, flexWrap: 'wrap' }}>
                        <button
                          type="button"
                          className="btn btn-tonal btn-md"
                          disabled={!!workBusy || !allowAttach}
                          onClick={() => {
                            setLinkOpen(true)
                            setWorkErr('')
                            setWorkMsg('')
                          }}
                        >
                          <Icon name="link" size={18} />
                          Add link
                        </button>
                        <button type="button" className="btn btn-tonal btn-md" disabled={!!workBusy || !allowAttach} onClick={() => fileRef.current?.click()}>
                          <Icon name={workBusy === 'file' ? 'sync' : 'upload_file'} size={18} cls={workBusy === 'file' ? 'spin' : ''} />
                          {workBusy === 'file' ? 'Uploading…' : 'Upload file'}
                        </button>
                        <input ref={fileRef} type="file" hidden onChange={onFile} />
                      </div>
                    )}
                    {!allowAttach ? (
                      <div className="note" style={{ marginTop: 6 }}>
                        This assignment doesn’t allow students to add their own resources — attach inside Teams instead.
                      </div>
                    ) : null}
                    {workMsg ? (
                      <div className="workstat">
                        <Icon name="check_circle" size={16} />
                        {workMsg}
                      </div>
                    ) : null}
                    {workErr ? (
                      <div className="stateinfo late" style={{ marginTop: 8 }}>
                        <Icon name="error" size={18} />
                        <span>{workErr}</span>
                      </div>
                    ) : null}
                  </>
                ) : null}

                {/* after turn-in: read-only list of what was handed in */}
                {showSubmitted && (work.submittedResources.length || work.resources.length) ? (
                  <>
                    <div className="sectheader">Submitted work</div>
                    <div className="listmini">
                      {(work.submittedResources.length ? work.submittedResources : work.resources).map((r) => (
                        <div key={r.id} className="workrow">
                          <Icon name={resourceIcon({ resource: { '@odata.type': r.type } })} size={20} />
                          {workName(r)}
                        </div>
                      ))}
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
