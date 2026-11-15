import { useEffect } from 'react'
import Icon from './Icon.jsx'

const fmtTime = (ts) => (ts ? new Date(ts).toLocaleString(undefined, { month: 'short', day: 'numeric', hour: 'numeric', minute: '2-digit' }) : '—')

/** Settings sheet: theme, data source, session status, cache controls. */
export default function SettingsSheet({ open, onClose, theme, setTheme, source, meta, syncedAt, onRefresh, onClear }) {
  useEffect(() => {
    if (!open) return
    const onKey = (e) => e.key === 'Escape' && onClose?.()
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [open, onClose])

  return (
    <>
      <div className={`scrim${open ? ' on' : ''}`} onClick={onClose} />
      <div className={`sheet-container center${open ? ' on' : ''}`} role="dialog" aria-modal="true" aria-label="Settings">
        <div className="sheet center">
          <div className="hdr">
            <div className="titles">
              <div className="t-title-l">Settings</div>
            </div>
            <button type="button" className="iconbtn" aria-label="Close" onClick={onClose}>
              <Icon name="close" size={22} />
            </button>
          </div>
          <div className="sheet-body" style={{ maxWidth: 520 }}>
            <div className="sectheader">Appearance</div>
            <div className="radio-row">
              {[
                ['system', 'auto_awesome', 'System'],
                ['light', 'light_mode', 'Light'],
                ['dark', 'dark_mode', 'Dark'],
              ].map(([val, icon, label]) => (
                <button
                  key={val}
                  type="button"
                  className={`radio-pill t-label-m${theme === val ? ' sel' : ''}`}
                  aria-pressed={theme === val}
                  onClick={() => setTheme(val)}
                >
                  <Icon name={icon} size={18} />
                  {label}
                </button>
              ))}
            </div>

            <div className="sectheader">Data</div>
            <div className="setrow">
              <Icon name={source.kind === 'demo' ? 'science' : 'cloud_done'} size={20} />
              <div className="grow">
                <div className="t-body-m">{source.kind === 'demo' ? 'Demo data' : 'Live — from your Teams session'}</div>
                <div className="note">Last synced {fmtTime(syncedAt)}</div>
              </div>
            </div>
            {source.kind === 'extension' ? (
              <>
                <div className="setrow">
                  <Icon name="badge" size={20} />
                  <div className="grow">
                    <div className="t-body-m">
                      {meta?.upn || 'Signed-in student'}
                      {meta?.strategy ? ` · via ${meta.strategy === 'graph' ? 'Graph API' : meta.strategy === 'edu' ? 'edu.microsoft.com' : meta.strategy}` : ''}
                    </div>
                    <div className="note">
                      {meta?.capturedAt ? `Token captured ${fmtTime(meta.capturedAt)}` : 'No session token captured yet'}
                      {Array.isArray(meta?.seen) && meta.seen.length ? ` · seen: ${meta.seen.slice(0, 3).join(', ')}` : ''}
                    </div>
                  </div>
                </div>
                {Array.isArray(meta?.attempts) && meta.attempts.length ? (
                  <div className="setrow" style={{ alignItems: 'flex-start' }}>
                    <Icon name="bug_report" size={20} />
                    <div className="grow">
                      <div className="t-body-m">API attempts</div>
                      <div className="note" style={{ wordBreak: 'break-word' }}>
                        {meta.attempts
                          .map((a) => `${a.strategy}:${a.status || a.code}${a.graphCode ? `:${a.graphCode}` : ''} (${String(a.aud || '').replace(/^https?:\/\//, '')})`)
                          .join(' · ')}
                      </div>
                    </div>
                  </div>
                ) : null}
              </>
            ) : null}
            <div style={{ display: 'flex', gap: 10, marginTop: 8, flexWrap: 'wrap' }}>
              <button type="button" className="btn btn-tonal btn-md" onClick={onRefresh}>
                <Icon name="sync" size={18} />
                Refresh now
              </button>
              <button type="button" className="btn btn-text btn-md" onClick={onClear}>
                Clear cached data
              </button>
            </div>

            <div className="divider" />
            <div className="prose" style={{ fontSize: 13 }}>
              <strong>About</strong>
              <p>
                Assignments+ is a student-side companion for Teams Assignments. It reuses the session your own Teams Web tab
                already holds to read your assignments, attach links or files to your submission, and — only when you press
                Turn in — to submit them. Everything stays on this device; nothing is uploaded anywhere except your own
                school account. No admin permissions involved.
              </p>
            </div>
          </div>
        </div>
      </div>
    </>
  )
}
