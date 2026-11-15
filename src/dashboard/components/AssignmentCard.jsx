import Icon from './Icon.jsx'
import { ClassAvatar } from './Chips.jsx'
import { shortDueLabel } from '../../shared/normalize.js'
import { bestClassName } from '../../shared/classes.js'

/** One assignment row — elevated M3 card with class avatar, due chip and status. */
export default function AssignmentCard({ a, cls, onOpen }) {
  const state = a.state // derived in App: {bucket, progress}
  const late = state.bucket === 'overdue'
  const today = state.bucket === 'today'

  let statusIcon = null
  if (state.progress === 'graded') statusIcon = <Icon name="verified" size={19} cls="statusicon done" />
  else if (state.progress === 'submitted') statusIcon = <Icon name="task_alt" size={19} cls="statusicon" />
  else if (state.progress === 'reassigned') statusIcon = <Icon name="assignment_return" size={19} cls="statusicon" />
  else if (state.progress === 'excused') statusIcon = <Icon name="interests" size={19} cls="statusicon" />
  else if (late) statusIcon = <Icon name="error" size={19} cls="statusicon" style={{ color: 'var(--sys-error)' }} />
  else statusIcon = <Icon name="schedule" size={19} cls="statusicon" />

  return (
    <article
      className="card outlined clickable acard"
      role="button"
      tabIndex={0}
      onClick={() => onOpen(a)}
      onKeyDown={(e) => {
        if (e.key === 'Enter' || e.key === ' ') {
          e.preventDefault()
          onOpen(a)
        }
      }}
      aria-label={`Open ${a.title}`}
    >
      <ClassAvatar cls={cls} />
      <div className="body">
        <div className="title">{a.title}</div>
        <div className="sub">
          {bestClassName(cls?.name, a.className) || 'Unknown class'} {cls?.description ? <span className="sep"> {cls.description}</span> : null}
        </div>
        {state.progress === 'graded' && a.submission?.score != null ? (
          <div>
            <span className="gradechip done t-label-m">
              {a.submission.score}
              {a.maxPoints ? ` / ${a.maxPoints}` : ''}
            </span>
          </div>
        ) : null}
      </div>
      <div className="meta">
        <span className={`due t-label-m${late ? ' late' : today ? ' today' : ''}`}>
          {late ? <Icon name="error" size={14} /> : today ? <Icon name="schedule" size={14} /> : null}
          {shortDueLabel(a)}
        </span>
        {statusIcon}
      </div>
    </article>
  )
}
