import Icon from './Icon.jsx'

export function Section({ title, count, children }) {
  return (
    <section>
      <div className="section-h">
        <span className="name">{title}</span>
        {count != null ? <span className="count">{count}</span> : null}
      </div>
      <div className="assignments">{children}</div>
    </section>
  )
}

export function SkeletonList() {
  return (
    <div className="assignments" aria-busy="true">
      {[0, 1, 2, 3].map((i) => (
        <div key={i} className="card outlined" style={{ padding: 14, display: 'flex', gap: 14 }}>
          <div className="skel" style={{ width: 40, height: 40, borderRadius: 20, flex: 'none' }} />
          <div style={{ flex: 1 }}>
            <div className="skel" style={{ height: 15, width: '62%', marginBottom: 8 }} />
            <div className="skel" style={{ height: 12, width: '38%' }} />
          </div>
          <div className="skel" style={{ height: 22, width: 70, borderRadius: 6 }} />
        </div>
      ))}
    </div>
  )
}

export function StateAction({ icon, headline, body, actions }) {
  return (
    <div className="statewrap">
      <div className="bigicon">
        <Icon name={icon} size={64} />
      </div>
      <div className="t-title-l">{headline}</div>
      <div className="caption t-body-m">{body}</div>
      <div className="actions">{actions}</div>
    </div>
  )
}
