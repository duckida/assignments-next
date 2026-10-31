import Icon from './Icon.jsx'

/** M3 filter chip — optional class avatar/dot + selected state layer */
export function FilterChip({ selected, onClick, children, leading, ariaLabel }) {
  return (
    <button
      type="button"
      className={`chip filter${selected ? ' sel' : ''}`}
      aria-pressed={!!selected}
      aria-label={ariaLabel}
      onClick={onClick}
    >
      {leading}
      <span className="label t-label-m">{children}</span>
      {selected ? <Icon name="check" size={16} cls="lead" /> : null}
    </button>
  )
}

/** M3 segmented button (single-select) */
export function Segment({ options, value, onChange }) {
  return (
    <div className="segment" role="group">
      {options.map((o) => (
        <button key={o.value} type="button" aria-pressed={value === o.value} onClick={() => onChange(o.value)}>
          {o.icon ? <Icon name={o.icon} size={17} /> : null}
          {o.label}
        </button>
      ))}
    </div>
  )
}

export function IconButton({ name, label, onClick, filled, spin }) {
  return (
    <button type="button" className={`iconbtn${filled ? ' filled' : ''}`} aria-label={label} title={label} onClick={onClick}>
      <Icon name={name} size={22} cls={spin ? 'spin' : ''} />
    </button>
  )
}

/** Class avatar: initial + hue derived from the class id */
export function ClassAvatar({ cls, small }) {
  const h = hueFor(cls?.id)
  const initial = (cls?.name || '?').trim().charAt(0).toUpperCase()
  return (
    <span className={`avatar${small ? ' avatar-sm' : ''}`} style={{ ['--h']: h }} aria-hidden="true">
      {initial}
    </span>
  )
}

function hueFor(id) {
  if (!id) return 260
  let h = 0
  for (const c of id) h = (h * 31 + c.charCodeAt(0)) >>> 0
  return h % 360
}
