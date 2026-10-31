/** Material Symbols Rounded icon (font vendored locally — works offline in the extension). */
export default function Icon({ name, size = 24, cls = '', style }) {
  return (
    <span
      className={`msym ${cls}`.trim()}
      style={{ fontSize: size, ...style }}
      aria-hidden="true"
      focusable="false"
    >
      {name}
    </span>
  )
}
