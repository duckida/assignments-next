import { useEffect, useState } from 'react'
import Icon from './Icon.jsx'
import { describeResource, isGraphUrl } from '../../shared/resources.js'

function resourceIcon(type, isImage) {
  if (isImage) return 'image'
  if (/Link/i.test(type)) return 'link'
  if (/Word/i.test(type)) return 'description'
  if (/Excel/i.test(type)) return 'table_chart'
  if (/PowerPoint/i.test(type)) return 'co_present'
  if (/Pdf/i.test(type)) return 'picture_as_pdf'
  if (/Media/i.test(type)) return 'play_circle'
  return 'attach_file'
}

/** The same authenticated file/image path for teacher resources and student work. */
export default function ResourceItem({ resource, source, onOpen, fallbackUrl, title, children }) {
  const info = describeResource(resource)
  const [preview, setPreview] = useState(null)
  const graph = isGraphUrl(info.url)
  // Images always resolve through the service worker (Graph URLs need auth).
  // File resources whose display name was missing get one metadata read: Graph
  // reports the real name and whether the file is an image (for inline preview).
  const loadUrl = info.isImage ? info.imageUrl : info.hasName ? null : info.apiUrl || (graph ? info.url : null)
  useEffect(() => {
    let alive = true
    setPreview(null)
    if (!loadUrl) return
    const load = async () => {
      try {
        const res = await source.resource?.({ op: info.isImage ? 'image' : 'preview', url: loadUrl })
        if (!alive) return
        if (res?.ok) setPreview(res)
        else if (info.isImage) setPreview({ error: res?.error || 'Preview unavailable.' })
      } catch {
        if (alive && info.isImage) setPreview({ error: 'Preview unavailable.' })
      }
    }
    load()
    return () => { alive = false }
  }, [source, loadUrl, info.isImage])

  const name = title || (info.hasName ? info.name : preview?.name) || info.name
  const href = info.url || fallbackUrl
  const showsImage = info.isImage || !!preview?.isImage
  const label = <>
    <Icon name={resourceIcon(info.type, showsImage)} size={20} />
    <span className="ellipsis">{name}</span>
    {href ? <Icon name="open_in_new" size={16} style={{ opacity: 0.6 }} /> : null}
  </>
  return (
    <div className="resource-item">
      <div className="resource-row">
        {!href ? (
          <span className="reslink static">{label}</span>
        ) : graph ? (
          <button type="button" className="reslink" onClick={() => onOpen(info.url)}>{label}</button>
        ) : (
          <a className="reslink" href={href} target="_blank" rel="noopener noreferrer">{label}</a>
        )}
        {children}
      </div>
      {showsImage && loadUrl ? (
        preview?.dataUrl ? (
          <img className="resource-image" src={preview.dataUrl} alt={name} loading="lazy" decoding="async" />
        ) : preview?.error ? (
          <div className="image-placeholder" role="status">{preview.error}</div>
        ) : (
          <div className="image-placeholder" role="status">Loading image…</div>
        )
      ) : null}
    </div>
  )
}
