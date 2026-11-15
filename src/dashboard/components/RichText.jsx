import { createElement, useEffect, useMemo, useState } from 'react'
import { sanitizeHtml } from '../../shared/normalize.js'
import { isGraphUrl, resourceUrl } from '../../shared/resources.js'

function InlineImage({ src, source, ...props }) {
  const embedded = /^data:image\//i.test(src)
  const [image, setImage] = useState(null)
  useEffect(() => {
    if (!src || embedded) return
    let alive = true
    setImage(null)
    const load = async () => {
      try {
        const result = await source.resource?.({ op: 'image', url: src })
        if (!alive) return
        setImage(result?.ok ? result : { error: result?.error || 'Image unavailable.' })
      } catch {
        if (alive) setImage({ error: 'Could not load this image.' })
      }
    }
    load()
    return () => { alive = false }
  }, [src, source, embedded])
  if (!src) return null
  const imageSrc = embedded ? src : image?.dataUrl
  return imageSrc ? (
    <img {...props} src={imageSrc} loading="lazy" decoding="async" />
  ) : (
    <span className="image-placeholder" role="status">{image?.error || `Loading ${props.alt || 'image'}…`}</span>
  )
}

/** Keep images in their original position while resolving protected URLs first. */
export default function RichText({ html, source, onOpenResource, style }) {
  const nodes = useMemo(() => {
    const safe = sanitizeHtml(html)
    return Array.from(new DOMParser().parseFromString(safe, 'text/html').body.childNodes)
  }, [html])
  const renderNode = (node, key) => {
    if (node.nodeType === 3) return node.textContent
    if (node.nodeType !== 1) return null
    const tag = node.tagName.toLowerCase()
    const children = Array.from(node.childNodes).map((child, index) => renderNode(child, `${key}-${index}`))
    if (tag === 'img') {
      return <InlineImage key={key} source={source} src={node.getAttribute('src')} alt={node.getAttribute('alt') || ''} title={node.getAttribute('title') || undefined} width={node.getAttribute('width') || undefined} height={node.getAttribute('height') || undefined} />
    }
    if (tag === 'a') {
      const href = node.getAttribute('href')
      if (isGraphUrl(href)) {
        return <button key={key} type="button" className="prose-link" onClick={() => onOpenResource(resourceUrl(href))}>{children}</button>
      }
      return <a key={key} href={href || undefined} target="_blank" rel="noopener noreferrer">{children}</a>
    }
    return createElement(tag, { key }, tag === 'br' ? undefined : children)
  }
  return <div className="prose" style={style}>{nodes.map((node, index) => renderNode(node, String(index)))}</div>
}
