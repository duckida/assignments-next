/** Resource URLs are not necessarily browser links: Graph drive URLs need auth. */
const GRAPH_ORIGIN = 'https://graph.microsoft.com'
const GRAPH_ORIGINS = new Set([GRAPH_ORIGIN, 'https://graph.microsoft.us'])
const IMAGE_EXT = /\.(?:avif|bmp|gif|ico|jpe?g|png|svg|webp)(?:$|[?#\s])/i

export function resourceUrl(value) {
  if (typeof value !== 'string' || !value.trim()) return null
  try {
    const text = value.trim()
    const url = /^\/(?:v1\.0|beta)\//i.test(text) ? new URL(text, GRAPH_ORIGIN) : new URL(text)
    if (!['http:', 'https:'].includes(url.protocol) || url.username || url.password) return null
    return url.href
  } catch {
    return null
  }
}

export function isGraphUrl(value) {
  const url = resourceUrl(value)
  return !!url && GRAPH_ORIGINS.has(new URL(url).origin)
}

/** All documented drive addressing forms, including /content and path-based URLs. */
export function driveItemUrl(value) {
  if (!isGraphUrl(value)) return null
  const url = new URL(resourceUrl(value))
  const path = url.pathname
  const drive = /^\/(?:v1\.0|beta)\/(?:drives\/[^/]+|(?:groups|sites|users)\/[^/]+\/drive|me\/drive|drive)\/(?:items\/[^/]+|root(?::|\/|$))/i
  const share = /^\/(?:v1\.0|beta)\/shares\/[^/]+\/driveItem(?:\/|$)/i
  if ((!drive.test(path) && !share.test(path)) || /\/thumbnails(?:\/|$)/i.test(path)) return null
  url.pathname = path.replace(/\/content\/?$/i, '').replace(/\/$/, '').replace(/:$/, '')
  url.search = ''
  url.hash = ''
  return url.href
}

export function isImageResource(resource) {
  const r = resource?.resource || resource || {}
  if (/^image\//i.test(r.mimeType || r.file?.mimeType || '')) return true
  return [r.displayName, r.name, resource?.displayName, r.fileUrl, r.webUrl, r.link].some((v) => IMAGE_EXT.test(String(v || '')))
}

/** Graph nests displayName in resource; internal EDU responses can be flat. */
export function describeResource(resource) {
  const r = resource?.resource || resource || {}
  const link = resourceUrl(r.link)
  const webUrl = resourceUrl(r.webUrl)
  const fileUrl = resourceUrl(r.fileUrl)
  // Prefer a browser (SharePoint) URL for clicks, but use Graph's file URL for
  // protected image data — its bearer token is handled by the service worker.
  const url = link || webUrl || fileUrl
  const name = [r.displayName, resource?.displayName, r.name].find(
    (n) => typeof n === 'string' && n.trim() && !/^(?:resource|file)$/i.test(n.trim())
  )
  const isImage = isImageResource(resource)
  return {
    url,
    // Authenticated Graph drive URL for this file, when the API supplied one.
    apiUrl: isGraphUrl(fileUrl) ? fileUrl : null,
    imageUrl: isImage ? (isGraphUrl(fileUrl) ? fileUrl : url) : null,
    name: name?.trim() || (url ? new URL(url).hostname.replace(/^www\./, '') : 'Resource'),
    hasName: !!name,
    type: String(r['@odata.type'] || r.type || ''),
    isImage,
    isFile: !!r.fileUrl || /education(?:File|Media|Word|Excel|PowerPoint|Pdf)Resource/i.test(r['@odata.type'] || r.type || ''),
  }
}
