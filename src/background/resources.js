import { GraphError, graphFetch } from './graph.js'
import { driveItemUrl, isGraphUrl, isImageResource, resourceUrl } from '../shared/resources.js'

const MAX_IMAGE = 20 * 1024 * 1024

/** Return bytes as a data URL: runtime messaging is JSON, and SWs have no blob URLs. */
export async function imageData(url, { token, credentials = 'omit', mimeType } = {}) {
  let res
  try {
    res = await fetch(url, {
      headers: token ? { Authorization: `Bearer ${token}` } : {},
      credentials,
      signal: AbortSignal.timeout(15_000),
    })
  } catch (e) {
    throw new GraphError('image-network', 0, 'Could not reach the image host.', String(e))
  }
  if (!res.ok) throw new GraphError('image', res.status, 'Could not load this image.')
  const type = (res.headers.get('content-type') || '').split(';')[0].trim().toLowerCase()
  const mime = !type || type === 'application/octet-stream' ? mimeType || type : type
  if (!/^image\/[a-z0-9.+-]+$/i.test(mime)) throw new GraphError('not-image', 400, 'This resource did not return an image.')
  if (Number(res.headers.get('content-length')) > MAX_IMAGE) {
    await res.body?.cancel()
    throw new GraphError('image-too-large', 413, 'This image is too large to preview. Open the file to view it.')
  }
  const reader = res.body?.getReader()
  const chunks = []
  let size = 0
  if (reader) {
    try {
      for (;;) {
        const { done, value } = await reader.read()
        if (done) break
        size += value.byteLength
        if (size > MAX_IMAGE) {
          await reader.cancel()
          throw new GraphError('image-too-large', 413, 'This image is too large to preview. Open the file to view it.')
        }
        chunks.push(value)
      }
    } finally {
      reader.releaseLock()
    }
  } else {
    const bytes = new Uint8Array(await res.arrayBuffer())
    if (bytes.length > MAX_IMAGE) throw new GraphError('image-too-large', 413, 'This image is too large to preview.')
    chunks.push(bytes)
    size = bytes.length
  }
  if (!size) throw new GraphError('image-empty', 400, 'This image is empty.')
  let binary = ''
  for (const chunk of chunks) {
    for (let offset = 0; offset < chunk.length; offset += 8192) binary += String.fromCharCode(...chunk.subarray(offset, offset + 8192))
  }
  return `data:${mime};base64,${btoa(binary)}`
}

/** Resolve Graph file references with auth; never navigate a tab to the API URL. */
export async function graphResource(token, value, op) {
  const url = driveItemUrl(value)
  if (!url) throw new GraphError('unsupported-resource', 400, 'This Graph resource is not a file link.')
  // No $select: driveItem's default representation includes name/webUrl/file/size
  // and @microsoft.graph.downloadUrl on every supported drive addressing form.
  const item = await graphFetch(token, url)
  const name = item?.name || ''
  const isImage = isImageResource(item)
  const webUrl = resourceUrl(item?.webUrl)
  const result = { ok: true, name, isImage, url: webUrl && !isGraphUrl(webUrl) ? webUrl : null }
  if (op === 'open') {
    if (!result.url) throw new GraphError('no-web-url', 404, 'This file has no browser link. Open it in Teams.')
    return result
  }
  if (op === 'image' || isImage) {
    if (item?.size > MAX_IMAGE) throw new GraphError('image-too-large', 413, 'This image is too large to preview. Open the file to view it.')
    const downloadUrl = resourceUrl(item?.['@microsoft.graph.downloadUrl'])
    if (!downloadUrl || !downloadUrl.startsWith('https://') || isGraphUrl(downloadUrl)) {
      throw new GraphError('no-download-url', 404, 'This image has no download link. Open the file to view it.')
    }
    // Graph's short-lived download URL is already authenticated. No bearer token
    // is sent to SharePoint/CDN, included in the DOM, or saved to local storage.
    result.dataUrl = await imageData(downloadUrl, { mimeType: item.file?.mimeType })
    result.isImage = true
  }
  return result
}
