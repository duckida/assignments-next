/** Minimal unverified JWT decoding + Graph-audience detection.
 * Tokens are captured from the Authorization header the Teams Web client itself sends;
 * they are never modified, and are only used to replay read-only requests to the
 * same endpoints the user's session already accesses.
 */

const GRAPH_APP_ID = '00000003-0000-0000-c000-000000000000' // "aud" of Graph API tokens

function base64UrlDecode(seg) {
  // base64url -> base64 -> JSON string
  const b64 = seg.replace(/-/g, '+').replace(/_/g, '/')
  const pad = b64.length % 4 === 0 ? b64 : b64 + '='.repeat(4 - (b64.length % 4))
  const bin = atob(pad)
  try {
    return decodeURIComponent(
      Array.from(bin)
        .map((c) => '%' + c.charCodeAt(0).toString(16).padStart(2, '0'))
        .join('')
    )
  } catch {
    return bin
  }
}

/** @returns {object|null} decoded JWT claims or null when unparseable */
export function decodeJwtPayload(token) {
  if (typeof token !== 'string') return null
  const parts = token.split('.')
  if (parts.length !== 3) return null
  try {
    return JSON.parse(base64UrlDecode(parts[1]))
  } catch {
    return null
  }
}

function audIsGraph(aud) {
  return (
    aud === GRAPH_APP_ID ||
    String(aud).startsWith('https://graph.microsoft.com') ||
    String(aud).startsWith('https://edu.graph.microsoft.com')
  )
}

/** true when the token is an access token aimed at Microsoft Graph */
export function isGraphAccessToken(token) {
  const p = decodeJwtPayload(token)
  if (!p) return false
  if (p.aud && !audIsGraph(p.aud)) return false
  // must look like an access token (id tokens also have aud: a clientId, already rejected above)
  return !!p.aud && (p.exp || p.iat)
}

/** Kept for future use (e.g. AAD-Graph legacy shapes); not used by the hook */
export const AUD_GRAPH = GRAPH_APP_ID

/** Summary used for storing "best" token per audience */
export function tokenSummary(token) {
  const p = decodeJwtPayload(token)
  if (!p) return null
  const scp = Array.isArray(p.scp) ? p.scp.join(' ') : p.scp || p.roles || ''
  return {
    token,
    aud: p.aud,
    scp,
    app: p.appid || p.azp || p.azpacr ? p.appid || p.azp : p.azp || p.appid || '',
    tid: p.tid || '',
    upn: p.upn || p.preferred_username || '',
    expMs: p.exp ? p.exp * 1000 : 0,
  }
}
