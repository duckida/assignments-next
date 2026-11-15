/** Pure view-model mapping for education assignments. Used by both the live
 * (Graph) path and the demo path so the UI has one data contract.
 */

/** @typedef {{
 *  id: string, classId: string, title: string,
 *  dueTs: number|null, closeTs: number|null, assignedTs: number|null,
 *  status: string, allowLate: boolean|null,
 *  maxPoints: number|null, webUrl: string|null, lastModified: number|null,
 *  submission?: { id:string, status:string, submittedTs:number|null, returnedTs:number|null, score:number|null, feedback:string|null, assignmentId:string }
 * }} AssignmentVM
 */

const VISIBLE_ASSIGNMENT_STATUSES = new Set(['assigned', 'published', 'inactive', 'unknownFutureValue'])

export function mapAssignment(raw) {
  if (!raw || !raw.id || !raw.classId) return null
  return {
    id: raw.id,
    classId: raw.classId,
    className: raw.className || raw.classDisplayName || null,
    title: raw.displayName || 'Untitled assignment',
    dueTs: raw.dueDateTime ? Date.parse(raw.dueDateTime) : null,
    closeTs: raw.closeDateTime ? Date.parse(raw.closeDateTime) : null,
    assignedTs: raw.assignedDateTime ? Date.parse(raw.assignedDateTime) : null,
    status: raw.status || 'assigned',
    allowLate: raw.allowLateSubmissions ?? null,
    // v1.0 grading: { "@odata.type": "#microsoft.graph.educationAssignmentPointsGradeType", maxPoints: n }
    maxPoints:
      raw.grading && raw.grading['@odata.type'] && raw.grading['@odata.type'].endsWith('PointsGradeType')
        ? (raw.grading.maxPoints ?? null)
        : null,
    webUrl: raw.webUrl || null,
    lastModified: raw.lastModifiedDateTime ? Date.parse(raw.lastModifiedDateTime) : null,
  }
}

/** Extract grade + feedback from outcomes collection */
export function mapSubmission(raw, assignmentId) {
  if (!raw) return null
  let score = null
  let feedback = null
  const outcomes = Array.isArray(raw.outcomes) ? raw.outcomes : []
  for (const o of outcomes) {
    const t = o && o['@odata.type']
    if (!t) continue
    if (t.endsWith('educationPointsOutcome')) {
      const pts = o.points || {}
      if (pts.points != null) score = Number(pts.points)
      if (pts.publishedPoints != null && score == null) score = Number(pts.publishedPoints)
    } else if (t.endsWith('educationFeedbackOutcome')) {
      if (o.feedback) feedback = o.feedback.text || o.feedback.feedback || feedback
    } else if (o.feedback) {
      feedback = o.feedback.text || feedback
    }
  }
  return {
    assignmentId: assignmentId ?? raw.assignmentId,
    id: raw.id || '',
    status: raw.status || 'working',
    submittedTs: raw.submittedDateTime ? Date.parse(raw.submittedDateTime) : null,
    returnedTs: raw.returnedDateTime ? Date.parse(raw.returnedDateTime) : null,
    score,
    feedback,
  }
}

/** ---------- derived view state ---------- */

export const ONE = { min: 60_000, hour: 3_600_000, day: 86_400_000 }

/** bucket: 'overdue' | 'today' | 'soon'(<=3d) | 'later' | 'none' */
export function dueBucket(a, now = Date.now()) {
  if (progress(a) !== 'todo') return 'none'
  if (a.dueTs == null) return 'none'
  if (a.dueTs < now) return 'overdue'
  if (sameDay(a.dueTs, now)) return 'today'
  if (a.dueTs - now <= 3 * ONE.day) return 'soon'
  return 'later'
}

/** progress: 'todo' | 'submitted' | 'graded' | 'reassigned' | 'excused' */
export function progress(a) {
  const s = a.submission
  if (!s) return 'todo'
  if (s.status === 'excused') return 'excused'
  if (s.status === 'returned' || s.status === 'graded') return s.score != null || s.feedback ? 'graded' : 'submitted'
  if (s.status === 'reassigned') return 'reassigned'
  if (s.status === 'submitted') return 'submitted'
  return 'todo'
}

export function sameDay(ts, ref) {
  const a = new Date(ts)
  const b = new Date(ref)
  return a.getFullYear() === b.getFullYear() && a.getMonth() === b.getMonth() && a.getDate() === b.getDate()
}

/** Human label for a due timestamp, e.g. "Due today · 3:00 PM" / "Fri · Oct 24" */
export function dueLabel(a, now = Date.now()) {
  if (a.dueTs == null) return 'No due date'
  const d = new Date(a.dueTs)
  const time = d.toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' })
  const t = new Date(now)
  const days = Math.round((startOfDay(a.dueTs) - startOfDay(now)) / ONE.day)
  if (days < -1) {
    const base = d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
    return `Due ${base}`
  }
  if (days === -1) return 'Due yesterday · ' + time
  if (days === 0) return 'Due today · ' + time
  if (days === 1) return 'Due tomorrow · ' + time
  if (days <= 6) {
    const wd = d.toLocaleDateString(undefined, { weekday: 'short' })
    return `Due ${wd} · ${time}`
  }
  return `Due ${d.toLocaleDateString(undefined, { month: 'short', day: 'numeric' })} · ${time}`
}

function startOfDay(ts) {
  const d = new Date(ts)
  d.setHours(0, 0, 0, 0)
  return d.getTime()
}

export function shortDueLabel(a, now = Date.now()) {
  if (a.dueTs == null) return 'No due'
  const days = Math.round((startOfDay(a.dueTs) - startOfDay(now)) / ONE.day)
  if (days === 0) return new Date(a.dueTs).toLocaleTimeString(undefined, { hour: 'numeric', minute: '2-digit' })
  if (days === 1) return 'Tomorrow'
  if (days < 0) {
    return new Date(a.dueTs).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
  }
  if (days <= 6) return new Date(a.dueTs).toLocaleDateString(undefined, { weekday: 'short' })
  return new Date(a.dueTs).toLocaleDateString(undefined, { month: 'short', day: 'numeric' })
}

/** ---------- sorting / grouping ---------- */

export function sortAssignments(list, now = Date.now()) {
  return [...list].sort((a, b) => {
    const pa = todoFirst(a)
    const pb = todoFirst(b)
    if (pa !== pb) return pa - pb
    const da = a.dueTs ?? Infinity
    const db = b.dueTs ?? Infinity
    if (da !== db) return da - db
    return a.title.localeCompare(b.title)
  })
}

function todoFirst(a) {
  const p = progress(a)
  if (p === 'todo') return 0
  if (p === 'reassigned') return 1
  return 2 // submitted / graded / excused sink to bottom
}

/** ---------- instructions sanitizer (dashboard-only, runs in DOM context) ---------- */

const ALLOWED_TAGS = new Set([
  'P','DIV','BR','B','I','EM','STRONG','U','UL','OL','LI','A','SPAN','IMG',
  'H1','H2','H3','H4','H5','BLOCKQUOTE','PRE','CODE','TABLE','THEAD','TBODY','TR','TD','TH',
])
const ALLOWED_ATTRS = new Set(['href'])
// attributes kept on <img> (src validated separately — http(s) or data:image only)
const IMG_ATTRS = new Set(['src', 'alt', 'title', 'width', 'height'])

/** Whitelist-sanitize teacher-authored instructions HTML. Links open in a new tab;
 *  images are kept (lazy-loaded) when their src is http(s) or an inline image.
 */
export function sanitizeHtml(html) {
  if (!html) return ''
  let doc
  try {
    doc = new DOMParser().parseFromString(html, 'text/html')
  } catch {
    return escapeHtml(String(html))
  }
  const walk = (node) => {
    for (const child of Array.from(node.childNodes)) {
      if (child.nodeType === 1 /* ELEMENT */) {
        walk(child)
        const tag = child.tagName.toUpperCase()
        if (tag === 'SCRIPT' || tag === 'STYLE' || tag === 'IFRAME' || tag === 'SVG' || tag === 'MATH' || tag === 'OBJECT') {
          child.remove()
          continue
        }
        if (!ALLOWED_TAGS.has(tag)) {
          // unwrap: replace with its children
          const parent = child.parentNode
          while (child.firstChild) parent.insertBefore(child.firstChild, child)
          parent.removeChild(child)
          continue
        }
        for (const attr of Array.from(child.attributes)) {
          const name = attr.name.toLowerCase()
          const keep = tag === 'A' ? ALLOWED_ATTRS.has(name) : tag === 'IMG' && IMG_ATTRS.has(name)
          if (!keep) child.removeAttribute(attr.name)
        }
        if (tag === 'A') {
          const href = child.getAttribute('href') || ''
          if (/^\s*javascript:/i.test(href)) child.removeAttribute('href')
          child.setAttribute('target', '_blank')
          child.setAttribute('rel', 'noopener noreferrer')
        } else if (tag === 'IMG') {
          const src = (child.getAttribute('src') || '').trim()
          // http(s) remote images and inline raster/SVG data images are safe in an <img>
          // (SVG loaded as an image can never execute script). Anything else is dropped.
          if (!/^(https?:|data:image\/)/i.test(src)) {
            child.remove()
            continue
          }
          child.setAttribute('loading', 'lazy')
          child.setAttribute('decoding', 'async')
        }
      } else if (child.nodeType === 8) {
        child.remove()
      }
    }
  }
  walk(doc.body)
  return doc.body.innerHTML
}

export function escapeHtml(s) {
  return String(s).replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]))
}

/** Stable pseudo-random hue from a class id (for avatars/badges) */
export function classHue(id) {
  if (!id) return 260
  let h = 0
  for (const c of id) h = (h * 31 + c.charCodeAt(0)) >>> 0
  return h % 360
}
