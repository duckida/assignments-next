/** Class IDs are GUIDs in Graph, but casing can differ between API responses. */
export const classKey = (id) => String(id || '').trim().toLowerCase()

/** A placeholder must not hide a real name from another source. */
export function bestClassName(...values) {
  return values.find((v) => typeof v === 'string' && v.trim() && !/^(?:class|team|unknown class)$/i.test(v.trim()))?.trim() || null
}

export function assignmentClassName(a) {
  return bestClassName(a?.className, a?.classDisplayName, a?.teamName, a?.teamDisplayName, a?.class?.displayName, a?.team?.displayName)
}

export function mergeAssignmentClasses(classes = [], assignments = []) {
  const byId = new Map()
  for (const c of classes || []) {
    if (!c?.id) continue
    const key = classKey(c.id)
    const prev = byId.get(key)
    byId.set(key, { ...prev, ...c, name: bestClassName(c.name, c.displayName, prev?.name) })
  }
  for (const a of assignments || []) {
    if (!a?.classId) continue
    const key = classKey(a.classId)
    const prev = byId.get(key)
    byId.set(key, { ...prev, id: a.classId, name: bestClassName(prev?.name, assignmentClassName(a)) })
  }
  return [...byId.values()]
}
