import { graphBatch, graphPaged } from './graph.js'
import { assignmentClassName, bestClassName, classKey, mergeAssignmentClasses } from '../shared/classes.js'

const ROSTER_SCOPES = /(?:EduRoster\.|Team\.ReadBasic|TeamSettings\.Read|Group(?:Member)?\.Read|Directory\.Read)/i

/** Assignments and roster access may be in different tokens for the same student. */
export function sameAccount(a, b) {
  if (a.tid && b.tid && a.tid !== b.tid) return false
  if (a.oid && b.oid && a.oid !== b.oid) return false
  return !(a.upn && b.upn && a.upn.toLowerCase() !== b.upn.toLowerCase())
}

export async function resolveClasses(assignments, candidates, cached = []) {
  const byId = new Map(mergeAssignmentClasses(Array.isArray(cached) ? cached : [], assignments).map((c) => [classKey(c.id), c]))
  const wanted = [...new Set(assignments.map((a) => a.classId).filter(Boolean))]
  const wantedKeys = new Set(wanted.map(classKey))
  const missing = () => wanted.filter((id) => !bestClassName(byId.get(classKey(id))?.name))
  const merge = (c, id = c?.id) => {
    if (!id || !c) return
    const key = classKey(id)
    const prev = byId.get(key)
    byId.set(key, {
      ...prev,
      id: prev?.id || id,
      name: bestClassName(c.displayName, c.name, prev?.name),
      description: c.description || prev?.description || '',
      webUrl: c.webUrl || prev?.webUrl || null,
    })
  }
  const tokens = [...candidates].sort((a, b) => Number(ROSTER_SCOPES.test(b.scp || '')) - Number(ROSTER_SCOPES.test(a.scp || '')))
  for (const c of tokens) {
    if (!missing().length) break
    const scopes = String(c.scp || '')
    // joinedTeams explicitly does not support $select/$top. An education token
    // alone need not carry Team.ReadBasic.All, so try the other captured tokens.
    const lists = await Promise.allSettled([
      graphPaged(c.token, '/education/me/classes?$select=id,displayName,description&$top=100'),
      graphPaged(c.token, '/me/joinedTeams'),
    ])
    for (const result of lists) {
      if (result.status === 'fulfilled') result.value.forEach((cls) => merge(cls))
    }

    const paths = new Map()
    for (const id of missing()) {
      const key = encodeURIComponent(id)
      // Some tenants expose names on the group/team but not the education roster.
      if (!scopes || /EduRoster\./i.test(scopes)) paths.set(`/education/classes/${key}?$select=id,displayName,description`, id)
      if (!scopes || /Team\.|TeamSettings\.|Directory\./i.test(scopes)) paths.set(`/teams/${key}`, id)
      if (!scopes || /Group|Directory\./i.test(scopes)) paths.set(`/groups/${key}?$select=id,displayName,description`, id)
    }
    if (paths.size) {
      const names = await graphBatch(c.token, [...paths.keys()]).catch(() => new Map())
      for (const [path, body] of names) merge(body, paths.get(path))
    }
  }

  // Keep the assignment's own name if the roster was missing it.
  for (const a of assignments) {
    const name = assignmentClassName(a)
    if (name && !bestClassName(byId.get(classKey(a.classId))?.name)) merge({ name }, a.classId)
  }
  return [...byId.values()].filter((c) => wantedKeys.has(classKey(c.id)))
}
