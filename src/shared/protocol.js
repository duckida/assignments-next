/** Message protocol shared between all extension contexts.
 * Bundled independently into background, content bridge and dashboard bundles.
 */

// Message types sent over chrome.runtime messaging
export const MSG = {
  // content bridge -> service worker
  TOKEN: 'mad3/token',            // { payload: {token, expMs, scp} }
  SNAPSHOT: 'mad3/snapshot',      // { payload: {url, ok, body} } raw edu JSON seen on the wire
  // dashboard -> service worker
  OPEN_DASHBOARD: 'mad3/open-dashboard', // injected FAB -> SW (open the full dashboard tab)
  SYNC: 'mad3/sync',              // {} -> { ok, code?, data?, meta? }
  DETAIL: 'mad3/detail',          // { classId, id } -> { ok, code?, assignment? }
  SUBMIT: 'mad3/submit',          // { classId, id, submissionId?, action: 'submit'|'unsubmit' } -> { ok, error? }
  SUBMISSION: 'mad3/submission',  // { op: 'list'|'addLink'|'addFile'|'remove'|'openFile', classId, assignmentId, ... }
  TITLE: 'mad3/title',            // { url } -> { ok, title? }  (fetch link's HTML <title>)
  CLEAR: 'mad3/clear',            // {} -> { ok }
}

// Marker for window.postMessage traffic between MAIN world and the isolated bridge
export const WINDOW_MARKER = '__mad3'
export const WT = {
  TOKEN: 'TOKEN',
  SNAPSHOT: 'SNAPSHOT',
}
