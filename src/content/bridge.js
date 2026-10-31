/**
 * bridge.js — isolated-world content script on Teams Web pages.
 * Relays MAIN-world capture events into the extension service worker, and
 * injects a small floating FAB that opens the Material dashboard.
 */
import { MSG, WINDOW_MARKER, WT } from '../shared/protocol.js'

const PAB = (() => {
  try {
    return chrome && typeof chrome.runtime !== 'undefined'
  } catch {
    return false
  }
})()

function send(type, payload) {
  if (!PAB) return
  try {
    const r = chrome.runtime.sendMessage({ type, payload })
    if (r && typeof r.catch === 'function') r.catch(() => {})
  } catch {}
}

/* relay MAIN-world events */
window.addEventListener('message', (event) => {
  try {
    if (event.source !== window) return
    const d = event.data
    if (!d || d[WINDOW_MARKER] !== true) return
    if (d.type === WT.TOKEN) send(MSG.TOKEN, d.payload)
    else if (d.type === WT.SNAPSHOT) send(MSG.SNAPSHOT, d.payload)
  } catch {}
})

/* floating action button — Material 3 tonal FAB, bottom-right, shadow-DOM isolated */
function injectFab() {
  if (PAB === false) return
  if (document.querySelector('mad3-fab')) return
  const host = document.createElement('mad3-fab')
  document.documentElement.appendChild(host)
  const shadow = host.attachShadow({ mode: 'open' })
  const style = document.createElement('style')
  style.textContent = `
    :host { all: initial; }
    .fab {
      position: fixed; z-index: 2147483646; right: 18px; bottom: 18px;
      width: 44px; height: 44px; border-radius: 14px 14px 14px 6px;
      background: #7c3aed; color: #fff; border: 0; cursor: pointer;
      box-shadow: 0 4px 8px 3px rgba(0,0,0,.15), 0 1px 3px rgba(0,0,0,.3);
      display: flex; align-items: center; justify-content: center;
      font: 600 15px/1 system-ui, sans-serif; letter-spacing: .1px;
      transition: transform .18s cubic-bezier(.2,0,0,1), box-shadow .18s;
      user-select: none; -webkit-user-select: none;
    }
    .fab:hover { box-shadow: 0 6px 10px 4px rgba(0,0,0,.15), 0 2px 3px rgba(0,0,0,.3); }
    .fab:active { transform: scale(.96); }
    .tip {
      position: fixed; right: 70px; bottom: 22px; z-index: 2147483646;
      background: #322f35; color: #f5eff7; border-radius: 6px; padding: 5px 8px;
      font: 14px/1.2 system-ui, sans-serif; white-space: nowrap; pointer-events: none;
      opacity: 0; transition: opacity .15s;
    }
    .fab:hover + .tip { opacity: 1; }
  `
  const btn = document.createElement('button')
  btn.className = 'fab'
  btn.textContent = 'A+'
  btn.title = 'Assignments+ (Material UI for Teams)'
  const tip = document.createElement('span')
  tip.className = 'tip'
  tip.textContent = 'Assignments+ dashboard'
  btn.addEventListener('click', () => send(MSG.OPEN_DASHBOARD, {}))
  shadow.appendChild(style)
  shadow.appendChild(btn)
  shadow.appendChild(tip)
}

if (document.readyState === 'loading') {
  document.addEventListener('DOMContentLoaded', injectFab, { once: true })
} else {
  injectFab()
}
