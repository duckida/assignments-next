#!/usr/bin/env node
/** Extension build: esbuild the three non-ESM contexts (SW, bridge, MAIN capture),
 * then Vite-build the React dashboard (+ demo preview page) into the same dist/. */
import { build } from 'esbuild'
import { execSync } from 'node:child_process'
import { cpSync, mkdirSync, readFileSync, rmSync, writeFileSync } from 'node:fs'
import { resolve, dirname } from 'node:path'
import { fileURLToPath } from 'node:url'

const root = dirname(fileURLToPath(import.meta.url))
const dist = resolve(root, '..', 'dist')
const shared = {
  bundle: true,
  format: 'iife',
  target: ['chrome111'],
  minify: false,
  legalComments: 'none',
  sourcemap: false,
  logLevel: 'info',
}

console.log('[build] cleaning dist/')
rmSync(dist, { recursive: true, force: true })
mkdirSync(dist, { recursive: true })

console.log('[build] esbuild: background / bridge / capture')
await build({ ...shared, entryPoints: ['src/background/index.js'], outfile: 'dist/assets/background.js' })
await build({ ...shared, entryPoints: ['src/content/bridge.js'], outfile: 'dist/assets/bridge.js' })
await build({ ...shared, entryPoints: ['src/inject/capture.js'], outfile: 'dist/assets/capture.js' })

console.log('[build] vite: dashboard + preview')
execSync('npx vite build', { stdio: 'inherit', cwd: resolve(root, '..') })

console.log('[build] staging manifest + icons')
cpSync(resolve(root, '..', 'public', 'manifest.json'), resolve(dist, 'manifest.json'))
cpSync(resolve(root, '..', 'public', 'icons'), resolve(dist, 'icons'), { recursive: true })

// sanity: every js path referenced by the manifest must exist
import { existsSync } from 'node:fs'
const manifest = JSON.parse(readFileSync(resolve(dist, 'manifest.json'), 'utf8'))
const refs = [...(manifest.content_scripts ?? []).flatMap((c) => c.js), manifest.background.service_worker]
for (const ref of refs) {
  if (!existsSync(resolve(dist, ref))) throw new Error('manifest references missing file: ' + ref)
}
console.log('[build] dist ready:', dist)
