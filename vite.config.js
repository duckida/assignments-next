import { defineConfig } from 'vite'
import react from '@vitejs/plugin-react'

// Builds two HTML pages into dist/:
//   index.html   -> the extension dashboard page (chrome.tabs.create)
//   preview.html -> standalone demo harness (chokes chrome.* away, loads mock data)
export default defineConfig({
  plugins: [react()],
  base: './',
  build: {
    outDir: 'dist',
    emptyOutDir: false, // build-extension.mjs wipes dist/ and stages non-page assets
    rollupOptions: {
      input: {
        index: 'index.html',
        preview: 'preview.html',
      },
    },
  },
})
