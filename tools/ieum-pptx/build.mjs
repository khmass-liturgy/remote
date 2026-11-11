import { build } from 'esbuild'
import path from 'node:path'
const SRC = process.env.IEUM_SRC, here = process.cwd()
const r = await build({
  entryPoints: ['entry.tsx'], bundle: true, format: 'iife', globalName: 'IeumPptx', platform: 'browser',
  target: 'es2022', minify: true, outfile: 'ieum-pptx.js', metafile: true, legalComments: 'none',
  jsx: 'automatic', nodePaths: [path.join(here, 'node_modules')],
  alias: { 'node:crypto': './shim-crypto.js', 'node:zlib': './shim-crypto.js', 'node:fs': './shim-crypto.js', 'node:stream/promises': './shim-crypto.js', 'node:path': './shim-crypto.js', 'node:os': './shim-crypto.js', '@ieum-slides': path.join(SRC, 'apps/slides/src') },
  inject: ['./inject-buffer.js'],
  plugins: [{ name: 'no-ui-i18n', setup(b) { b.onResolve({ filter: /\/i18n\/locale$/ }, () => ({ path: path.join(here, 'shim-i18n.js') })) } }],
  define: { 'process.env.NODE_ENV': '"production"', global: 'globalThis' },
  loader: { '.png': 'dataurl', '.svg': 'dataurl', '.woff2': 'dataurl', '.ttf': 'dataurl' },
  logLimit: 40,
})
const ins = Object.entries(r.metafile.inputs).map(([k, v]) => [v.bytes, k]).sort((a, b) => b[0] - a[0]).slice(0, 15)
console.log(ins.map(([b, k]) => (b / 1024 | 0) + 'K ' + k.replace(SRC, '')).join('\n'))
