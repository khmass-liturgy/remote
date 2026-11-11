// Browser stand-ins for the few node:crypto / node:zlib calls reached at import time.
export function randomUUID() { return crypto.randomUUID() }
export function createHash() {
  const parts = []
  return {
    update(b) { parts.push(b); return this },
    digest() { let h = 2166136261; for (const p of parts) for (let i = 0; i < p.length; i += 97) h = Math.imul(h ^ p[i], 16777619); return (h >>> 0).toString(16) + ':' + parts.reduce((n, p) => n + p.length, 0) },
  }
}
export function deflateSync() { throw new Error('not available in the browser viewer') }
export default { randomUUID, createHash, deflateSync }
