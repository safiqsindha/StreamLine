// Generated IDs are random and never reused. Task IDs come from the source
// file; see `fallbackTaskId` for files that do not carry one.

export function newId(prefix: 'sch' | 'bl' | 'imp'): string {
  return `${prefix}_${globalThis.crypto.randomUUID().replace(/-/g, '')}`
}

/**
 * Deterministic fallback task ID for sources without an ID column:
 * a hash of the workstream and name path, plus an occurrence counter for exact duplicates.
 * Never derived from row position or display label.
 */
export function fallbackTaskId(parts: Array<string | null>, occurrence: number): string {
  const key = parts.map((p) => (p ?? '').trim().toLowerCase()).join('␟')
  let h = 0x811c9dc5
  for (let i = 0; i < key.length; i++) {
    h ^= key.charCodeAt(i)
    h = Math.imul(h, 0x01000193)
  }
  const base = `h${(h >>> 0).toString(16).padStart(8, '0')}`
  return occurrence > 0 ? `${base}-${occurrence + 1}` : base
}
