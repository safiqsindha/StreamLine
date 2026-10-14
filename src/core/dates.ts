// Calendar dates are handled as UTC day numbers so no time zone can shift them.

const DAY_MS = 86_400_000

export function dayNum(iso: string): number {
  return Math.round(Date.UTC(+iso.slice(0, 4), +iso.slice(5, 7) - 1, +iso.slice(8, 10)) / DAY_MS)
}

export function isoOf(day: number): string {
  return new Date(day * DAY_MS).toISOString().slice(0, 10)
}

export function addDays(iso: string, n: number): string {
  return isoOf(dayNum(iso) + n)
}

export function isIsoDate(v: unknown): v is string {
  if (typeof v !== 'string' || !/^\d{4}-\d{2}-\d{2}$/.test(v)) return false
  return isoOf(dayNum(v)) === v
}

export function weekday(day: number): number {
  return new Date(day * DAY_MS).getUTCDay()
}

export function fmtShort(iso: string): string {
  return new Date(dayNum(iso) * DAY_MS).toLocaleDateString('en-US', { month: 'short', day: 'numeric', timeZone: 'UTC' })
}

export function fmtLong(iso: string): string {
  return new Date(iso).toLocaleString('en-GB', {
    day: 'numeric',
    month: 'short',
    year: 'numeric',
    hour: '2-digit',
    minute: '2-digit',
  })
}

/** Calendar duration in days, inclusive. Null when either date is missing. */
export function spanDays(start: string | null, finish: string | null): number | null {
  if (!start || !finish) return null
  return dayNum(finish) - dayNum(start) + 1
}
