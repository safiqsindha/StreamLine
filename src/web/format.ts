import { fmtLong, fmtShort } from '../core/dates'
import type { Presentation, Task } from '../core/model'
import { S } from '../core/strings'

/** Shown name: presentation label when set, else the source name. */
export function displayName(task: Task, presentation: Presentation | null): string {
  return presentation?.tasks[task.id]?.displayLabel || task.name
}

export function hasLabelOverride(task: Task, presentation: Presentation | null): boolean {
  const label = presentation?.tasks[task.id]?.displayLabel
  return !!label && label !== task.name
}

export const shortDate = (iso: string | null): string => (iso ? fmtShort(iso) : S.unavailable)

export const longDateTime = (iso: string): string => fmtLong(iso)

export function laneToken(index: number): string {
  return `var(--lane-${(index % 6) + 1})`
}

export function slug(s: string): string {
  return s.toLowerCase().replace(/[^a-z0-9]+/g, '-').replace(/^-|-$/g, '') || 'schedule'
}

export function localTodayISO(): string {
  const d = new Date()
  return `${d.getFullYear()}-${String(d.getMonth() + 1).padStart(2, '0')}-${String(d.getDate()).padStart(2, '0')}`
}

export function download(name: string, blob: Blob): void {
  const url = URL.createObjectURL(blob)
  const a = document.createElement('a')
  a.href = url
  a.download = name
  document.body.appendChild(a)
  a.click()
  a.remove()
  setTimeout(() => URL.revokeObjectURL(url), 1500)
}
