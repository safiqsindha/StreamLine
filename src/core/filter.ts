// Task filters. They choose which tasks are shown and never alter a date.

import { dayNum } from './dates'
import type { Presentation, Task } from './model'
import { groupKeyOf } from './order'
import type { Filters } from './views'

/**
 * Overlap requires both start and finish; a task missing either is excluded when
 * an overlap window is set, rather than guessing from the one date it has.
 */
export function createTaskFilter(presentation: Presentation, filters: Filters | undefined): (t: Task) => boolean {
  const f = filters ?? {}
  const groups = f.groups && f.groups.length ? new Set(f.groups) : null
  const text = f.text?.trim().toLowerCase() || null
  const crit = f.criticality && f.criticality.length ? new Set(f.criticality) : null
  const win = f.overlap ? { start: dayNum(f.overlap.start), finish: dayNum(f.overlap.finish) } : null
  return (t) => {
    if (groups && !groups.has(groupKeyOf(t, presentation))) return false
    if (crit && !crit.has(t.critical)) return false
    if (text) {
      const label = presentation.tasks[t.id]?.displayLabel ?? ''
      if (!t.name.toLowerCase().includes(text) && !label.toLowerCase().includes(text)) return false
    }
    if (win) {
      if (!t.start || !t.finish) return false
      if (dayNum(t.start) > win.finish || dayNum(t.finish) < win.start) return false
    }
    return true
  }
}
