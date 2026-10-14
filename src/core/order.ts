// Display groups and the one task ordering comparator.

import { S } from './strings'
import type { Presentation, Schedule, Task } from './model'

export const UNGROUPED_LABEL = S.workstreams.ungrouped
export const UNGROUPED_KEY = ''

export interface DisplayGroup {
  key: string
  label: string
}

export function groupKeyOf(task: Task, presentation: Presentation): string {
  return presentation.tasks[task.id]?.displayGroup ?? task.workstream ?? UNGROUPED_KEY
}

function sourceLabel(schedule: Schedule, key: string): string {
  if (key === UNGROUPED_KEY) return UNGROUPED_LABEL
  const ws = schedule.workstreams.find((w) => w.id === key)
  return ws ? ws.name : key
}

/**
 * Display groups in display order: presentation.groupOrder first (those that exist),
 * then remaining groups in source workstream order, then other display groups
 * alphabetically, with the ungrouped group last.
 */
export function displayGroups(schedule: Schedule, presentation: Presentation): DisplayGroup[] {
  const present = new Map<string, DisplayGroup>()
  for (const t of schedule.tasks) {
    const key = groupKeyOf(t, presentation)
    if (!present.has(key)) present.set(key, { key, label: sourceLabel(schedule, key) })
  }
  const out: DisplayGroup[] = []
  const take = (g: DisplayGroup | undefined) => {
    if (g && present.delete(g.key)) out.push(g)
  }
  const byKeyOrLabel = (name: string) => {
    const direct = present.get(name)
    if (direct) return direct
    for (const g of present.values()) if (g.label === name && g.key !== UNGROUPED_KEY) return g
    return undefined
  }
  for (const name of presentation.groupOrder ?? []) take(byKeyOrLabel(name))
  for (const ws of schedule.workstreams) take(present.get(ws.id) ?? byKeyOrLabel(ws.name))
  const ungrouped = present.get(UNGROUPED_KEY)
  present.delete(UNGROUPED_KEY)
  const rest = [...present.values()].sort((a, b) => cmp(a.label, b.label) || cmp(a.key, b.key))
  out.push(...rest)
  if (ungrouped) out.push(ungrouped)
  return out
}

function cmp(a: string, b: string): number {
  return a < b ? -1 : a > b ? 1 : 0
}

/**
 * The single ordering used by every view, pagination and keyboard focus.
 * Order of precedence: display group position, presentation `order` (defined before
 * undefined, lower first), source file order (index in schedule.tasks, which keeps
 * children after their parent as the file lists them), then task ID.
 * Start dates are deliberately not used: PMs expect the order of their file.
 */
export function compareTasks(schedule: Schedule, presentation: Presentation): (a: Task, b: Task) => number {
  const groupRank = new Map(displayGroups(schedule, presentation).map((g, i) => [g.key, i]))
  const fileIndex = new Map(schedule.tasks.map((t, i) => [t.id, i]))
  return (a, b) => {
    const ga = groupRank.get(groupKeyOf(a, presentation)) ?? Infinity
    const gb = groupRank.get(groupKeyOf(b, presentation)) ?? Infinity
    if (ga !== gb) return ga - gb
    const oa = presentation.tasks[a.id]?.order
    const ob = presentation.tasks[b.id]?.order
    if (oa !== ob) {
      if (oa === undefined) return 1
      if (ob === undefined) return -1
      if (oa !== ob) return oa - ob
    }
    const ia = fileIndex.get(a.id) ?? 0
    const ib = fileIndex.get(b.id) ?? 0
    if (ia !== ib) return ia - ib
    return cmp(a.id, b.id)
  }
}
