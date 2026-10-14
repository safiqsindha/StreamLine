// View resolution: canonical schedule + presentation + view state → ordered rows.
// Every view derives from the same canonical data through this one function.

import { dayNum } from './dates'
import type { Criticality, Presentation, Schedule, Task } from './model'
import { createTaskFilter } from './filter'
import { S } from './strings'
import { compareTasks, displayGroups, groupKeyOf } from './order'
import { buildDependencyIndex } from './reachability'

export { UNGROUPED_LABEL, compareTasks, displayGroups } from './order'
export { downstreamOf, predecessorsOf, successorsOf } from './reachability'

export const MILESTONES_LABEL = S.workstreams.milestones

export type ViewId = 'executive' | 'detailed' | 'dependencies' | 'workstreams' | 'critical'

export const VIEW_IDS: ViewId[] = ['executive', 'detailed', 'dependencies', 'workstreams', 'critical']

export interface Filters {
  /** Display group keys to include. Empty or absent = all. */
  groups?: string[]
  /** Case-insensitive substring over source name and display label. */
  text?: string
  criticality?: Criticality[]
  /** Keep tasks whose start..finish overlaps this window. Tasks without dates are excluded when set. */
  overlap?: { start: string; finish: string }
}

export interface ViewState {
  view: ViewId
  filters?: Filters
  /** Display group keys that are collapsed. */
  collapsed?: string[]
  /** Workstreams drilldown: show exactly this display group. Unknown key → empty result, never widened. */
  focusGroup?: string
  /** 1-based. */
  page?: number
  pageSize?: number
}

export const DEFAULT_PAGE_SIZE = 20

export type Row =
  | {
      kind: 'group'
      key: string
      label: string
      /** Tasks in this group after view + filter selection. */
      count: number
      collapsed: boolean
      start: string | null
      finish: string | null
    }
  | {
      kind: 'task'
      taskId: string
      groupKey: string
      /** Outline depth from source parentId, for indentation. */
      depth: number
    }
  | {
      /** Executive roll-up of one display group. */
      kind: 'summary'
      key: string
      label: string
      start: string | null
      finish: string | null
      /** 'critical' if any source-critical task; 'not-critical' if all report not-critical; else 'unavailable'. */
      critical: Criticality
    }
  | {
      kind: 'milestones'
      label: string
      taskIds: string[]
    }

export interface HiddenDependency {
  /** Visible task whose related task is not on screen. */
  taskId: string
  direction: 'predecessor' | 'successor'
  count: number
  reason: 'filtered' | 'collapsed' | 'other-page' | 'view'
}

export interface ResolvedView {
  view: ViewId
  rows: Row[]
  page: number
  pageCount: number
  pageSize: number
  reductions: {
    /** Source tasks the view's selection policy excluded. */
    byView: number
    byFilters: number
    onOtherPages: number
    inCollapsedGroups: number
  }
  hiddenDeps: HiddenDependency[]
  /** Set when the view cannot be shown honestly; the UI renders an explanation instead of rows. */
  emptyReason?: 'no-criticality-data' | 'no-matching-tasks' | 'unknown-group' | 'no-tasks' | 'no-critical-tasks'
  /** Display groups available in this schedule, in display order. */
  groups: Array<{ key: string; label: string }>
}

type Bucket = 'view' | 'filtered' | 'collapsed' | 'other-page' | 'visible'

interface Span {
  start: string | null
  finish: string | null
}

function spanOf(tasks: Task[]): Span {
  let start: string | null = null
  let finish: string | null = null
  for (const t of tasks) {
    if (t.start && (start === null || dayNum(t.start) < dayNum(start))) start = t.start
    if (t.finish && (finish === null || dayNum(t.finish) > dayNum(finish))) finish = t.finish
  }
  return { start, finish }
}

function aggregateCritical(tasks: Task[]): Criticality {
  if (tasks.some((t) => t.critical === 'critical')) return 'critical'
  if (tasks.length > 0 && tasks.every((t) => t.critical === 'not-critical')) return 'not-critical'
  return 'unavailable'
}

function depthOf(byId: Map<string, Task>, task: Task): number {
  let depth = 0
  const seen = new Set<string>([task.id])
  let cur = task.parentId ? byId.get(task.parentId) : undefined
  while (cur && !seen.has(cur.id)) {
    depth++
    seen.add(cur.id)
    cur = cur.parentId ? byId.get(cur.parentId) : undefined
  }
  return depth
}

/**
 * Bucket accounting: every task lands in exactly one bucket, assigned in the order
 * view, filters, collapsed, pages; whatever remains is visible on the returned page.
 * In the executive view, milestones are the only individually shown tasks, so every
 * other task counts as byView (it is represented by its group summary), and a
 * milestone removed by filters counts as byFilters.
 */
export function resolveView(schedule: Schedule, presentation: Presentation, state: ViewState): ResolvedView {
  const { view } = state
  const groups = displayGroups(schedule, presentation)
  const compare = compareTasks(schedule, presentation)
  const byId = new Map(schedule.tasks.map((t) => [t.id, t]))
  const passes = createTaskFilter(presentation, state.filters)
  const collapsedKeys = new Set(state.collapsed ?? [])
  const pageSize = Math.max(1, Math.floor(state.pageSize ?? DEFAULT_PAGE_SIZE) || DEFAULT_PAGE_SIZE)

  const bucket = new Map<string, Bucket>()
  const rows: Row[] = []
  let page = 1
  let pageCount = 1
  let emptyReason: ResolvedView['emptyReason']

  const finish = (): ResolvedView => {
    const reductions = { byView: 0, byFilters: 0, onOtherPages: 0, inCollapsedGroups: 0 }
    for (const b of bucket.values()) {
      if (b === 'view') reductions.byView++
      else if (b === 'filtered') reductions.byFilters++
      else if (b === 'collapsed') reductions.inCollapsedGroups++
      else if (b === 'other-page') reductions.onOtherPages++
    }
    const hiddenDeps: HiddenDependency[] = []
    const index = buildDependencyIndex(schedule)
    const shownIds = schedule.tasks.filter((t) => bucket.get(t.id) === 'visible').sort(compare)
    const tally = (ids: string[]) => {
      const counts = new Map<string, number>()
      for (const id of ids) {
        const b = bucket.get(id)
        if (b && b !== 'visible') counts.set(b, (counts.get(b) ?? 0) + 1)
      }
      return counts
    }
    const reasons: Array<HiddenDependency['reason']> = ['view', 'filtered', 'collapsed', 'other-page']
    for (const t of shownIds) {
      for (const [direction, ids] of [
        ['predecessor', index.predecessors.get(t.id) ?? []],
        ['successor', index.successors.get(t.id) ?? []],
      ] as const) {
        const counts = tally(ids)
        for (const reason of reasons) {
          const count = counts.get(reason) ?? 0
          if (count > 0) hiddenDeps.push({ taskId: t.id, direction, count, reason })
        }
      }
    }
    return { view, rows, page, pageCount, pageSize, reductions, hiddenDeps, emptyReason, groups: groups.map((g) => ({ ...g })) }
  }

  if (schedule.tasks.length === 0) {
    emptyReason = 'no-tasks'
    return finish()
  }

  const sorted = [...schedule.tasks].sort(compare)

  if (view === 'executive') {
    const passing = sorted.filter(passes)
    const milestones = passing.filter((t) => t.milestone)
    for (const t of sorted) {
      if (!t.milestone) bucket.set(t.id, 'view')
      else bucket.set(t.id, passes(t) ? 'visible' : 'filtered')
    }
    if (milestones.length) rows.push({ kind: 'milestones', label: MILESTONES_LABEL, taskIds: milestones.map((t) => t.id) })
    for (const g of groups) {
      const members = passing.filter((t) => groupKeyOf(t, presentation) === g.key)
      if (!members.length) continue
      rows.push({ kind: 'summary', key: g.key, label: g.label, ...spanOf(members), critical: aggregateCritical(members) })
    }
    if (passing.length === 0) emptyReason = 'no-matching-tasks'
    return finish()
  }

  let selected: Task[]
  if (view === 'critical') {
    if (schedule.tasks.every((t) => t.critical === 'unavailable')) emptyReason = 'no-criticality-data'
    selected = sorted.filter((t) => t.critical === 'critical')
    if (!emptyReason && selected.length === 0) emptyReason = 'no-critical-tasks'
  } else if (view === 'workstreams' && state.focusGroup !== undefined) {
    const focus = state.focusGroup
    selected = sorted.filter((t) => groupKeyOf(t, presentation) === focus)
    if (!groups.some((g) => g.key === focus)) emptyReason = 'unknown-group'
  } else {
    selected = sorted
  }
  const selectedIds = new Set(selected.map((t) => t.id))
  for (const t of sorted) if (!selectedIds.has(t.id)) bucket.set(t.id, 'view')

  const filtered = selected.filter(passes)
  const filteredIds = new Set(filtered.map((t) => t.id))
  for (const t of selected) if (!filteredIds.has(t.id)) bucket.set(t.id, 'filtered')
  if (!emptyReason && selected.length > 0 && filtered.length === 0) emptyReason = 'no-matching-tasks'

  if (emptyReason === 'no-criticality-data' || emptyReason === 'unknown-group') return finish()

  const focused = view === 'workstreams' ? state.focusGroup : undefined

  if (view === 'critical') {
    pageCount = Math.max(1, Math.ceil(filtered.length / pageSize))
    page = clampPage(state.page, pageCount)
    const from = (page - 1) * pageSize
    filtered.forEach((t, i) => {
      const onPage = i >= from && i < from + pageSize
      bucket.set(t.id, onPage ? 'visible' : 'other-page')
      if (onPage) rows.push({ kind: 'task', taskId: t.id, groupKey: groupKeyOf(t, presentation), depth: depthOf(byId, t) })
    })
    return finish()
  }

  const isCollapsed = (key: string) => key !== focused && collapsedKeys.has(key)
  const members = groups
    .map((g) => ({ g, tasks: filtered.filter((t) => groupKeyOf(t, presentation) === g.key) }))
    .filter((m) => m.tasks.length > 0)

  let paged = 0
  for (const m of members) {
    if (isCollapsed(m.g.key)) {
      for (const t of m.tasks) bucket.set(t.id, 'collapsed')
    } else paged += m.tasks.length
  }
  pageCount = Math.max(1, Math.ceil(paged / pageSize))
  page = clampPage(state.page, pageCount)

  let n = 0
  for (const m of members) {
    const closed = isCollapsed(m.g.key)
    const header = (): Row => ({ kind: 'group', key: m.g.key, label: m.g.label, count: m.tasks.length, collapsed: closed, ...spanOf(m.tasks) })
    if (closed) {
      if (Math.min(pageCount, Math.floor(n / pageSize) + 1) === page) rows.push(header())
      continue
    }
    const onPage: Task[] = []
    m.tasks.forEach((t, i) => {
      const pos = n + i
      const here = pos >= (page - 1) * pageSize && pos < page * pageSize
      bucket.set(t.id, here ? 'visible' : 'other-page')
      if (here) onPage.push(t)
    })
    n += m.tasks.length
    if (!onPage.length) continue
    rows.push(header())
    for (const t of onPage) rows.push({ kind: 'task', taskId: t.id, groupKey: m.g.key, depth: depthOf(byId, t) })
  }
  return finish()
}

function clampPage(page: number | undefined, pageCount: number): number {
  const p = Math.floor(page ?? 1)
  if (!Number.isFinite(p)) return 1
  return Math.min(Math.max(p, 1), pageCount)
}
