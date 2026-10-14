// View resolution: canonical schedule + presentation + view state → ordered rows.
// Every view derives from the same canonical data through this one function.

import type { Criticality, Presentation, Schedule } from './model'

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
  emptyReason?: 'no-criticality-data' | 'no-matching-tasks' | 'unknown-group' | 'no-tasks'
  /** Display groups available in this schedule, in display order. */
  groups: Array<{ key: string; label: string }>
}

export function resolveView(_schedule: Schedule, _presentation: Presentation, _state: ViewState): ResolvedView {
  throw new Error('not implemented')
}
