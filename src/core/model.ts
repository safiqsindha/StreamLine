// Canonical schedule model. Source facts and presentation overrides live in
// separate types and separate stores; they are never merged into one object.

export type Criticality = 'critical' | 'not-critical' | 'unavailable'

export type DependencyType = 'FS' | 'SS' | 'FF' | 'SF' | 'unknown'

export type SourceKind = 'xlsx' | 'mspdi'

export interface Dependency {
  predecessorId: string
  type: DependencyType
  lagDays: number | null
}

/** A task exactly as the source file states it. `null` means the file did not provide the value. */
export interface Task {
  id: string
  name: string
  workstream: string | null
  owner: string | null
  start: string | null // ISO YYYY-MM-DD
  finish: string | null
  milestone: boolean
  critical: Criticality
  /** Total float in days, when the source provides it. */
  floatDays: number | null
  parentId: string | null
  isSummary: boolean
  deps: Dependency[]
}

export interface Workstream {
  id: string
  name: string
}

export interface ScheduleSource {
  kind: SourceKind
  filename: string
  importedAt: string // ISO 8601
}

export interface Schedule {
  id: string
  name: string
  createdAt: string
  source: ScheduleSource
  workstreams: Workstream[]
  tasks: Task[]
}

export interface ScheduleMeta {
  id: string
  name: string
  createdAt: string
  source: ScheduleSource
  taskCount: number
}

// ---------------------------------------------------------------------------
// Presentation: everything a user may change. Keyed by schedule ID.

export type Emphasis = 'none' | 'highlight' | 'muted'

export interface TaskPresentation {
  displayLabel?: string
  /** Display group name. Overrides the source workstream for grouping only. */
  displayGroup?: string
  /** Explicit sort position within its group. Lower sorts first. */
  order?: number
  emphasis?: Emphasis
}

export interface Presentation {
  scheduleId: string
  updatedAt: string
  /** Display name for the schedule; the source name is kept in Schedule.name. */
  displayName?: string
  tasks: Record<string, TaskPresentation>
  /** Display ordering of groups (by display group name). */
  groupOrder?: string[]
  theme?: string
}

export function emptyPresentation(scheduleId: string): Presentation {
  return { scheduleId, updatedAt: new Date(0).toISOString(), tasks: {} }
}

// ---------------------------------------------------------------------------
// Baselines: immutable snapshots of source dates.

export interface BaselineTask {
  id: string
  name: string
  start: string | null
  finish: string | null
  milestone: boolean
}

export interface Baseline {
  id: string
  scheduleId: string
  name: string
  createdAt: string
  tasks: BaselineTask[]
}

export function toMeta(s: Schedule): ScheduleMeta {
  return { id: s.id, name: s.name, createdAt: s.createdAt, source: s.source, taskCount: s.tasks.length }
}
