import type { SourceKind, Task, Workstream } from '../model'

/** A field the canonical model needs, mapped from a source column. */
export type MappedField =
  | 'id'
  | 'name'
  | 'workstream'
  | 'owner'
  | 'start'
  | 'finish'
  | 'milestone'
  | 'critical'
  | 'float'
  | 'parent'
  | 'summary'
  | 'predecessors'

/** Column mapping for spreadsheet sources: field → source column header (exact text). */
export type ColumnMapping = Partial<Record<MappedField, string>>

export interface ImportIssue {
  /** 1-based source row, when the issue belongs to a row. */
  row?: number
  /** Plain-language, actionable message. Shown to the user verbatim. */
  message: string
  severity: 'skipped' | 'warning'
}

/** The parsed, validated result. Nothing is persisted until it is accepted. */
export interface ParsedSchedule {
  kind: SourceKind
  filename: string
  name: string
  workstreams: Workstream[]
  tasks: Task[]
  /** Columns found in the file (spreadsheets) or fields found (MSPDI). */
  detectedColumns: string[]
  /** Mapping actually used. For MSPDI this is fixed by the schema. */
  mapping: ColumnMapping
  /** Source columns or fields that could not be mapped to anything. */
  unmapped: string[]
  issues: ImportIssue[]
}

export interface ImportPreview {
  previewId: string
  parsed: ParsedSchedule
  summary: {
    rowCount: number
    laneCount: number
    dateRange: { start: string; finish: string } | null
    hasCriticality: boolean
    hasFloat: boolean
  }
  /** Present when the import replaces an existing schedule. */
  reconciliation?: Reconciliation
}

export interface Reconciliation {
  targetScheduleId: string
  added: string[]
  removed: string[]
  changed: Array<{ id: string; fields: string[] }>
  /** Presentation overrides whose task no longer exists. Retained, not dropped. */
  orphanedOverrides: string[]
}

/** Raised when the file cannot be imported at all. The message is shown to the user verbatim. */
export class ImportError extends Error {
  constructor(message: string) {
    super(message)
    this.name = 'ImportError'
  }
}
