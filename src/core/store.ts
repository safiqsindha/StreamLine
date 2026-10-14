import type { Baseline, Presentation, Schedule, ScheduleMeta } from './model'

/**
 * The single persistence boundary. Nothing else reads or writes storage.
 * Swapping the backing store means writing one new class that implements this.
 */
export interface Store {
  listSchedules(): Promise<ScheduleMeta[]>
  getSchedule(id: string): Promise<Schedule | null>
  putSchedule(s: Schedule): Promise<void>
  deleteSchedule(id: string): Promise<void>
  getPresentation(scheduleId: string): Promise<Presentation | null>
  putPresentation(scheduleId: string, p: Presentation): Promise<void>
  listBaselines(scheduleId: string): Promise<Baseline[]>
  /** Creates a new baseline record. Rejects if a baseline with this ID already exists. */
  putBaseline(scheduleId: string, b: Baseline): Promise<void>
  deleteBaseline(scheduleId: string, baselineId: string): Promise<void>
}

/** Raised for storage failures so callers can show a visible message instead of throwing raw errors. */
export class StoreError extends Error {
  constructor(message: string, readonly cause?: unknown) {
    super(message)
    this.name = 'StoreError'
  }
}
