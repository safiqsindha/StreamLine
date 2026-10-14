import type { Dependency, Presentation, Schedule, Task } from '../model'
import type { ParsedSchedule, Reconciliation } from './types'

const SCALAR_FIELDS = ['name', 'workstream', 'owner', 'start', 'finish', 'milestone', 'critical', 'floatDays', 'parentId', 'isSummary'] as const

function depKey(d: Dependency): string {
  return `${d.predecessorId}|${d.type}|${d.lagDays ?? ''}`
}

function depsEqual(a: Dependency[], b: Dependency[]): boolean {
  if (a.length !== b.length) return false
  const left = a.map(depKey).sort()
  const right = b.map(depKey).sort()
  return left.every((k, i) => k === right[i])
}

function changedFields(a: Task, b: Task): string[] {
  const fields: string[] = SCALAR_FIELDS.filter((f) => a[f] !== b[f])
  if (!depsEqual(a.deps, b.deps)) fields.push('deps')
  return fields
}

/** Compare an existing schedule with a re-import, by task ID. Pure. */
export function reconcile(existing: Schedule, incoming: ParsedSchedule, presentation: Presentation | null): Reconciliation {
  const before = new Map(existing.tasks.map((t) => [t.id, t]))
  const after = new Map(incoming.tasks.map((t) => [t.id, t]))

  const added: string[] = []
  const changed: Reconciliation['changed'] = []
  for (const [id, task] of after) {
    const old = before.get(id)
    if (!old) {
      added.push(id)
      continue
    }
    const fields = changedFields(old, task)
    if (fields.length > 0) changed.push({ id, fields })
  }
  const removed = [...before.keys()].filter((id) => !after.has(id))
  const orphanedOverrides = presentation ? Object.keys(presentation.tasks).filter((id) => !after.has(id)) : []

  return { targetScheduleId: existing.id, added, removed, changed, orphanedOverrides }
}
