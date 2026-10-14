// Reachability over stated dependency links. This is graph traversal, not CPM:
// nothing here computes dates, float or criticality.

import type { Schedule } from './model'

export interface DependencyIndex {
  predecessors: Map<string, string[]>
  successors: Map<string, string[]>
}

/** Links whose predecessor is not a task in the schedule are ignored. Each pair is listed once. */
export function buildDependencyIndex(schedule: Schedule): DependencyIndex {
  const ids = new Set(schedule.tasks.map((t) => t.id))
  const predecessors = new Map<string, string[]>()
  const successors = new Map<string, string[]>()
  for (const id of ids) {
    predecessors.set(id, [])
    successors.set(id, [])
  }
  for (const t of schedule.tasks) {
    const seen = new Set<string>()
    for (const d of t.deps) {
      if (!ids.has(d.predecessorId) || seen.has(d.predecessorId)) continue
      seen.add(d.predecessorId)
      predecessors.get(t.id)!.push(d.predecessorId)
      successors.get(d.predecessorId)!.push(t.id)
    }
  }
  return { predecessors, successors }
}

export function predecessorsOf(schedule: Schedule, taskId: string): string[] {
  return buildDependencyIndex(schedule).predecessors.get(taskId) ?? []
}

export function successorsOf(schedule: Schedule, taskId: string): string[] {
  return buildDependencyIndex(schedule).successors.get(taskId) ?? []
}

/**
 * Every task reachable by following successor links from `taskId`, transitively.
 * Cycle-safe. The starting task is never included in its own result.
 */
export function downstreamOf(schedule: Schedule, taskId: string): Set<string> {
  const { successors } = buildDependencyIndex(schedule)
  const out = new Set<string>()
  const stack = [taskId]
  while (stack.length) {
    const cur = stack.pop()!
    for (const next of successors.get(cur) ?? []) {
      if (next === taskId || out.has(next)) continue
      out.add(next)
      stack.push(next)
    }
  }
  return out
}
