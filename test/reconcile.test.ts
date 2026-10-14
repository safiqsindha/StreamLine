import { describe, expect, it } from 'vitest'
import { reconcile } from '../src/core/import/reconcile'
import type { ParsedSchedule } from '../src/core/import/types'
import type { Presentation, Schedule, Task } from '../src/core/model'

function task(id: string, over: Partial<Task> = {}): Task {
  return {
    id,
    name: `Task ${id}`,
    workstream: 'Avionics',
    owner: null,
    start: '2026-01-05',
    finish: '2026-01-09',
    milestone: false,
    critical: 'unavailable',
    floatDays: null,
    parentId: null,
    isSummary: false,
    deps: [],
    ...over,
  }
}

function schedule(tasks: Task[]): Schedule {
  return {
    id: 'sch_1',
    name: 'Program',
    createdAt: '2026-01-01T00:00:00.000Z',
    source: { kind: 'mspdi', filename: 'a.xml', importedAt: '2026-01-01T00:00:00.000Z' },
    workstreams: [],
    tasks,
  }
}

function parsed(tasks: Task[]): ParsedSchedule {
  return { kind: 'mspdi', filename: 'b.xml', name: 'Program', workstreams: [], tasks, detectedColumns: [], mapping: {}, unmapped: [], issues: [] }
}

describe('reconcile', () => {
  it('reports no differences for identical content', () => {
    const t = [task('1', { deps: [{ predecessorId: '2', type: 'FS', lagDays: 0 }] }), task('2')]
    const r = reconcile(schedule(t), parsed(structuredClone(t)), null)
    expect(r).toEqual({ targetScheduleId: 'sch_1', added: [], removed: [], changed: [], orphanedOverrides: [] })
  })

  it('reports added and removed tasks by ID', () => {
    const r = reconcile(schedule([task('1'), task('2')]), parsed([task('2'), task('3')]), null)
    expect(r.added).toEqual(['3'])
    expect(r.removed).toEqual(['1'])
  })

  it('lists which source fields changed', () => {
    const old = task('1')
    const next = task('1', { name: 'Renamed', finish: '2026-02-01', critical: 'critical', floatDays: 0, owner: 'Sam', milestone: true })
    const r = reconcile(schedule([old]), parsed([next]), null)
    expect(r.changed).toEqual([{ id: '1', fields: ['name', 'owner', 'finish', 'milestone', 'critical', 'floatDays'] }])
  })

  it('detects workstream, parent and summary changes', () => {
    const r = reconcile(schedule([task('1')]), parsed([task('1', { workstream: 'Power', parentId: '9', isSummary: true })]), null)
    expect(r.changed[0].fields).toEqual(['workstream', 'parentId', 'isSummary'])
  })

  it('treats dependency order as irrelevant but lag and type as changes', () => {
    const a = { predecessorId: 'a', type: 'FS' as const, lagDays: 0 }
    const b = { predecessorId: 'b', type: 'SS' as const, lagDays: 2 }
    const same = reconcile(schedule([task('1', { deps: [a, b] })]), parsed([task('1', { deps: [b, a] })]), null)
    expect(same.changed).toEqual([])
    const lag = reconcile(schedule([task('1', { deps: [a] })]), parsed([task('1', { deps: [{ ...a, lagDays: 3 }] })]), null)
    expect(lag.changed).toEqual([{ id: '1', fields: ['deps'] }])
    const nullLag = reconcile(schedule([task('1', { deps: [a] })]), parsed([task('1', { deps: [{ ...a, lagDays: null }] })]), null)
    expect(nullLag.changed).toEqual([{ id: '1', fields: ['deps'] }])
  })

  it('reports overrides for tasks missing from the incoming file as orphaned', () => {
    const p: Presentation = { scheduleId: 'sch_1', updatedAt: 'x', tasks: { '1': { emphasis: 'muted' }, '2': { displayLabel: 'Two' } } }
    const r = reconcile(schedule([task('1'), task('2')]), parsed([task('2')]), p)
    expect(r.orphanedOverrides).toEqual(['1'])
  })

  it('does not mutate its inputs', () => {
    const s = schedule([task('1')])
    const i = parsed([task('1', { name: 'x' })])
    const snap = JSON.stringify([s, i])
    reconcile(s, i, null)
    expect(JSON.stringify([s, i])).toBe(snap)
  })
})
