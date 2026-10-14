import { describe, expect, it } from 'vitest'
import { emptyPresentation } from '../src/core/model'
import type { Criticality, Dependency, Presentation, Schedule, Task } from '../src/core/model'
import { UNGROUPED_LABEL, compareTasks, downstreamOf, predecessorsOf, resolveView, successorsOf } from '../src/core/views'
import type { Row, ViewState } from '../src/core/views'

function task(id: string, o: Partial<Task> = {}): Task {
  return {
    id,
    name: `Task ${id}`,
    workstream: 'a',
    owner: null,
    start: '2026-01-05',
    finish: '2026-01-09',
    milestone: false,
    critical: 'unavailable',
    floatDays: null,
    parentId: null,
    isSummary: false,
    deps: [],
    ...o,
  }
}

function dep(predecessorId: string): Dependency {
  return { predecessorId, type: 'FS', lagDays: null }
}

function schedule(tasks: Task[], workstreams = [{ id: 'a', name: 'Alpha' }, { id: 'b', name: 'Beta' }]): Schedule {
  return {
    id: 's1',
    name: 'S',
    createdAt: '2026-01-01T00:00:00Z',
    source: { kind: 'xlsx', filename: 's.xlsx', importedAt: '2026-01-01T00:00:00Z' },
    workstreams,
    tasks,
  }
}

const pres = (p: Partial<Presentation> = {}): Presentation => ({ ...emptyPresentation('s1'), ...p })
const taskIds = (rows: Row[]) => rows.flatMap((r) => (r.kind === 'task' ? [r.taskId] : []))
const total = (r: ReturnType<typeof resolveView>) =>
  r.reductions.byView + r.reductions.byFilters + r.reductions.onOtherPages + r.reductions.inCollapsedGroups

describe('ordering', () => {
  it('uses file order, ignoring start dates', () => {
    const s = schedule([task('z', { start: '2026-03-01' }), task('y', { start: '2026-01-01' })])
    const r = resolveView(s, pres(), { view: 'detailed' })
    expect(taskIds(r.rows)).toEqual(['z', 'y'])
  })

  it('applies presentation order before undefined, then file order, then ID', () => {
    const s = schedule([task('c'), task('b'), task('a'), task('d')])
    const p = pres({ tasks: { a: { order: 2 }, d: { order: 1 } } })
    const cmp = compareTasks(s, p)
    expect([...s.tasks].sort(cmp).map((t) => t.id)).toEqual(['d', 'a', 'c', 'b'])
  })

  it('breaks ties by ID when file positions are identical', () => {
    const t = task('x')
    const dup = { ...task('w') }
    const s = schedule([t])
    const cmp = compareTasks(s, pres())
    const same = (a: Task, b: Task) => cmp(a, b)
    expect(same(t, { ...t, id: 'x' })).toBe(0)
    expect(same(dup, t)).not.toBe(0)
  })

  it('moves a task with displayGroup and orders groups', () => {
    const s = schedule([task('a1'), task('a2'), task('b1', { workstream: 'b' })])
    const r = resolveView(s, pres({ tasks: { a2: { displayGroup: 'Custom' } } }), { view: 'detailed' })
    expect(r.groups.map((g) => g.key)).toEqual(['a', 'b', 'Custom'])
    const keys = r.rows.filter((x) => x.kind === 'task').map((x) => (x.kind === 'task' ? x.groupKey : ''))
    expect(keys).toEqual(['a', 'b', 'Custom'])
  })

  it('honours groupOrder, then source order, alphabetical, ungrouped last', () => {
    const s = schedule([
      task('n', { workstream: null }),
      task('z', { workstream: 'zeta' }),
      task('m', { workstream: 'mu' }),
      task('b', { workstream: 'b' }),
      task('a'),
    ])
    const r = resolveView(s, pres({ groupOrder: ['mu', 'missing'] }), { view: 'detailed' })
    expect(r.groups.map((g) => g.key)).toEqual(['mu', 'a', 'b', 'zeta', ''])
    expect(r.groups[4].label).toBe(UNGROUPED_LABEL)
    expect(r.groups[1].label).toBe('Alpha')
  })
})

describe('views', () => {
  const s = schedule([
    task('a1', { critical: 'critical' }),
    task('a2', { critical: 'not-critical' }),
    task('m1', { milestone: true, start: '2026-02-01', finish: '2026-02-01', critical: 'not-critical' }),
    task('b1', { workstream: 'b', critical: 'not-critical' }),
    task('b2', { workstream: 'b', critical: 'not-critical', parentId: 'b1' }),
  ])

  it('executive: milestones row then summaries with tri-state critical', () => {
    const r = resolveView(s, pres(), { view: 'executive' })
    expect(r.rows.map((x) => x.kind)).toEqual(['milestones', 'summary', 'summary'])
    const [ms, a, b] = r.rows
    expect(ms).toMatchObject({ label: 'Key milestones', taskIds: ['m1'] })
    expect(a).toMatchObject({ key: 'a', critical: 'critical', start: '2026-01-05', finish: '2026-02-01' })
    expect(b).toMatchObject({ key: 'b', critical: 'not-critical' })
    expect(r.reductions.byView).toBe(4)
    expect(total(r)).toBe(4)
  })

  it('executive: unavailable when no task reports criticality, and mixed is unavailable', () => {
    const none = schedule([task('x'), task('y')])
    expect(resolveView(none, pres(), { view: 'executive' }).rows[0]).toMatchObject({ critical: 'unavailable' })
    const mixed = schedule([task('x', { critical: 'not-critical' }), task('y')])
    expect(resolveView(mixed, pres(), { view: 'executive' }).rows[0]).toMatchObject({ critical: 'unavailable' })
  })

  it('executive: summary span ignores null dates and includes source summary tasks', () => {
    const sch = schedule([task('p', { isSummary: true, start: '2026-01-01', finish: '2026-04-01' }), task('q', { start: null, finish: null })])
    expect(resolveView(sch, pres(), { view: 'executive' }).rows[0]).toMatchObject({ start: '2026-01-01', finish: '2026-04-01' })
    const nulls = schedule([task('q', { start: null, finish: null })])
    expect(resolveView(nulls, pres(), { view: 'executive' }).rows[0]).toMatchObject({ start: null, finish: null })
  })

  it('detailed and dependencies show group headers and all tasks with depth', () => {
    for (const view of ['detailed', 'dependencies'] as const) {
      const r = resolveView(s, pres(), { view })
      expect(r.rows.map((x) => x.kind)).toEqual(['group', 'task', 'task', 'task', 'group', 'task', 'task'])
      expect(r.rows.find((x) => x.kind === 'task' && x.taskId === 'b2')).toMatchObject({ depth: 1 })
      expect(r.reductions).toEqual({ byView: 0, byFilters: 0, onOtherPages: 0, inCollapsedGroups: 0 })
    }
  })

  it('workstreams focusGroup shows exactly that group, expanded', () => {
    const r = resolveView(s, pres(), { view: 'workstreams', focusGroup: 'b', collapsed: ['b'] })
    expect(r.rows[0]).toMatchObject({ kind: 'group', key: 'b', collapsed: false })
    expect(taskIds(r.rows)).toEqual(['b1', 'b2'])
    expect(r.reductions.byView).toBe(3)
    expect(r.groups).toHaveLength(2)
  })

  it('workstreams unknown focusGroup is empty and never widened', () => {
    const r = resolveView(s, pres(), { view: 'workstreams', focusGroup: 'Beta' })
    expect(r.rows).toEqual([])
    expect(r.emptyReason).toBe('unknown-group')
    expect(total(r)).toBe(5)
  })

  it('workstreams without focus behaves like detailed', () => {
    const r = resolveView(s, pres(), { view: 'workstreams' })
    expect(taskIds(r.rows)).toEqual(['a1', 'a2', 'm1', 'b1', 'b2'])
  })

  it('critical: flat critical tasks only', () => {
    const r = resolveView(s, pres(), { view: 'critical' })
    expect(r.rows).toEqual([{ kind: 'task', taskId: 'a1', groupKey: 'a', depth: 0 }])
    expect(r.reductions.byView).toBe(4)
  })

  it('critical: no criticality data', () => {
    const r = resolveView(schedule([task('x'), task('y')]), pres(), { view: 'critical' })
    expect(r.rows).toEqual([])
    expect(r.emptyReason).toBe('no-criticality-data')
  })

  it('critical: data present but none critical', () => {
    const r = resolveView(schedule([task('x', { critical: 'not-critical' })]), pres(), { view: 'critical' })
    expect(r.rows).toEqual([])
    expect(r.emptyReason).toBe('no-critical-tasks')
  })

  it('critical: filters removing every critical task reports no-matching-tasks', () => {
    const r = resolveView(s, pres(), { view: 'critical', filters: { text: 'nothing' } })
    expect(r.emptyReason).toBe('no-matching-tasks')
    expect(r.reductions.byFilters).toBe(1)
  })

  it('empty schedule reports no-tasks', () => {
    const r = resolveView(schedule([]), pres(), { view: 'detailed' })
    expect(r.rows).toEqual([])
    expect(r.emptyReason).toBe('no-tasks')
    expect(r.pageCount).toBe(1)
  })
})

describe('filters', () => {
  const s = schedule([
    task('a1', { name: 'Design review', critical: 'critical' }),
    task('a2', { name: 'Build', critical: 'not-critical', start: '2026-03-01', finish: '2026-03-10' }),
    task('b1', { workstream: 'b', name: 'Test', start: null, finish: null }),
  ])

  it('text matches source name and display label, case-insensitively', () => {
    const p = pres({ tasks: { b1: { displayLabel: 'Quality Gate' } } })
    expect(taskIds(resolveView(s, p, { view: 'detailed', filters: { text: 'DESIGN' } }).rows)).toEqual(['a1'])
    const r = resolveView(s, p, { view: 'detailed', filters: { text: 'quality' } })
    expect(taskIds(r.rows)).toEqual(['b1'])
    expect(r.reductions.byFilters).toBe(2)
  })

  it('criticality filter includes unavailable', () => {
    const r = resolveView(s, pres(), { view: 'detailed', filters: { criticality: ['unavailable' as Criticality] } })
    expect(taskIds(r.rows)).toEqual(['b1'])
  })

  it('groups filter limits to display groups', () => {
    const r = resolveView(s, pres(), { view: 'detailed', filters: { groups: ['b'] } })
    expect(taskIds(r.rows)).toEqual(['b1'])
  })

  it('overlap requires both dates and an intersecting window', () => {
    const r = resolveView(s, pres(), { view: 'detailed', filters: { overlap: { start: '2026-03-10', finish: '2026-03-20' } } })
    expect(taskIds(r.rows)).toEqual(['a2'])
    const edge = resolveView(s, pres(), { view: 'detailed', filters: { overlap: { start: '2026-01-09', finish: '2026-01-09' } } })
    expect(taskIds(edge.rows)).toEqual(['a1'])
  })

  it('all tasks filtered out reports no-matching-tasks', () => {
    const r = resolveView(s, pres(), { view: 'detailed', filters: { text: 'zzz' } })
    expect(r.rows).toEqual([])
    expect(r.emptyReason).toBe('no-matching-tasks')
    expect(r.groups).toHaveLength(2)
  })
})

describe('collapse and pagination', () => {
  it('collapsed groups drop task rows and count them', () => {
    const s = schedule([task('a1'), task('a2'), task('b1', { workstream: 'b' })])
    const r = resolveView(s, pres(), { view: 'detailed', collapsed: ['a'] })
    expect(r.rows).toEqual([
      { kind: 'group', key: 'a', label: 'Alpha', count: 2, collapsed: true, start: '2026-01-05', finish: '2026-01-09' },
      { kind: 'group', key: 'b', label: 'Beta', count: 1, collapsed: false, start: '2026-01-05', finish: '2026-01-09' },
      { kind: 'task', taskId: 'b1', groupKey: 'b', depth: 0 },
    ])
    expect(r.reductions.inCollapsedGroups).toBe(2)
    expect(total(r) + 1).toBe(3)
  })

  const many = schedule([
    ...Array.from({ length: 25 }, (_, i) => task(`a${String(i).padStart(2, '0')}`)),
    ...Array.from({ length: 20 }, (_, i) => task(`b${String(i).padStart(2, '0')}`, { workstream: 'b' })),
  ])

  it('paginates 45 tasks over 3 pages, repeating group headers', () => {
    const state: ViewState = { view: 'detailed' }
    const p1 = resolveView(many, pres(), { ...state, page: 1 })
    const p2 = resolveView(many, pres(), { ...state, page: 2 })
    const p3 = resolveView(many, pres(), { ...state, page: 3 })
    expect([p1.pageCount, p1.pageSize]).toEqual([3, 20])
    expect(taskIds(p1.rows)).toHaveLength(20)
    expect(taskIds(p2.rows)).toHaveLength(20)
    expect(taskIds(p3.rows)).toHaveLength(5)
    expect(p1.rows[0]).toMatchObject({ kind: 'group', key: 'a', count: 25 })
    expect(p2.rows[0]).toMatchObject({ kind: 'group', key: 'a', count: 25 })
    expect(p2.rows.filter((r) => r.kind === 'group').map((r) => (r.kind === 'group' ? r.key : ''))).toEqual(['a', 'b'])
    expect(p3.rows.filter((r) => r.kind === 'group')).toHaveLength(1)
    expect(p3.rows[0]).toMatchObject({ key: 'b' })
    expect(p1.reductions.onOtherPages).toBe(25)
    expect(p2.reductions.onOtherPages).toBe(25)
    expect(p3.reductions.onOtherPages).toBe(40)
    const all = [...taskIds(p1.rows), ...taskIds(p2.rows), ...taskIds(p3.rows)]
    expect(new Set(all).size).toBe(45)
  })

  it('clamps the page and honours pageSize', () => {
    expect(resolveView(many, pres(), { view: 'detailed', page: 99 }).page).toBe(3)
    expect(resolveView(many, pres(), { view: 'detailed', page: -4 }).page).toBe(1)
    const r = resolveView(many, pres(), { view: 'detailed', pageSize: 10, page: 5 })
    expect([r.pageCount, r.page]).toEqual([5, 5])
  })

  it('buckets sum to the total with view, filters, collapse and pages combined', () => {
    const state: ViewState = { view: 'workstreams', focusGroup: 'a', filters: { text: 'task a1' }, collapsed: [], page: 1, pageSize: 5 }
    const r = resolveView(many, pres(), state)
    expect(r.reductions.byView).toBe(20)
    expect(r.reductions.byFilters).toBe(15)
    expect(taskIds(r.rows)).toHaveLength(5)
    expect(r.reductions.onOtherPages).toBe(5)
    expect(total(r) + taskIds(r.rows).length).toBe(45)

    const c = resolveView(many, pres(), { view: 'detailed', collapsed: ['b'], page: 2 })
    expect(c.reductions.inCollapsedGroups).toBe(20)
    expect(total(c) + taskIds(c.rows).length).toBe(45)
    expect(c.rows.map((x) => x.kind)).toEqual(['group', 'task', 'task', 'task', 'task', 'task', 'group'])
  })
})

describe('hidden dependencies', () => {
  const s = schedule([
    task('a1'),
    task('a2', { deps: [dep('a1'), dep('b1'), dep('nope')] }),
    task('a3', { deps: [dep('a2')] }),
    task('b1', { workstream: 'b', critical: 'critical' }),
    task('b2', { workstream: 'b', deps: [dep('a2')] }),
  ])

  it('reports by direction and reason', () => {
    const r = resolveView(s, pres(), {
      view: 'detailed',
      filters: { text: 'task a' },
      collapsed: [],
      pageSize: 2,
    })
    expect(taskIds(r.rows)).toEqual(['a1', 'a2'])
    expect(r.hiddenDeps).toEqual(
      expect.arrayContaining([
        { taskId: 'a2', direction: 'predecessor', count: 1, reason: 'filtered' },
        { taskId: 'a2', direction: 'successor', count: 1, reason: 'other-page' },
        { taskId: 'a2', direction: 'successor', count: 1, reason: 'filtered' },
      ]),
    )
    expect(r.hiddenDeps.find((h) => h.taskId === 'a1' && h.direction === 'successor')).toBeUndefined()
  })

  it('reports collapsed and view reasons', () => {
    const c = resolveView(s, pres(), { view: 'detailed', collapsed: ['b'] })
    expect(c.hiddenDeps).toContainEqual({ taskId: 'a2', direction: 'predecessor', count: 1, reason: 'collapsed' })
    expect(c.hiddenDeps).toContainEqual({ taskId: 'a2', direction: 'successor', count: 1, reason: 'collapsed' })
    const v = resolveView(s, pres(), { view: 'critical' })
    expect(v.hiddenDeps).toEqual([{ taskId: 'b1', direction: 'successor', count: 1, reason: 'view' }])
  })

  it('emits nothing when every related task is visible', () => {
    expect(resolveView(s, pres(), { view: 'detailed' }).hiddenDeps).toEqual([])
  })
})

describe('reachability', () => {
  const s = schedule([
    task('a'),
    task('b', { deps: [dep('a')] }),
    task('c', { deps: [dep('b')] }),
    task('d', { deps: [dep('c'), dep('a')] }),
    task('x', { deps: [dep('d')] }),
  ])

  it('walks successors transitively', () => {
    expect([...downstreamOf(s, 'a')].sort()).toEqual(['b', 'c', 'd', 'x'])
    expect([...downstreamOf(s, 'c')].sort()).toEqual(['d', 'x'])
    expect(downstreamOf(s, 'x').size).toBe(0)
    expect(downstreamOf(s, 'missing').size).toBe(0)
  })

  it('is cycle-safe and excludes the start task', () => {
    const cyc = schedule([task('a', { deps: [dep('c')] }), task('b', { deps: [dep('a')] }), task('c', { deps: [dep('b')] })])
    expect([...downstreamOf(cyc, 'a')].sort()).toEqual(['b', 'c'])
    const self = schedule([task('a', { deps: [dep('a')] })])
    expect(downstreamOf(self, 'a').size).toBe(0)
  })

  it('exposes direct links', () => {
    expect(successorsOf(s, 'a').sort()).toEqual(['b', 'd'])
    expect(predecessorsOf(s, 'd').sort()).toEqual(['a', 'c'])
  })
})
