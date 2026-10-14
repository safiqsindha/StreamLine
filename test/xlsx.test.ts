import { describe, expect, it } from 'vitest'
import { fallbackTaskId } from '../src/core/ids'
import { ImportError } from '../src/core/import/types'
import { parseXlsx } from '../src/core/import/xlsx'
import { d, makeXlsx, writeSampleFixtures, type Cell } from './fixtures/makeXlsx'
import { mkdtemp, readFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'

const H = ['Task ID', 'Task Name', 'Workstream', 'Owner', 'Start', 'Finish']

async function parse(header: string[], rows: Cell[][], mapping?: Parameters<typeof parseXlsx>[2]) {
  return parseXlsx(await makeXlsx(header, rows), 'Program Alpha.xlsx', mapping)
}

async function failure(p: Promise<unknown>): Promise<string> {
  try {
    await p
  } catch (e) {
    expect(e).toBeInstanceOf(ImportError)
    return (e as Error).message
  }
  throw new Error('expected ImportError')
}

describe('parseXlsx', () => {
  it('auto-detects columns by synonym, ignoring case and spacing', async () => {
    const p = await parse(
      ['wbs id', 'ACTIVITY', 'Swim Lane', 'Assigned To', 'Begin', 'End Date', 'Notes'],
      [['1', 'Kickoff', 'Design', 'Avery', d('2026-01-05'), d('2026-01-09'), 'x']],
    )
    expect(p.kind).toBe('xlsx')
    expect(p.name).toBe('Program Alpha')
    expect(p.detectedColumns).toEqual(['wbs id', 'ACTIVITY', 'Swim Lane', 'Assigned To', 'Begin', 'End Date', 'Notes'])
    expect(p.mapping).toMatchObject({
      id: 'wbs id', name: 'ACTIVITY', workstream: 'Swim Lane', owner: 'Assigned To', start: 'Begin', finish: 'End Date',
    })
    expect(p.unmapped).toEqual(['Notes'])
    expect(p.tasks[0]).toMatchObject({
      id: '1', name: 'Kickoff', workstream: 'Design', owner: 'Avery', start: '2026-01-05', finish: '2026-01-09',
    })
    expect(p.workstreams).toEqual([{ id: 'Design', name: 'Design' }])
  })

  it('finds the header row below a title block', async () => {
    const bytes = await makeXlsx(H, [['1', 'A', 'L', 'O', d('2026-01-05'), d('2026-01-06')]], {
      preamble: [['Program schedule'], [], ['Exported 2026-10-01']],
    })
    const p = await parseXlsx(bytes, 'x.xlsx')
    expect(p.tasks).toHaveLength(1)
    expect(p.issues).toEqual([])
  })

  it('uses an explicit mapping with exact header text', async () => {
    const p = await parse(
      ['Ref', 'Deliverable', 'Kick', 'Done', 'Task Name'],
      [['A1', 'Widget', d('2026-02-02'), d('2026-02-06'), 'ignored']],
      { id: 'Ref', name: 'Deliverable', start: 'Kick', finish: 'Done' },
    )
    expect(p.mapping).toEqual({ id: 'Ref', name: 'Deliverable', start: 'Kick', finish: 'Done' })
    expect(p.unmapped).toEqual(['Task Name'])
    expect(p.tasks[0]).toMatchObject({ id: 'A1', name: 'Widget', start: '2026-02-02', finish: '2026-02-06' })
  })

  it('reads Excel date cells by UTC parts, including near midnight', async () => {
    const p = await parse(H, [
      ['1', 'A', null, null, new Date('2026-03-01T00:00:00Z'), new Date('2026-12-31T00:00:00Z')],
      ['2', 'B', null, null, d('2024-02-29'), d('2024-03-01')],
    ])
    expect(p.tasks.map((t) => [t.start, t.finish])).toEqual([
      ['2026-03-01', '2026-12-31'],
      ['2024-02-29', '2024-03-01'],
    ])
  })

  it('parses common text date forms', async () => {
    const p = await parse(H, [
      ['1', 'A', null, null, '2026-10-14', '10/20/2026'],
      ['2', 'B', null, null, '14-Oct-2026', '14 Oct 2026'],
      ['3', 'C', null, null, 'Oct 14, 2026', '2026-10-14T08:00:00'],
    ])
    expect(p.tasks.map((t) => [t.start, t.finish])).toEqual([
      ['2026-10-14', '2026-10-20'],
      ['2026-10-14', '2026-10-14'],
      ['2026-10-14', '2026-10-14'],
    ])
    expect(p.issues).toEqual([])
  })

  it('warns on an unparseable date and leaves it null', async () => {
    const p = await parse(H, [['1', 'A', null, null, 'soon', d('2026-01-09')]])
    expect(p.tasks[0].start).toBeNull()
    expect(p.tasks[0].finish).toBe('2026-01-09')
    expect(p.issues).toHaveLength(1)
    expect(p.issues[0]).toMatchObject({ row: 2, severity: 'warning' })
    expect(p.issues[0].message).toContain('Row 2')
    expect(p.issues[0].message).toContain('"Start"')
  })

  it('rejects a file with no date columns, with the exact message', async () => {
    const msg = await failure(parse(['Task', 'Owner'], [['A', 'x']]))
    expect(msg).toBe('This file has no recognizable date columns. Streamline needs a start and finish column for each task.')
  })

  it('rejects a file with no name column', async () => {
    const msg = await failure(parse(['Start', 'Finish'], [[d('2026-01-01'), d('2026-01-02')]]))
    expect(msg).toMatch(/task name/i)
  })

  it('rejects bytes that are not an xlsx', async () => {
    const msg = await failure(parseXlsx(new TextEncoder().encode('not a workbook'), 'x.xlsx'))
    expect(msg).toMatch(/\.xlsx/)
  })

  it('rejects a file with no usable rows', async () => {
    const msg = await failure(parse(H, []))
    expect(msg).toMatch(/no usable tasks/i)
  })

  it('skips finish-before-start rows with the real sheet row number', async () => {
    const rows: Cell[][] = []
    for (let i = 0; i < 46; i++) rows.push([`T${i}`, `Task ${i}`, null, null, d('2026-01-05'), d('2026-01-06')])
    rows.push(['BAD', 'Backwards', null, null, d('2026-02-10'), d('2026-02-01')]) // sheet row 48
    const p = await parse(H, rows)
    expect(p.tasks.find((t) => t.id === 'BAD')).toBeUndefined()
    expect(p.issues).toEqual([
      { row: 48, severity: 'skipped', message: 'Row 48 has a finish date before its start date. That row was skipped.' },
    ])
  })

  it('skips blank rows silently and nameless rows with an issue', async () => {
    const p = await parse(H, [
      ['1', 'A', null, null, d('2026-01-05'), d('2026-01-06')],
      [null, null, null, null, null, null],
      ['3', null, 'Lane', null, d('2026-01-05'), d('2026-01-06')],
    ])
    expect(p.tasks).toHaveLength(1)
    expect(p.issues).toEqual([{ row: 4, severity: 'skipped', message: 'Row 4 has no task name. That row was skipped.' }])
  })

  it('leaves absent values null and never infers', async () => {
    const p = await parse(['Task Name', 'Start', 'Finish'], [
      ['Only name', null, null],
      ['Same day', d('2026-01-05'), d('2026-01-05')],
    ])
    expect(p.tasks[0]).toMatchObject({
      workstream: null, owner: null, start: null, finish: null, milestone: false, floatDays: null,
      parentId: null, isSummary: false, deps: [],
    })
    expect(p.tasks[1].milestone).toBe(false)
    expect(p.workstreams).toEqual([])
  })

  it('treats blank owner and workstream cells as null', async () => {
    const p = await parse(H, [['1', 'A', '  ', '', d('2026-01-05'), d('2026-01-06')]])
    expect(p.tasks[0].workstream).toBeNull()
    expect(p.tasks[0].owner).toBeNull()
  })

  it('reads criticality as tri-state', async () => {
    const p = await parse(
      ['ID', 'Task', 'Start', 'Finish', 'Critical'],
      ['yes', 'TRUE', 1, 'Y', 'Critical', 'no', 'false', 0, 'n', null, 'maybe'].map(
        (v, i) => [`${i}`, `T${i}`, d('2026-01-05'), d('2026-01-06'), v] as Cell[],
      ),
    )
    expect(p.tasks.map((t) => t.critical)).toEqual([
      'critical', 'critical', 'critical', 'critical', 'critical',
      'not-critical', 'not-critical', 'not-critical', 'not-critical',
      'unavailable', 'unavailable',
    ])
  })

  it('marks all tasks unavailable when there is no critical column', async () => {
    const p = await parse(H, [['1', 'A', null, null, d('2026-01-05'), d('2026-01-06')]])
    expect(p.tasks.every((t) => t.critical === 'unavailable')).toBe(true)
  })

  it('reads float, milestone, parent and summary columns', async () => {
    const p = await parse(
      ['ID', 'Task', 'Start', 'Finish', 'Float', 'Type', 'Parent', 'Summary'],
      [
        ['1', 'Phase', d('2026-01-05'), d('2026-01-30'), 4, 'Task', null, 'Yes'],
        ['2', 'Child', d('2026-01-05'), d('2026-01-09'), '3d', 'Milestone', '1', 'No'],
        ['3', 'Other', d('2026-01-05'), d('2026-01-09'), '2 days', null, '1', null],
        ['4', 'Bad float', d('2026-01-05'), d('2026-01-09'), 'lots', null, null, null],
      ],
    )
    expect(p.tasks.map((t) => t.floatDays)).toEqual([4, 3, 2, null])
    expect(p.tasks.map((t) => t.milestone)).toEqual([false, true, false, false])
    expect(p.tasks.map((t) => t.parentId)).toEqual([null, '1', '1', null])
    expect(p.tasks.map((t) => t.isSummary)).toEqual([true, false, false, false])
    expect(p.issues.filter((i) => i.severity === 'warning')).toHaveLength(1)
  })

  it('parses dependency types and lags, defaulting type to unknown', async () => {
    const dates = [d('2026-01-05'), d('2026-01-06')]
    const p = await parse(
      ['ID', 'Task', 'Start', 'Finish', 'Predecessors'],
      [
        ['1', 'A', ...dates, null],
        ['2', 'B', ...dates, '1'],
        ['3', 'C', ...dates, '1FS'],
        ['4', 'D', ...dates, '1SS+2d'],
        ['5', 'E', ...dates, '1FF-1; 2'],
        ['6', 'F', ...dates, '1 SF +3 days, 3'],
      ],
    )
    const deps = (id: string) => p.tasks.find((t) => t.id === id)!.deps
    expect(deps('1')).toEqual([])
    expect(deps('2')).toEqual([{ predecessorId: '1', type: 'unknown', lagDays: null }])
    expect(deps('3')).toEqual([{ predecessorId: '1', type: 'FS', lagDays: null }])
    expect(deps('4')).toEqual([{ predecessorId: '1', type: 'SS', lagDays: 2 }])
    expect(deps('5')).toEqual([
      { predecessorId: '1', type: 'FF', lagDays: -1 },
      { predecessorId: '2', type: 'unknown', lagDays: null },
    ])
    expect(deps('6')).toEqual([
      { predecessorId: '1', type: 'SF', lagDays: 3 },
      { predecessorId: '3', type: 'unknown', lagDays: null },
    ])
    expect(p.issues).toEqual([])
  })

  it('accepts numeric predecessor cells and hyphenated IDs', async () => {
    const dates = [d('2026-01-05'), d('2026-01-06')]
    const p = await parse(
      ['ID', 'Task', 'Start', 'Finish', 'Predecessors'],
      [['10', 'A', ...dates, null], ['B-1', 'B', ...dates, 10], ['C', 'C', ...dates, 'B-1FS+1d']],
    )
    expect(p.tasks[1].deps).toEqual([{ predecessorId: '10', type: 'unknown', lagDays: null }])
    expect(p.tasks[2].deps).toEqual([{ predecessorId: 'B-1', type: 'FS', lagDays: 1 }])
  })

  it('drops predecessors that are not in the file, with a warning', async () => {
    const dates = [d('2026-01-05'), d('2026-01-06')]
    const p = await parse(
      ['ID', 'Task', 'Start', 'Finish', 'Predecessors'],
      [['1', 'A', ...dates, '99FS, 1'], ['2', 'B', ...dates, '1']],
    )
    expect(p.tasks[0].deps).toEqual([])
    expect(p.tasks[1].deps).toHaveLength(1)
    expect(p.issues).toHaveLength(2)
    expect(p.issues[0]).toMatchObject({ row: 2, severity: 'warning' })
    expect(p.issues[0].message).toContain('"99"')
    expect(p.issues[0].message).toContain('dropped')
  })

  it('drops all dependencies with a warning when there is no ID column', async () => {
    const p = await parse(
      ['Task', 'Start', 'Finish', 'Predecessors'],
      [['A', d('2026-01-05'), d('2026-01-06'), null], ['B', d('2026-01-05'), d('2026-01-06'), 'A']],
    )
    expect(p.tasks[1].deps).toEqual([])
    expect(p.issues).toHaveLength(1)
    expect(p.issues[0]).toMatchObject({ row: 3, severity: 'warning' })
    expect(p.issues[0].message).toMatch(/no ID column/)
    expect(p.issues[0].message).toMatch(/dropped/)
  })

  it('skips repeated source IDs after the first', async () => {
    const p = await parse(H, [
      ['1', 'First', null, null, d('2026-01-05'), d('2026-01-06')],
      ['1', 'Second', null, null, d('2026-01-05'), d('2026-01-06')],
    ])
    expect(p.tasks.map((t) => t.name)).toEqual(['First'])
    expect(p.issues).toHaveLength(1)
    expect(p.issues[0]).toMatchObject({ row: 3, severity: 'skipped' })
    expect(p.issues[0].message).toContain('"1"')
  })

  describe('fallback task IDs', () => {
    const nh = ['Task Name', 'Workstream', 'Start', 'Finish']
    const row = (name: string, ws: string | null): Cell[] => [name, ws, d('2026-01-05'), d('2026-01-06')]

    it('are deterministic across parses and use fallbackTaskId', async () => {
      const rows = [row('A', 'Design'), row('B', 'Build')]
      const a = await parse(nh, rows)
      const b = await parse(nh, rows)
      expect(a.tasks.map((t) => t.id)).toEqual(b.tasks.map((t) => t.id))
      expect(a.tasks[0].id).toBe(fallbackTaskId(['Design', 'A'], 0))
    })

    it('survive row reordering', async () => {
      const a = await parse(nh, [row('A', 'Design'), row('B', 'Build'), row('C', null)])
      const b = await parse(nh, [row('C', null), row('A', 'Design'), row('B', 'Build')])
      const idOf = (p: typeof a, name: string) => p.tasks.find((t) => t.name === name)!.id
      for (const n of ['A', 'B', 'C']) expect(idOf(a, n)).toBe(idOf(b, n))
    })

    it('give duplicate names distinct, stable IDs', async () => {
      const rows = [row('Review', 'Design'), row('Review', 'Design'), row('Review', 'Build')]
      const a = await parse(nh, rows)
      const b = await parse(nh, rows)
      const ids = a.tasks.map((t) => t.id)
      expect(new Set(ids).size).toBe(3)
      expect(ids).toEqual(b.tasks.map((t) => t.id))
      expect(ids[1]).toBe(fallbackTaskId(['Design', 'Review'], 1))
    })
  })

  it('lists workstreams in first-appearance order', async () => {
    const p = await parse(H, [
      ['1', 'A', 'Build', null, d('2026-01-05'), d('2026-01-06')],
      ['2', 'B', 'Design', null, d('2026-01-05'), d('2026-01-06')],
      ['3', 'C', 'Build', null, d('2026-01-05'), d('2026-01-06')],
      ['4', 'D', null, null, d('2026-01-05'), d('2026-01-06')],
    ])
    expect(p.workstreams).toEqual([
      { id: 'Build', name: 'Build' },
      { id: 'Design', name: 'Design' },
    ])
  })

  it('imports the generated sample program', async () => {
    const dir = await mkdtemp(join(tmpdir(), 'streamline-fixture-'))
    const path = await writeSampleFixtures(dir)
    const p = await parseXlsx(new Uint8Array(await readFile(path)), 'evt-program.xlsx')
    expect(p.tasks).toHaveLength(25)
    expect(p.workstreams.map((w) => w.name)).toEqual(['Design', 'Build', 'Validation'])
    expect(p.tasks.filter((t) => t.milestone)).toHaveLength(2)
    expect(p.tasks.some((t) => t.critical === 'unavailable')).toBe(true)
    expect(p.issues).toEqual([])
  })
})
