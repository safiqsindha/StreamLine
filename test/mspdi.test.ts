import { describe, expect, it } from 'vitest'
import { parseMspdi } from '../src/core/import/mspdi'
import { ImportError } from '../src/core/import/types'
import { buildMspdi, sampleMspdi, type TaskSpec } from './fixtures/mspdi'

const NOT_MSPDI = 'This .xml is not an MSPDI export. In Project, use File → Save As → XML.'

function parse(tasks: TaskSpec[], extra: Partial<Parameters<typeof buildMspdi>[0]> = {}) {
  return parseMspdi(buildMspdi({ name: 'Test', tasks, ...extra }), 'test.xml')
}

const task = (p: ReturnType<typeof parse>, id: string) => p.tasks.find((t) => t.id === id)!

describe('validation', () => {
  it('rejects non-MSPDI xml with the exact message', () => {
    expect(() => parseMspdi('<note><to>x</to></note>', 'a.xml')).toThrow(NOT_MSPDI)
    expect(() => parseMspdi(buildMspdi({ tasks: [], xmlns: null }), 'a.xml')).toThrow(NOT_MSPDI)
    expect(() => parseMspdi(buildMspdi({ tasks: [], xmlns: 'urn:other' }), 'a.xml')).toThrow(NOT_MSPDI)
    expect(() => parseMspdi(buildMspdi({ tasks: [], root: 'Schedule' }), 'a.xml')).toThrow(NOT_MSPDI)
  })

  it('accepts a versioned namespace', () => {
    const p = parseMspdi(buildMspdi({ tasks: [{ uid: 1, name: 'A' }], xmlns: 'http://schemas.microsoft.com/project/2007' }), 'a.xml')
    expect(p.tasks).toHaveLength(1)
  })

  it('throws ImportError for malformed xml', () => {
    const bad = '<Project xmlns="http://schemas.microsoft.com/project"><Tasks><Task></Tasks>'
    expect(() => parseMspdi(bad, 'a.xml')).toThrow(ImportError)
    expect(() => parseMspdi('not xml at all', 'a.xml')).toThrow(ImportError)
  })
})

describe('tasks', () => {
  it('keeps UIDs as strings and skips the project summary task', () => {
    const p = parse([
      { uid: '007', name: 'Leading zeros' },
      { uid: 12, name: 'Twelve' },
    ])
    expect(p.tasks.map((t) => t.id)).toEqual(['007', '12'])
    expect(p.issues).toEqual([])
    expect(p.kind).toBe('mspdi')
  })

  it('takes the date part with no time zone shift', () => {
    const p = parse([
      { uid: 1, name: 'Late', start: '2026-10-08T23:59:59', finish: '2026-10-09T00:00:00' },
      { uid: 2, name: 'Offset', start: '2026-10-08T08:00:00-08:00', finish: '2026-10-08T17:00:00Z' },
      { uid: 3, name: 'No dates' },
    ])
    expect(task(p, '1').start).toBe('2026-10-08')
    expect(task(p, '1').finish).toBe('2026-10-09')
    expect(task(p, '2').start).toBe('2026-10-08')
    expect(task(p, '2').finish).toBe('2026-10-08')
    expect(task(p, '3').start).toBeNull()
    expect(task(p, '3').finish).toBeNull()
  })

  it('maps milestone and summary flags', () => {
    const p = parse([
      { uid: 1, name: 'S', summary: true },
      { uid: 2, name: 'M', level: 2, milestone: true },
    ])
    expect(task(p, '1').isSummary).toBe(true)
    expect(task(p, '1').milestone).toBe(false)
    expect(task(p, '2').milestone).toBe(true)
  })

  it('reads criticality as a tri-state', () => {
    const p = parse([
      { uid: 1, name: 'Yes', critical: true },
      { uid: 2, name: 'No', critical: false },
      { uid: 3, name: 'Absent' },
    ])
    expect(task(p, '1').critical).toBe('critical')
    expect(task(p, '2').critical).toBe('not-critical')
    expect(task(p, '3').critical).toBe('unavailable')
  })

  it('converts total slack from tenths of minutes using MinutesPerDay', () => {
    const p = parse(
      [
        { uid: 1, name: 'Two days', totalSlack: 9600 },
        { uid: 2, name: 'Fractional', totalSlack: 7200 },
        { uid: 3, name: 'Negative', totalSlack: -4800 },
        { uid: 4, name: 'Absent' },
      ],
      { minutesPerDay: 480 },
    )
    expect(task(p, '1').floatDays).toBe(2)
    expect(task(p, '2').floatDays).toBe(1.5)
    expect(task(p, '3').floatDays).toBe(-1)
    expect(task(p, '4').floatDays).toBeNull()

    const q = parse([{ uid: 1, name: 'A', totalSlack: 6000 }], { minutesPerDay: 600 })
    expect(task(q, '1').floatDays).toBe(1)
    const r = parse([{ uid: 1, name: 'A', totalSlack: 4800 }])
    expect(task(r, '1').floatDays).toBe(1)
  })

  it('skips tasks that finish before they start', () => {
    const p = parse([
      { uid: 5, name: 'Backwards', start: '2026-10-10T08:00:00', finish: '2026-10-09T17:00:00' },
      { uid: 6, name: 'Fine', start: '2026-10-10T08:00:00', finish: '2026-10-10T17:00:00' },
    ])
    expect(p.tasks.map((t) => t.id)).toEqual(['6'])
    expect(p.issues).toContainEqual(
      expect.objectContaining({
        severity: 'skipped',
        message: 'Task 5 ("Backwards") has a finish date before its start date. That task was skipped.',
      }),
    )
  })

  it('skips tasks without a name', () => {
    const p = parse([{ uid: 1 }, { uid: 2, name: 'Named' }])
    expect(p.tasks.map((t) => t.id)).toEqual(['2'])
    expect(p.issues).toHaveLength(1)
    expect(p.issues[0].severity).toBe('skipped')
  })

  it('handles a single task and decodes entities', () => {
    const p = parse([{ uid: 1, name: 'R&D <phase> "one"' }])
    expect(p.tasks).toHaveLength(1)
    expect(p.tasks[0].name).toBe('R&D <phase> "one"')
  })

  it('names the schedule from the project, then the filename', () => {
    expect(parse([{ uid: 1, name: 'A' }]).name).toBe('Test')
    const t = parseMspdi(buildMspdi({ title: 'Titled', tasks: [] }), 'f.xml')
    expect(t.name).toBe('Titled')
    const n = parseMspdi(buildMspdi({ tasks: [] }), 'my-plan.v2.xml')
    expect(n.name).toBe('my-plan.v2')
  })
})

describe('hierarchy and workstreams', () => {
  const tasks: TaskSpec[] = [
    { uid: 1, name: 'Design', level: 1, summary: true },
    { uid: 2, name: 'Sub', level: 2, summary: true },
    { uid: 3, name: 'Leaf', level: 3 },
    { uid: 4, name: 'Sibling', level: 2 },
    { uid: 5, name: 'Standalone', level: 1 },
    { uid: 6, name: 'Build', level: 1, summary: true },
    { uid: 7, name: 'Part', level: 2 },
  ]

  it('derives parentId from outline level', () => {
    const p = parse(tasks)
    expect(p.tasks.map((t) => [t.id, t.parentId])).toEqual([
      ['1', null],
      ['2', '1'],
      ['3', '2'],
      ['4', '1'],
      ['5', null],
      ['6', null],
      ['7', '6'],
    ])
  })

  it('falls back to OutlineNumber when OutlineLevel is absent', () => {
    const xml = buildMspdi({ tasks: [], projectSummary: false }).replace(
      '<Tasks></Tasks>',
      '<Tasks><Task><UID>1</UID><Name>A</Name><OutlineNumber>1</OutlineNumber></Task><Task><UID>2</UID><Name>B</Name><OutlineNumber>1.1</OutlineNumber></Task></Tasks>',
    )
    const p = parseMspdi(xml, 'a.xml')
    expect(task(p, '2').parentId).toBe('1')
  })

  it('uses the top-level summary as the workstream', () => {
    const p = parse(tasks)
    expect(p.tasks.map((t) => [t.id, t.workstream])).toEqual([
      ['1', 'Design'],
      ['2', 'Design'],
      ['3', 'Design'],
      ['4', 'Design'],
      ['5', null],
      ['6', 'Build'],
      ['7', 'Build'],
    ])
    expect(p.workstreams).toEqual([
      { id: 'Design', name: 'Design' },
      { id: 'Build', name: 'Build' },
    ])
  })

  it('reattaches children of a skipped summary to the nearest kept ancestor', () => {
    const p = parse([
      { uid: 1, name: 'Top', level: 1, summary: true },
      { uid: 2, level: 2, summary: true },
      { uid: 3, name: 'Child', level: 3 },
    ])
    expect(task(p, '3').parentId).toBe('1')
  })
})

describe('owners', () => {
  it('joins assigned resource names in resource UID order', () => {
    const p = parse(
      [
        { uid: 1, name: 'A' },
        { uid: 2, name: 'B' },
        { uid: 3, name: 'C' },
      ],
      {
        resources: [
          { uid: 1, name: 'Avery' },
          { uid: 2, name: 'Jordan' },
          { uid: 3, name: 'Sam' },
        ],
        assignments: [
          { taskUid: 1, resourceUid: 3 },
          { taskUid: 1, resourceUid: 1 },
          { taskUid: 2, resourceUid: 2 },
        ],
      },
    )
    expect(task(p, '1').owner).toBe('Avery, Sam')
    expect(task(p, '2').owner).toBe('Jordan')
    expect(task(p, '3').owner).toBeNull()
  })
})

describe('dependencies', () => {
  const base: TaskSpec[] = [
    { uid: 1, name: 'P1' },
    { uid: 2, name: 'P2' },
    { uid: 3, name: 'P3' },
    { uid: 4, name: 'P4' },
    { uid: 5, name: 'P5' },
  ]

  it('maps link types and converts lag to days', () => {
    const p = parse([
      ...base,
      {
        uid: 10,
        name: 'Succ',
        preds: [
          { uid: 1, type: 0, lag: 4800, lagFormat: 7 },
          { uid: 2, type: 1, lag: 0, lagFormat: 7 },
          { uid: 3, type: 2, lag: -2400, lagFormat: 7 },
          { uid: 4, type: 3, lag: 9600, lagFormat: 7 },
          { uid: 5 },
        ],
      },
    ])
    expect(task(p, '10').deps).toEqual([
      { predecessorId: '1', type: 'FF', lagDays: 1 },
      { predecessorId: '2', type: 'FS', lagDays: 0 },
      { predecessorId: '3', type: 'SF', lagDays: -0.5 },
      { predecessorId: '4', type: 'SS', lagDays: 2 },
      { predecessorId: '5', type: 'unknown', lagDays: null },
    ])
  })

  it('treats a missing or unrecognised type as unknown', () => {
    const p = parse([...base, { uid: 10, name: 'S', preds: [{ uid: 1 }, { uid: 2, type: 9, lag: 0 }] }])
    expect(task(p, '10').deps.map((d) => d.type)).toEqual(['unknown', 'unknown'])
  })

  it('converts elapsed and non-day formats', () => {
    const p = parse([
      ...base,
      {
        uid: 10,
        name: 'S',
        preds: [
          { uid: 1, type: 1, lag: 14400, lagFormat: 8 },
          { uid: 2, type: 1, lag: 4800, lagFormat: 5 },
          { uid: 3, type: 1, lag: 48000, lagFormat: 9 },
        ],
      },
    ])
    expect(task(p, '10').deps.map((d) => d.lagDays)).toEqual([1, 1, 10])
  })

  it('leaves percentage lags unavailable with a warning', () => {
    const p = parse([...base, { uid: 10, name: 'S', preds: [{ uid: 1, type: 1, lag: 2500, lagFormat: 19 }] }])
    expect(task(p, '10').deps[0].lagDays).toBeNull()
    expect(p.issues.some((i) => i.severity === 'warning' && /percentage/.test(i.message))).toBe(true)
  })

  it('drops links to tasks that are not in the file, with a warning', () => {
    const p = parse([{ uid: 1, name: 'A' }, { uid: 2, name: 'B', preds: [{ uid: 1, type: 1 }, { uid: 99, type: 1 }] }])
    expect(task(p, '2').deps.map((d) => d.predecessorId)).toEqual(['1'])
    expect(p.issues).toContainEqual({
      message: 'Task 2 ("B") depends on task 99, which is not in the file. That link was dropped.',
      severity: 'warning',
    })
  })
})

describe('reporting', () => {
  it('reports detected columns, mapping and unmapped fields', () => {
    const p = parse([{ uid: 1, name: 'A', extra: '<Duration>PT8H0M0S</Duration><Cost>0</Cost>' }])
    expect(p.detectedColumns).toEqual(expect.arrayContaining(['UID', 'Name', 'Duration', 'Cost']))
    expect(p.unmapped).toEqual(expect.arrayContaining(['Duration', 'Cost', 'ID']))
    expect(p.unmapped).not.toContain('UID')
    expect(p.mapping.id).toBe('UID')
    expect(new Set(p.detectedColumns).size).toBe(p.detectedColumns.length)
  })
})

describe('sample', () => {
  it('imports the synthetic sample cleanly', () => {
    const p = parseMspdi(sampleMspdi(), 'evt-program.xml')
    expect(p.name).toBe('EVT Program')
    expect(p.issues).toEqual([])
    expect(p.tasks).toHaveLength(21)
    expect(p.workstreams.map((w) => w.name)).toEqual(['Design', 'Build', 'Validation'])
    expect(p.tasks.filter((t) => t.milestone)).toHaveLength(2)
    const types = new Set(p.tasks.flatMap((t) => t.deps.map((d) => d.type)))
    expect(types).toEqual(new Set(['FS', 'SS', 'FF', 'SF']))
    expect(task(p, '3').owner).toBe('Avery Chen, Jordan Patel')
  })
})
