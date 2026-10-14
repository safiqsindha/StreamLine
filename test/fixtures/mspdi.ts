import { mkdirSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { fileURLToPath } from 'node:url'

export interface LinkSpec {
  uid: number | string
  /** MSPDI link type code: 0 FF, 1 FS, 2 SF, 3 SS. Omit to leave the element out. */
  type?: number
  /** Tenths of a minute. Omit to leave the element out. */
  lag?: number
  lagFormat?: number
}

export interface TaskSpec {
  uid: number | string
  name?: string
  level?: number
  outlineNumber?: string
  start?: string
  finish?: string
  summary?: boolean
  milestone?: boolean
  /** undefined leaves the Critical element out. */
  critical?: boolean
  /** Tenths of a minute. */
  totalSlack?: number
  preds?: LinkSpec[]
  /** Raw extra child elements, e.g. `<Duration>PT8H0M0S</Duration>`. */
  extra?: string
}

export interface ResourceSpec {
  uid: number
  name: string
}

export interface AssignmentSpec {
  taskUid: number | string
  resourceUid: number
}

export interface MspdiSpec {
  name?: string
  title?: string
  minutesPerDay?: number
  xmlns?: string | null
  root?: string
  /** Prepend the UID 0 project summary task. Default true. */
  projectSummary?: boolean
  tasks: TaskSpec[]
  resources?: ResourceSpec[]
  assignments?: AssignmentSpec[]
}

export const MSPDI_NS = 'http://schemas.microsoft.com/project'

export function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;')
}

function el(name: string, value: string | number | boolean | undefined): string {
  if (value === undefined) return ''
  const v = typeof value === 'boolean' ? (value ? '1' : '0') : String(value)
  return `<${name}>${v}</${name}>`
}

function taskXml(t: TaskSpec, id: number): string {
  const level = t.level ?? 1
  const links = (t.preds ?? [])
    .map(
      (p) =>
        `<PredecessorLink>${el('PredecessorUID', p.uid)}${el('Type', p.type)}${el('LinkLag', p.lag)}${el('LagFormat', p.lagFormat)}</PredecessorLink>`,
    )
    .join('')
  return (
    '<Task>' +
    el('UID', t.uid) +
    el('ID', id) +
    (t.name === undefined ? '' : el('Name', esc(t.name))) +
    el('OutlineNumber', t.outlineNumber) +
    el('OutlineLevel', level) +
    el('Summary', t.summary ?? false) +
    el('Milestone', t.milestone ?? false) +
    el('Start', t.start) +
    el('Finish', t.finish) +
    el('Critical', t.critical) +
    el('TotalSlack', t.totalSlack) +
    (t.extra ?? '') +
    links +
    '</Task>'
  )
}

export function buildMspdi(spec: MspdiSpec): string {
  const root = spec.root ?? 'Project'
  const ns = spec.xmlns === null ? '' : ` xmlns="${spec.xmlns ?? MSPDI_NS}"`
  const tasks: TaskSpec[] = [...spec.tasks]
  if (spec.projectSummary !== false) {
    tasks.unshift({
      uid: 0,
      name: spec.name ?? 'Project',
      level: 0,
      summary: true,
      start: '2026-01-01T08:00:00',
      finish: '2026-12-31T17:00:00',
      critical: false,
    })
  }
  const resources = (spec.resources ?? []).map((r) => `<Resource>${el('UID', r.uid)}${el('Name', esc(r.name))}</Resource>`).join('')
  const assignments = (spec.assignments ?? [])
    .map((a) => `<Assignment>${el('UID', a.resourceUid * 1000 + Number(a.taskUid))}${el('TaskUID', a.taskUid)}${el('ResourceUID', a.resourceUid)}</Assignment>`)
    .join('')
  return (
    '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n' +
    `<${root}${ns}>` +
    el('Name', spec.name === undefined ? undefined : esc(spec.name)) +
    el('Title', spec.title === undefined ? undefined : esc(spec.title)) +
    el('MinutesPerDay', spec.minutesPerDay) +
    `<Tasks>${tasks.map((t, i) => taskXml(t, i)).join('')}</Tasks>` +
    `<Resources>${resources}</Resources>` +
    `<Assignments>${assignments}</Assignments>` +
    `</${root}>`
  )
}

const DAY = 4800 // one 8-hour working day, in tenths of a minute

export function sampleSpec(): MspdiSpec {
  const d = (date: string, end = false) => `${date}T${end ? '17:00:00' : '08:00:00'}`
  const tasks: TaskSpec[] = [
    { uid: 1, name: 'Design', level: 1, outlineNumber: '1', summary: true, start: d('2026-10-05'), finish: d('2026-11-13', true), critical: true, totalSlack: 0 },
    { uid: 2, name: 'Requirements review', level: 2, outlineNumber: '1.1', start: d('2026-10-05'), finish: d('2026-10-09', true), critical: true, totalSlack: 0 },
    { uid: 3, name: 'Concept design', level: 2, outlineNumber: '1.2', start: d('2026-10-12'), finish: d('2026-10-23', true), critical: true, totalSlack: 0, preds: [{ uid: 2, type: 1, lag: 0, lagFormat: 7 }] },
    { uid: 4, name: 'Thermal analysis', level: 2, outlineNumber: '1.3', start: d('2026-10-19'), finish: d('2026-10-30', true), critical: false, totalSlack: 3 * DAY, preds: [{ uid: 3, type: 3, lag: 2 * DAY, lagFormat: 7 }] },
    { uid: 5, name: 'Detailed design', level: 2, outlineNumber: '1.4', start: d('2026-10-26'), finish: d('2026-11-12', true), critical: true, totalSlack: 0, preds: [{ uid: 3, type: 1, lag: 0, lagFormat: 7 }] },
    { uid: 6, name: 'Design freeze', level: 2, outlineNumber: '1.5', start: d('2026-11-13'), finish: d('2026-11-13'), critical: true, totalSlack: 0, preds: [{ uid: 5, type: 1, lag: 0, lagFormat: 7 }, { uid: 4, type: 0, lag: 0, lagFormat: 7 }] },
    { uid: 7, name: 'Build', level: 1, outlineNumber: '2', summary: true, start: d('2026-11-16'), finish: d('2027-01-29', true), critical: true, totalSlack: 0 },
    { uid: 8, name: 'Order long-lead parts', level: 2, outlineNumber: '2.1', start: d('2026-11-16'), finish: d('2026-11-20', true), critical: false, totalSlack: 5 * DAY, preds: [{ uid: 6, type: 1, lag: 0, lagFormat: 7 }] },
    { uid: 9, name: 'Fabricate boards', level: 2, outlineNumber: '2.2', start: d('2026-11-23'), finish: d('2026-12-18', true), critical: true, totalSlack: 0, preds: [{ uid: 8, type: 1, lag: 5 * DAY, lagFormat: 7 }] },
    { uid: 10, name: 'Machine enclosure', level: 2, outlineNumber: '2.3', start: d('2026-11-23'), finish: d('2026-12-11', true), critical: false, totalSlack: 10 * DAY, preds: [{ uid: 6, type: 1, lag: 5 * DAY, lagFormat: 7 }] },
    { uid: 11, name: 'Board assembly', level: 2, outlineNumber: '2.4', summary: true, start: d('2026-12-21'), finish: d('2027-01-15', true), critical: true, totalSlack: 0 },
    { uid: 12, name: 'Populate and solder', level: 3, outlineNumber: '2.4.1', start: d('2026-12-21'), finish: d('2027-01-06', true), critical: true, totalSlack: 0, preds: [{ uid: 9, type: 1, lag: 0, lagFormat: 7 }] },
    { uid: 13, name: 'Inspect assemblies', level: 3, outlineNumber: '2.4.2', start: d('2027-01-04'), finish: d('2027-01-15', true), critical: true, totalSlack: 0, preds: [{ uid: 12, type: 3, lag: 3600, lagFormat: 5 }] },
    { uid: 14, name: 'Final integration', level: 2, outlineNumber: '2.5', start: d('2027-01-18'), finish: d('2027-01-29', true), critical: true, totalSlack: 0, preds: [{ uid: 13, type: 1, lag: 0, lagFormat: 7 }, { uid: 10, type: 1, lag: 0, lagFormat: 7 }] },
    { uid: 15, name: 'Build complete', level: 2, outlineNumber: '2.6', start: d('2027-01-29'), finish: d('2027-01-29'), milestone: true, critical: true, totalSlack: 0, preds: [{ uid: 14, type: 2, lag: -1 * DAY, lagFormat: 7 }] },
    { uid: 16, name: 'Validation', level: 1, outlineNumber: '3', summary: true, start: d('2027-02-01'), finish: d('2027-03-12', true), critical: true, totalSlack: 0 },
    { uid: 17, name: 'Write test procedures', level: 2, outlineNumber: '3.1', start: d('2027-02-01'), finish: d('2027-02-12', true), critical: false, totalSlack: 4 * DAY, preds: [{ uid: 15, type: 1, lag: 0, lagFormat: 7 }] },
    { uid: 18, name: 'Environmental tests', level: 2, outlineNumber: '3.2', start: d('2027-02-15'), finish: d('2027-03-05', true), critical: true, totalSlack: 0, preds: [{ uid: 17, type: 1, lag: 0, lagFormat: 7 }] },
    { uid: 19, name: 'Functional tests', level: 2, outlineNumber: '3.3', start: d('2027-02-15'), finish: d('2027-02-26', true), critical: false, totalSlack: 5 * DAY, preds: [{ uid: 17, type: 1, lag: 0, lagFormat: 7 }] },
    { uid: 20, name: 'Test report', level: 2, outlineNumber: '3.4', start: d('2027-03-08'), finish: d('2027-03-12', true), critical: true, totalSlack: 0, preds: [{ uid: 18, type: 1, lag: 0, lagFormat: 7 }, { uid: 19, type: 1, lag: 0, lagFormat: 7 }] },
    { uid: 21, name: 'Validation complete', level: 2, outlineNumber: '3.5', start: d('2027-03-12'), finish: d('2027-03-12'), milestone: true, critical: true, totalSlack: 0, preds: [{ uid: 20, type: 1, lag: 0, lagFormat: 7 }] },
  ]
  return {
    name: 'EVT Program',
    minutesPerDay: 480,
    tasks,
    resources: [
      { uid: 1, name: 'Avery Chen' },
      { uid: 2, name: 'Jordan Patel' },
      { uid: 3, name: 'Sam Okafor' },
      { uid: 4, name: 'Riley Novak' },
    ],
    assignments: [
      { taskUid: 2, resourceUid: 1 },
      { taskUid: 3, resourceUid: 1 },
      { taskUid: 3, resourceUid: 2 },
      { taskUid: 4, resourceUid: 2 },
      { taskUid: 5, resourceUid: 1 },
      { taskUid: 8, resourceUid: 3 },
      { taskUid: 9, resourceUid: 3 },
      { taskUid: 10, resourceUid: 3 },
      { taskUid: 12, resourceUid: 4 },
      { taskUid: 13, resourceUid: 4 },
      { taskUid: 14, resourceUid: 3 },
      { taskUid: 14, resourceUid: 4 },
      { taskUid: 17, resourceUid: 2 },
      { taskUid: 18, resourceUid: 2 },
      { taskUid: 19, resourceUid: 4 },
      { taskUid: 20, resourceUid: 1 },
    ],
  }
}

export function sampleMspdi(): string {
  return buildMspdi(sampleSpec())
}

/** Writes the synthetic sample to `<dir>/evt-program.xml` (default: test/fixtures/out). Returns the path. */
export function writeSampleMspdi(dir: string = fileURLToPath(new URL('./out', import.meta.url))): string {
  mkdirSync(dir, { recursive: true })
  const path = join(dir, 'evt-program.xml')
  writeFileSync(path, sampleMspdi(), 'utf8')
  return path
}
