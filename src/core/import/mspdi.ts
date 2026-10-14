import { XMLParser, XMLValidator } from 'fast-xml-parser'
import { isIsoDate } from '../dates'
import type { Dependency, DependencyType, Criticality, Task, Workstream } from '../model'
import { ImportError, type ColumnMapping, type ImportIssue, type ParsedSchedule } from './types'

const NOT_MSPDI = 'This .xml is not an MSPDI export. In Project, use File → Save As → XML.'
const XMLNS = /^http:\/\/schemas\.microsoft\.com\/project(\/.*)?$/i

// Project's schema default for <MinutesPerDay> (8 hours). Used only when the file omits the element.
const DEFAULT_MINUTES_PER_DAY = 480
const MINUTES_PER_ELAPSED_DAY = 1440

// Duration format codes used by <LagFormat>. Elapsed formats count calendar time, not working time.
const ELAPSED_FORMATS = new Set([4, 6, 8, 10, 12, 36, 38, 40, 42, 44])
const PERCENT_FORMATS = new Set([19, 20, 51, 52])

const LINK_TYPES: Record<string, DependencyType> = { '0': 'FF', '1': 'FS', '2': 'SF', '3': 'SS' }

// Task elements this importer reads. Everything else the file carries is reported as unmapped.
const USED_TASK_ELEMENTS = [
  'UID',
  'Name',
  'Start',
  'Finish',
  'Milestone',
  'Summary',
  'Critical',
  'TotalSlack',
  'OutlineLevel',
  'OutlineNumber',
  'PredecessorLink',
]

// MSPDI has no workstream field. A task's workstream is the name of its level-1 summary
// ancestor (or its own name when it is that summary). Level-1 tasks that are not summaries have none.
const MAPPING: ColumnMapping = {
  id: 'UID',
  name: 'Name',
  workstream: 'Name of the OutlineLevel 1 summary task',
  owner: 'Assignment → Resource Name',
  start: 'Start',
  finish: 'Finish',
  milestone: 'Milestone',
  critical: 'Critical',
  float: 'TotalSlack',
  parent: 'OutlineLevel',
  summary: 'Summary',
  predecessors: 'PredecessorLink',
}

type Node = Record<string, unknown>

interface RawLink {
  uid: string
  type: DependencyType
  lagDays: number | null
}

interface Entry {
  level: number
  kept: boolean
  uid: string
  name: string | null
}

const parser = new XMLParser({
  ignoreAttributes: false,
  attributeNamePrefix: '@_',
  parseTagValue: false,
  parseAttributeValue: false,
  trimValues: true,
  isArray: (name) => name === 'Task' || name === 'PredecessorLink' || name === 'Assignment' || name === 'Resource',
})

function isNode(v: unknown): v is Node {
  return typeof v === 'object' && v !== null && !Array.isArray(v)
}

function list(v: unknown): Node[] {
  if (v == null || v === '') return []
  return (Array.isArray(v) ? v : [v]).filter(isNode)
}

function text(n: Node, key: string): string | undefined {
  const v = n[key]
  const s = isNode(v) ? v['#text'] : v
  if (s == null || typeof s === 'object') return undefined
  const t = String(s).trim()
  return t === '' ? undefined : t
}

function flag(n: Node, key: string): boolean | undefined {
  const t = text(n, key)?.toLowerCase()
  if (t === '1' || t === 'true') return true
  if (t === '0' || t === 'false') return false
  return undefined
}

function int(n: Node, key: string): number | undefined {
  const t = text(n, key)
  if (t === undefined || !/^-?\d+$/.test(t)) return undefined
  return Number(t)
}

function datePart(v: string | undefined): string | null {
  if (!v) return null
  const d = v.slice(0, 10)
  return isIsoDate(d) ? d : null
}

function round(n: number, places: number): number {
  const f = 10 ** places
  return Math.round(n * f) / f
}

function levelOf(t: Node): number {
  const level = int(t, 'OutlineLevel')
  if (level !== undefined) return level
  const num = text(t, 'OutlineNumber')
  return num ? num.split('.').length : 1
}

export function parseMspdi(xml: string, filename: string): ParsedSchedule {
  const src = xml.charCodeAt(0) === 0xfeff ? xml.slice(1) : xml
  if (XMLValidator.validate(src) !== true) {
    throw new ImportError('This file is not valid XML and could not be read.')
  }
  let doc: Node
  try {
    doc = parser.parse(src) as Node
  } catch {
    throw new ImportError('This file is not valid XML and could not be read.')
  }
  const project = doc['Project']
  if (!isNode(project) || !XMLNS.test(String(project['@_xmlns'] ?? ''))) throw new ImportError(NOT_MSPDI)

  const minutesPerDay = int(project, 'MinutesPerDay')
  const mpd = minutesPerDay && minutesPerDay > 0 ? minutesPerDay : DEFAULT_MINUTES_PER_DAY

  const projectName = text(project, 'Name') ?? text(project, 'Title') ?? filename.replace(/\.[^./\\]*$/, '')

  const resourceNames = new Map<string, { name: string; order: number }>()
  for (const r of list(isNode(project['Resources']) ? project['Resources']['Resource'] : undefined)) {
    const uid = text(r, 'UID')
    const name = text(r, 'Name')
    if (uid && name) resourceNames.set(uid, { name, order: Number(uid) })
  }
  const owners = new Map<string, Map<string, { name: string; order: number }>>()
  for (const a of list(isNode(project['Assignments']) ? project['Assignments']['Assignment'] : undefined)) {
    const taskUid = text(a, 'TaskUID')
    const resUid = text(a, 'ResourceUID')
    const res = resUid ? resourceNames.get(resUid) : undefined
    if (!taskUid || !resUid || !res) continue
    let m = owners.get(taskUid)
    if (!m) owners.set(taskUid, (m = new Map()))
    m.set(resUid, res)
  }

  const issues: ImportIssue[] = []
  const detected = new Set<string>()
  const tasks: Task[] = []
  const links = new Map<string, RawLink[]>()
  const linkLabels = new Map<string, string>()
  const seen = new Set<string>()
  const stack: Entry[] = []

  const rawTasks = list(isNode(project['Tasks']) ? project['Tasks']['Task'] : undefined)
  rawTasks.forEach((t, index) => {
    const row = index + 1
    for (const key of Object.keys(t)) if (!key.startsWith('@_') && key !== '#text') detected.add(key)

    const uid = text(t, 'UID')
    const level = levelOf(t)
    if (uid === '0' || level === 0) return

    const name = text(t, 'Name') ?? null
    while (stack.length && stack[stack.length - 1].level >= level) stack.pop()
    const entry: Entry = { level, kept: false, uid: uid ?? '', name }
    const ancestors = [...stack]
    stack.push(entry)

    if (!uid) {
      issues.push({ row, message: `Row ${row} has no task UID. That task was skipped.`, severity: 'skipped' })
      return
    }
    if (seen.has(uid)) {
      issues.push({ row, message: `Task ${uid} appears more than once. The repeat was skipped.`, severity: 'skipped' })
      return
    }
    if (!name) {
      issues.push({ row, message: `Task ${uid} has no name. That task was skipped.`, severity: 'skipped' })
      return
    }
    const start = datePart(text(t, 'Start'))
    const finish = datePart(text(t, 'Finish'))
    if (start && finish && finish < start) {
      issues.push({
        row,
        message: `Task ${uid} ("${name}") has a finish date before its start date. That task was skipped.`,
        severity: 'skipped',
      })
      return
    }
    for (const [field, raw, parsed] of [
      ['Start', text(t, 'Start'), start],
      ['Finish', text(t, 'Finish'), finish],
    ] as const) {
      if (raw && !parsed) {
        issues.push({
          row,
          message: `Task ${uid} ("${name}") has a ${field} value Streamline could not read ("${raw}"). It was left blank.`,
          severity: 'warning',
        })
      }
    }

    entry.kept = true
    seen.add(uid)

    const parent = [...ancestors].reverse().find((a) => a.kept)
    const root = level === 1 ? entry : ancestors[0]?.level === 1 ? ancestors[0] : null
    const isSummary = flag(t, 'Summary') === true
    let workstream: string | null = null
    if (root && (root !== entry || isSummary)) workstream = root.name

    const crit = flag(t, 'Critical')
    const critical: Criticality = crit === undefined ? 'unavailable' : crit ? 'critical' : 'not-critical'

    const slack = int(t, 'TotalSlack')
    const floatDays = slack === undefined ? null : round(slack / 10 / mpd, 1)

    const assigned = owners.get(uid)
    const owner = assigned
      ? [...new Set([...assigned.values()].sort((a, b) => a.order - b.order).map((r) => r.name))].join(', ')
      : null

    const rawLinks: RawLink[] = []
    for (const p of list(t['PredecessorLink'])) {
      const pid = text(p, 'PredecessorUID')
      if (!pid) continue
      const type = LINK_TYPES[text(p, 'Type') ?? ''] ?? 'unknown'
      const lag = int(p, 'LinkLag')
      let lagDays: number | null = null
      if (lag !== undefined) {
        const format = int(p, 'LagFormat')
        if (format === undefined) {
          issues.push({
            row,
            message: `Task ${uid} ("${name}") has a lag from task ${pid} without a lag format, so Streamline cannot tell working days from calendar days. The lag is shown as unavailable.`,
            severity: 'warning',
          })
        } else if (PERCENT_FORMATS.has(format)) {
          issues.push({
            row,
            message: `Task ${uid} ("${name}") has a lag from task ${pid} stated as a percentage of duration. Streamline cannot convert that to days, so the lag is shown as unavailable.`,
            severity: 'warning',
          })
        } else {
          const perDay = ELAPSED_FORMATS.has(format) ? MINUTES_PER_ELAPSED_DAY : mpd
          lagDays = round(lag / 10 / perDay, 2)
        }
      }
      rawLinks.push({ uid: pid, type, lagDays })
    }
    links.set(uid, rawLinks)
    linkLabels.set(uid, name)

    tasks.push({
      id: uid,
      name,
      workstream,
      owner,
      start,
      finish,
      milestone: flag(t, 'Milestone') === true,
      critical,
      floatDays,
      parentId: parent ? parent.uid : null,
      isSummary,
      deps: [],
    })
  })

  for (const task of tasks) {
    const deps: Dependency[] = []
    for (const l of links.get(task.id) ?? []) {
      if (!seen.has(l.uid)) {
        issues.push({
          message: `Task ${task.id} ("${linkLabels.get(task.id)}") depends on task ${l.uid}, which is not in the file. That link was dropped.`,
          severity: 'warning',
        })
        continue
      }
      deps.push({ predecessorId: l.uid, type: l.type, lagDays: l.lagDays })
    }
    task.deps = deps
  }

  const workstreams: Workstream[] = []
  for (const t of tasks) {
    if (t.workstream && !workstreams.some((w) => w.id === t.workstream)) {
      workstreams.push({ id: t.workstream, name: t.workstream })
    }
  }

  const used = new Set(USED_TASK_ELEMENTS)
  const detectedColumns = [...detected]
  return {
    kind: 'mspdi',
    filename,
    name: projectName,
    workstreams,
    tasks,
    detectedColumns,
    mapping: { ...MAPPING },
    unmapped: detectedColumns.filter((c) => !used.has(c)),
    issues,
  }
}
