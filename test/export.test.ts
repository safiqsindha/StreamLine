import { createServer, type Server } from 'node:http'
import { mkdtemp, readFile, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AddressInfo } from 'node:net'
import { afterAll, beforeAll, describe, expect, it } from 'vitest'
import { createApp } from '../src/server/app'
import { FileStore } from '../src/server/fileStore'
import { exportTheme, svgDocument } from '../src/server/render'
import { renderTimeline } from '../src/core/render/timeline'
import { computeRange, createLayout } from '../src/core/layout'
import { resolveView } from '../src/core/views'
import { emptyPresentation, type Schedule, type Task } from '../src/core/model'
import type { ParsedSchedule } from '../src/core/import/types'

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

const tasks = [task('1'), task('2', { start: '2026-01-12', finish: '2026-01-23', deps: [{ predecessorId: '1', type: 'FS', lagDays: 0 }] }), task('3', { milestone: true, start: '2026-02-02', finish: '2026-02-02' })]

const fakeParse = async (_bytes: Uint8Array, filename: string): Promise<ParsedSchedule> => ({
  kind: 'mspdi',
  filename,
  name: 'Program Alpha',
  workstreams: [{ id: 'w1', name: 'Avionics' }],
  tasks,
  detectedColumns: [],
  mapping: {},
  unmapped: [],
  issues: [],
})

let dir: string
let server: Server
let base: string
let schedule: Schedule
let css: string

const exportReq = (body: unknown) =>
  fetch(`${base}/api/schedules/${schedule.id}/export`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify(body) })

beforeAll(async () => {
  css = await readFile(join(__dirname, '..', 'src', 'web', 'theme.css'), 'utf8')
  dir = await mkdtemp(join(tmpdir(), 'streamline-export-'))
  server = createServer(createApp(new FileStore(dir), { parse: fakeParse, themeCss: css }))
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', r))
  base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`
  const prev = await (await fetch(`${base}/api/schedules/import/preview`, { method: 'POST', headers: { 'X-Filename': 'a.xml' }, body: Buffer.from([1]) })).json()
  const acc = await fetch(`${base}/api/schedules/import/accept`, { method: 'POST', headers: { 'Content-Type': 'application/json' }, body: JSON.stringify({ previewId: prev.previewId, name: 'Program Alpha' }) })
  schedule = await acc.json()
})

afterAll(async () => {
  await new Promise((r) => server.close(r))
  await rm(dir, { recursive: true, force: true })
})

function expected(view: 'detailed' | 'critical', today: string): string {
  const presentation = emptyPresentation(schedule.id)
  const resolved = resolveView(schedule, presentation, { view })
  const dates: string[] = []
  for (const row of resolved.rows) {
    if (row.kind === 'task') {
      const t = schedule.tasks.find((x) => x.id === row.taskId)!
      dates.push(t.start!, t.finish!)
    } else if (row.kind === 'milestones') {
      for (const id of row.taskIds) dates.push(schedule.tasks.find((x) => x.id === id)!.start!)
    }
  }
  const range = computeRange(dates, 'week')!
  const layout = createLayout(range, 'week', resolved.rows.length)
  const theme = exportTheme(css)
  return renderTimeline({
    schedule,
    presentation,
    resolved,
    layout,
    showDeps: false,
    today,
    color: theme.color,
    font: theme.font,
    labels: { width: 250, title: `${schedule.name} · ${view === 'critical' ? 'Critical path' : 'Detailed'}` },
  })
}

describe('export', () => {
  it('svg is the renderTimeline output for the same state', async () => {
    const res = await exportReq({ format: 'svg', view: 'detailed', today: '2026-01-10' })
    expect(res.status).toBe(200)
    expect(res.headers.get('content-type')).toContain('image/svg+xml')
    expect(res.headers.get('content-disposition')).toBe('attachment; filename="program-alpha-detailed.svg"')
    const text = await res.text()
    expect(text.startsWith('<?xml version="1.0" encoding="UTF-8"?>\n<svg')).toBe(true)
    expect(text).toBe(svgDocument(expected('detailed', '2026-01-10')))
  })

  it('png has the signature and the SVG size at 2x', async () => {
    const svg = await (await exportReq({ format: 'svg', today: '2026-01-10' })).text()
    const w = Number(/<svg[^>]* width="([\d.]+)"/.exec(svg)![1])
    const h = Number(/<svg[^>]* height="([\d.]+)"/.exec(svg)![1])
    const res = await exportReq({ format: 'png', today: '2026-01-10' })
    expect(res.headers.get('content-type')).toBe('image/png')
    expect(res.headers.get('content-disposition')).toContain('program-alpha-detailed.png')
    const png = Buffer.from(await res.arrayBuffer())
    expect([...png.subarray(0, 8)]).toEqual([0x89, 0x50, 0x4e, 0x47, 0x0d, 0x0a, 0x1a, 0x0a])
    expect(png.subarray(12, 16).toString('ascii')).toBe('IHDR')
    expect(png.readUInt32BE(16)).toBe(Math.round(w * 2))
    expect(png.readUInt32BE(20)).toBe(Math.round(h * 2))
  })

  it('pdf is well formed with correct xref offsets', async () => {
    const res = await exportReq({ format: 'pdf', today: '2026-01-10' })
    expect(res.headers.get('content-type')).toBe('application/pdf')
    const pdf = Buffer.from(await res.arrayBuffer())
    const text = pdf.toString('latin1')
    expect(text.startsWith('%PDF-1.4')).toBe(true)
    expect(text.trimEnd().endsWith('%%EOF')).toBe(true)

    const startxref = Number(/startxref\n(\d+)\n%%EOF/.exec(text)![1])
    expect(text.slice(startxref, startxref + 4)).toBe('xref')
    const entries = [...text.slice(startxref).matchAll(/^(\d{10}) (\d{5}) ([nf]) $/gm)]
    expect(entries.length).toBe(6)
    entries.slice(1).forEach((m, i) => {
      const off = Number(m[1])
      expect(text.slice(off, off + `${i + 1} 0 obj`.length)).toBe(`${i + 1} 0 obj`)
    })
    expect(text).toContain('/Filter /FlateDecode')
    expect(text).toContain('/ColorSpace /DeviceRGB')
  })

  it('empty critical view is 422 with a plain message', async () => {
    for (const format of ['svg', 'png', 'pdf']) {
      const res = await exportReq({ format, view: 'critical' })
      expect(res.status).toBe(422)
      expect((await res.json()).error.message).toMatch(/did not say which tasks are critical/)
    }
  })

  it('rejects bad input with 400', async () => {
    expect((await exportReq({ format: 'gif' })).status).toBe(400)
    expect((await exportReq({})).status).toBe(400)
    expect((await exportReq({ format: 'svg', view: 'nope' })).status).toBe(400)
    expect((await exportReq({ format: 'svg', zoom: 3 })).status).toBe(400)
    expect((await exportReq({ format: 'svg', today: 'yesterday' })).status).toBe(400)
    expect((await exportReq({ format: 'svg', filters: { groups: 'a' } })).status).toBe(400)
  })

  it('keeps csv and json naming', async () => {
    const res = await exportReq({ format: 'csv' })
    expect(res.headers.get('content-disposition')).toBe('attachment; filename="program-alpha.csv"')
  })
})
