import { createServer, type Server } from 'node:http'
import { mkdtemp, rm } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import type { AddressInfo } from 'node:net'
import { afterAll, beforeAll, beforeEach, describe, expect, it } from 'vitest'
import { createApp } from '../src/server/app'
import { FileStore } from '../src/server/fileStore'
import { ImportError, type ParsedSchedule } from '../src/core/import/types'
import type { Task } from '../src/core/model'

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

let nextTasks: Task[] = []
const fakeParse = async (_bytes: Uint8Array, filename: string): Promise<ParsedSchedule> => {
  if (filename.endsWith('.bad')) throw new ImportError('This file is not a schedule.')
  return {
    kind: 'mspdi',
    filename,
    name: 'Program Alpha',
    workstreams: [{ id: 'w1', name: 'Avionics' }],
    tasks: nextTasks,
    detectedColumns: [],
    mapping: {},
    unmapped: [],
    issues: [],
  }
}

let dir: string
let server: Server
let base: string

beforeAll(async () => {
  dir = await mkdtemp(join(tmpdir(), 'streamline-'))
  server = createServer(createApp(new FileStore(dir), { parse: fakeParse, maxBodyBytes: 1024 }))
  await new Promise<void>((r) => server.listen(0, '127.0.0.1', r))
  base = `http://127.0.0.1:${(server.address() as AddressInfo).port}`
})

afterAll(async () => {
  await new Promise((r) => server.close(r))
  await rm(dir, { recursive: true, force: true })
})

beforeEach(() => {
  nextTasks = [task('1'), task('2', { name: 'Has, comma', owner: 'Sam "S"', deps: [{ predecessorId: '1', type: 'FS', lagDays: 2 }] }), task('3', { start: null, finish: null })]
})

async function call(method: string, path: string, body?: unknown, headers: Record<string, string> = {}) {
  const res = await fetch(base + path, {
    method,
    headers: body !== undefined && !(body instanceof Uint8Array) ? { 'Content-Type': 'application/json', ...headers } : headers,
    body: body === undefined ? undefined : body instanceof Uint8Array ? Buffer.from(body) : JSON.stringify(body),
  })
  const text = await res.text()
  let json: any
  try {
    json = JSON.parse(text)
  } catch {
    json = undefined
  }
  return { status: res.status, json, text, headers: res.headers }
}

async function importNew(filename = 'a.xml', name?: string) {
  const prev = await call('POST', '/api/schedules/import/preview', new Uint8Array([1, 2, 3]), { 'X-Filename': filename })
  expect(prev.status).toBe(200)
  const acc = await call('POST', '/api/schedules/import/accept', { previewId: prev.json.previewId, name })
  expect(acc.status).toBe(201)
  return acc.json
}

describe('server', () => {
  it('lists no schedules at first', async () => {
    const r = await call('GET', '/api/schedules')
    expect(r.status).toBe(200)
    expect(r.json).toEqual([])
  })

  it('previews, accepts and fetches a schedule', async () => {
    const prev = await call('POST', '/api/schedules/import/preview', new Uint8Array([1]), { 'X-Filename': 'a.xml' })
    expect(prev.json.summary).toMatchObject({ rowCount: 3, laneCount: 1, dateRange: { start: '2026-01-05', finish: '2026-01-09' }, hasCriticality: false, hasFloat: false })
    expect(prev.json.reconciliation).toBeUndefined()
    const acc = await call('POST', '/api/schedules/import/accept', { previewId: prev.json.previewId })
    expect(acc.json.id).toMatch(/^sch_/)
    expect(acc.json.name).toBe('Program Alpha')
    expect(acc.json.source).toMatchObject({ kind: 'mspdi', filename: 'a.xml' })
    const got = await call('GET', `/api/schedules/${acc.json.id}`)
    expect(got.json.tasks).toHaveLength(3)
    const list = await call('GET', '/api/schedules')
    expect(list.json.find((m: any) => m.id === acc.json.id).taskCount).toBe(3)
    const again = await call('POST', '/api/schedules/import/accept', { previewId: prev.json.previewId })
    expect(again.status).toBe(410)
  })

  it('reports import problems in plain language', async () => {
    expect((await call('POST', '/api/schedules/import/preview', new Uint8Array([1]))).status).toBe(400)
    const bad = await call('POST', '/api/schedules/import/preview', new Uint8Array([1]), { 'X-Filename': 'x.bad' })
    expect(bad.status).toBe(422)
    expect(bad.json.error.message).toBe('This file is not a schedule.')
    const big = await call('POST', '/api/schedules/import/preview', new Uint8Array(2048), { 'X-Filename': 'a.xml' })
    expect(big.status).toBe(413)
  })

  it('renames a schedule', async () => {
    const s = await importNew()
    const r = await call('PATCH', `/api/schedules/${s.id}`, { name: '  Renamed ' })
    expect(r.json).toMatchObject({ id: s.id, name: 'Renamed', taskCount: 3 })
    expect((await call('PATCH', `/api/schedules/${s.id}`, { name: '' })).status).toBe(400)
    expect((await call('GET', `/api/schedules/${s.id}`)).json.tasks).toEqual(s.tasks)
  })

  it('patches, merges and clears presentation', async () => {
    const s = await importNew()
    const p = `/api/schedules/${s.id}/presentation`
    expect((await call('GET', p)).json.tasks).toEqual({})
    await call('PATCH', p, { theme: 'print', tasks: { '1': { displayLabel: 'One', emphasis: 'highlight' } } })
    const merged = await call('PATCH', p, { tasks: { '1': { order: 3 } } })
    expect(merged.json.tasks['1']).toEqual({ displayLabel: 'One', emphasis: 'highlight', order: 3 })
    expect(merged.json.theme).toBe('print')
    const cleared = await call('PATCH', p, { theme: null, tasks: { '1': { displayLabel: '', emphasis: null } } })
    expect(cleared.json.tasks['1']).toEqual({ order: 3 })
    expect(cleared.json.theme).toBeUndefined()
    const removed = await call('PATCH', p, { tasks: { '1': null } })
    expect(removed.json.tasks).toEqual({})
    expect((await call('GET', p)).json.tasks).toEqual({})
  })

  it('validates presentation types', async () => {
    const s = await importNew()
    const p = `/api/schedules/${s.id}/presentation`
    expect((await call('PATCH', p, { tasks: { '1': { emphasis: 'loud' } } })).status).toBe(400)
    expect((await call('PATCH', p, { tasks: { '1': { order: 'x' } } })).status).toBe(400)
    expect((await call('PATCH', p, { groupOrder: 'x' })).status).toBe(400)
    expect((await call('PATCH', p, { bogus: 1 })).status).toBe(400)
  })

  it('rejects source fields', async () => {
    const s = await importNew()
    const p = `/api/schedules/${s.id}/presentation`
    const msg = 'Source fields are read-only. Change them in the source file and re-import.'
    for (const field of ['start', 'finish', 'owner', 'name', 'id', 'deps', 'critical', 'workstream']) {
      const r = await call('PATCH', p, { tasks: { '1': { [field]: 'x' } } })
      expect(r.status).toBe(400)
      expect(r.json.error.message).toBe(msg)
    }
    const top = await call('PATCH', p, { start: '2026-02-01' })
    expect(top.status).toBe(400)
    expect(top.json.error.message).toBe(msg)
    expect((await call('GET', p)).json.tasks).toEqual({})
  })

  it('rejects overrides for unknown tasks', async () => {
    const s = await importNew()
    const r = await call('PATCH', `/api/schedules/${s.id}/presentation`, { tasks: { nope: { displayLabel: 'x' } } })
    expect(r.status).toBe(400)
  })

  it('resets some or all overrides', async () => {
    const s = await importNew()
    const p = `/api/schedules/${s.id}/presentation`
    await call('PATCH', p, { tasks: { '1': { displayLabel: 'a' }, '2': { displayLabel: 'b' } } })
    const some = await call('POST', `${p}/reset`, { taskIds: ['1'] })
    expect(Object.keys(some.json.tasks)).toEqual(['2'])
    const all = await call('POST', `${p}/reset`)
    expect(all.json.tasks).toEqual({})
  })

  it('creates immutable baselines with new ids and deletes them', async () => {
    const s = await importNew()
    const b = `/api/schedules/${s.id}/baselines`
    const one = await call('POST', b, { name: 'Kickoff' })
    expect(one.status).toBe(201)
    expect(one.json.id).toMatch(/^bl_/)
    expect(one.json.tasks[0]).toEqual({ id: '1', name: 'Task 1', start: '2026-01-05', finish: '2026-01-09', milestone: false })
    await new Promise((r) => setTimeout(r, 5))
    const two = await call('POST', b, { name: 'Kickoff' })
    expect(two.json.id).not.toBe(one.json.id)
    const list = await call('GET', b)
    expect(list.json.map((x: any) => x.id)).toEqual([two.json.id, one.json.id])
    expect((await call('POST', b, {})).status).toBe(400)
    expect((await call('DELETE', `${b}/${one.json.id}`)).status).toBe(204)
    expect((await call('DELETE', `${b}/${one.json.id}`)).status).toBe(404)
    expect((await call('GET', b)).json).toHaveLength(1)
  })

  it('refuses to overwrite an existing baseline in the store', async () => {
    const store = new FileStore(await mkdtemp(join(tmpdir(), 'streamline-store-')))
    const b = { id: 'bl_x', scheduleId: 'sch_x', name: 'n', createdAt: new Date().toISOString(), tasks: [] }
    await store.putBaseline('sch_x', b)
    await expect(store.putBaseline('sch_x', { ...b, name: 'changed' })).rejects.toThrow(/already exists/)
    expect((await store.listBaselines('sch_x'))[0].name).toBe('n')
  })

  it('deleting a schedule removes its presentation and baselines', async () => {
    const s = await importNew()
    await call('PATCH', `/api/schedules/${s.id}/presentation`, { tasks: { '1': { displayLabel: 'a' } } })
    await call('POST', `/api/schedules/${s.id}/baselines`, { name: 'b' })
    expect((await call('DELETE', `/api/schedules/${s.id}`)).status).toBe(204)
    expect((await call('GET', `/api/schedules/${s.id}`)).status).toBe(404)
    const store = new FileStore(dir)
    expect(await store.getPresentation(s.id)).toBeNull()
    expect(await store.listBaselines(s.id)).toEqual([])
    expect((await call('GET', '/api/schedules')).json.some((m: any) => m.id === s.id)).toBe(false)
  })

  it('replace import keeps id and presentation and reports orphans', async () => {
    const s = await importNew()
    await call('PATCH', `/api/schedules/${s.id}/presentation`, { tasks: { '1': { displayLabel: 'one' }, '2': { displayLabel: 'two' } } })
    nextTasks = [task('2', { finish: '2026-02-01' }), task('4')]
    const prev = await call('POST', '/api/schedules/import/preview', new Uint8Array([1]), { 'X-Filename': 'b.xml', 'X-Replace-Schedule-Id': s.id })
    expect(prev.json.reconciliation).toEqual({
      targetScheduleId: s.id,
      added: ['4'],
      removed: ['1', '3'],
      changed: [{ id: '2', fields: ['name', 'owner', 'finish', 'deps'] }],
      orphanedOverrides: ['1'],
    })
    const acc = await call('POST', '/api/schedules/import/accept', { previewId: prev.json.previewId })
    expect(acc.status).toBe(200)
    expect(acc.json.id).toBe(s.id)
    expect(acc.json.createdAt).toBe(s.createdAt)
    expect(acc.json.name).toBe(s.name)
    expect(acc.json.source.filename).toBe('b.xml')
    expect(acc.json.tasks.map((t: Task) => t.id)).toEqual(['2', '4'])
    const pres = await call('GET', `/api/schedules/${s.id}/presentation`)
    expect(Object.keys(pres.json.tasks).sort()).toEqual(['1', '2'])
    const missing = await call('POST', '/api/schedules/import/preview', new Uint8Array([1]), { 'X-Filename': 'b.xml', 'X-Replace-Schedule-Id': 'sch_missing' })
    expect(missing.status).toBe(404)
  })

  it('exports csv and json from source data', async () => {
    const s = await importNew('a.xml', 'My Program')
    const csv = await call('POST', `/api/schedules/${s.id}/export`, { format: 'csv' })
    expect(csv.status).toBe(200)
    expect(csv.headers.get('content-type')).toContain('text/csv')
    expect(csv.headers.get('content-disposition')).toContain('My_Program.csv')
    const lines = csv.text.trim().split('\r\n')
    expect(lines[0]).toBe('id,name,workstream,owner,start,finish,milestone,critical,float_days,parent_id,is_summary,predecessors')
    expect(lines[1]).toBe('1,Task 1,Avionics,,2026-01-05,2026-01-09,false,unavailable,,,false,')
    expect(lines[2]).toBe('2,"Has, comma",Avionics,"Sam ""S""",2026-01-05,2026-01-09,false,unavailable,,,false,1:FS:2')
    expect(lines[3]).toBe('3,Task 3,Avionics,,,,false,unavailable,,,false,')
    const json = await call('POST', `/api/schedules/${s.id}/export`, { format: 'json' })
    expect(json.headers.get('content-disposition')).toContain('.json')
    expect(json.json.tasks[0].owner).toBeNull()
    const png = await call('POST', `/api/schedules/${s.id}/export`, { format: 'png' })
    expect(png.status).toBe(501)
    expect(png.json.error.message).toBe('Not available yet')
    expect((await call('POST', `/api/schedules/${s.id}/export`, { format: 'doc' })).status).toBe(400)
  })

  it('serves the OpenAPI document covering every route', async () => {
    const r = await call('GET', '/api/openapi.json')
    expect(r.json.openapi).toBe('3.1.0')
    expect(Object.keys(r.json.paths)).toContain('/api/schedules/{id}/export')
    for (const op of Object.values<any>(r.json.paths).flatMap((p) => Object.values<any>(p))) {
      expect(op.summary).toBeTruthy()
      expect(op['x-example']).toBeDefined()
    }
  })

  it('returns JSON errors for unknown routes, methods and ids', async () => {
    const nf = await call('GET', '/api/nothing')
    expect(nf.status).toBe(404)
    expect(nf.json.error.message).toBeTruthy()
    expect((await call('PUT', '/api/schedules')).status).toBe(405)
    expect((await call('GET', '/api/schedules/sch_missing')).status).toBe(404)
    expect((await call('GET', '/api/schedules/sch_missing/baselines')).status).toBe(404)
  })

  it('rejects path-traversal ids', async () => {
    for (const id of ['..%2F..%2Fetc', 'a%5Cb', 'a.b', '%00']) {
      const r = await call('GET', `/api/schedules/${id}`)
      expect(r.status).toBe(400)
      expect(r.text).not.toMatch(/\bat .*\.ts/)
    }
    const s = await importNew()
    expect((await call('DELETE', `/api/schedules/${s.id}/baselines/..%2F..%2Fschedules%2F${s.id}`)).status).toBe(400)
    const store = new FileStore(dir)
    await expect(store.getSchedule('../x')).rejects.toThrow(/not valid/)
    await expect(store.putBaseline('sch_a', { id: '../b', scheduleId: 'sch_a', name: 'x', createdAt: '', tasks: [] })).rejects.toThrow(/not valid/)
  })

  it('does not leak stack traces when the view cannot be built', async () => {
    const s = await importNew()
    const r = await call('GET', `/api/schedules/${s.id}/view?view=bogus`)
    expect(r.status).toBe(400)
    expect(r.text).not.toContain('node_modules')
  })
})
