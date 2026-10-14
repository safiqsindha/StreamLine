import { mkdir, readFile, readdir, rename, rm, writeFile } from 'node:fs/promises'
import { join, resolve } from 'node:path'
import { randomUUID } from 'node:crypto'
import type { Baseline, Presentation, Schedule, ScheduleMeta } from '../core/model'
import { toMeta } from '../core/model'
import { StoreError, type Store } from '../core/store'

const SAFE_ID = /^[A-Za-z0-9_-]{1,128}$/

export function isSafeId(id: string): boolean {
  return SAFE_ID.test(id)
}

function requireId(id: string, what: string): void {
  if (!isSafeId(id)) throw new StoreError(`That ${what} ID is not valid.`)
}

function isMissing(e: unknown): boolean {
  return typeof e === 'object' && e !== null && (e as { code?: string }).code === 'ENOENT'
}

export class FileStore implements Store {
  private readonly root: string

  constructor(dataDir = './data') {
    this.root = resolve(dataDir)
  }

  private schedulePath(id: string) {
    return join(this.root, 'schedules', `${id}.json`)
  }
  private presentationPath(id: string) {
    return join(this.root, 'presentation', `${id}.json`)
  }
  private baselineDir(scheduleId: string) {
    return join(this.root, 'baselines', scheduleId)
  }

  private async readJson<T>(path: string, what: string): Promise<T | null> {
    try {
      return JSON.parse(await readFile(path, 'utf8')) as T
    } catch (e) {
      if (isMissing(e)) return null
      throw new StoreError(`Streamline could not read the saved ${what}. Check that the data folder is readable.`, e)
    }
  }

  private async writeJson(path: string, value: unknown, what: string, exclusive = false): Promise<void> {
    const tmp = `${path}.${randomUUID()}.tmp`
    try {
      await mkdir(join(path, '..'), { recursive: true })
      await writeFile(tmp, JSON.stringify(value, null, 2), 'utf8')
      if (exclusive) {
        const exists = await readFile(path).then(
          () => true,
          (e) => {
            if (isMissing(e)) return false
            throw e
          },
        )
        if (exists) {
          await rm(tmp, { force: true })
          throw new StoreError(`A ${what} with this ID already exists and cannot be changed.`)
        }
      }
      await rename(tmp, path)
    } catch (e) {
      await rm(tmp, { force: true }).catch(() => undefined)
      if (e instanceof StoreError) throw e
      throw new StoreError(`Streamline could not save the ${what}. Check that the data folder is writable and has free space.`, e)
    }
  }

  private async listDir(dir: string): Promise<string[]> {
    try {
      return await readdir(dir)
    } catch (e) {
      if (isMissing(e)) return []
      throw new StoreError('Streamline could not read the data folder.', e)
    }
  }

  async listSchedules(): Promise<ScheduleMeta[]> {
    const files = (await this.listDir(join(this.root, 'schedules'))).filter((f) => f.endsWith('.json'))
    const metas: ScheduleMeta[] = []
    for (const f of files) {
      const s = await this.readJson<Schedule>(join(this.root, 'schedules', f), 'schedule')
      if (s) metas.push(toMeta(s))
    }
    return metas.sort((a, b) => (a.createdAt < b.createdAt ? 1 : a.createdAt > b.createdAt ? -1 : a.id.localeCompare(b.id)))
  }

  async getSchedule(id: string): Promise<Schedule | null> {
    requireId(id, 'schedule')
    return this.readJson<Schedule>(this.schedulePath(id), 'schedule')
  }

  async putSchedule(s: Schedule): Promise<void> {
    requireId(s.id, 'schedule')
    await this.writeJson(this.schedulePath(s.id), s, 'schedule')
  }

  async deleteSchedule(id: string): Promise<void> {
    requireId(id, 'schedule')
    try {
      await rm(this.schedulePath(id), { force: true })
      await rm(this.presentationPath(id), { force: true })
      await rm(this.baselineDir(id), { recursive: true, force: true })
    } catch (e) {
      throw new StoreError('Streamline could not delete the schedule. Check that the data folder is writable.', e)
    }
  }

  async getPresentation(scheduleId: string): Promise<Presentation | null> {
    requireId(scheduleId, 'schedule')
    return this.readJson<Presentation>(this.presentationPath(scheduleId), 'presentation')
  }

  async putPresentation(scheduleId: string, p: Presentation): Promise<void> {
    requireId(scheduleId, 'schedule')
    await this.writeJson(this.presentationPath(scheduleId), p, 'presentation')
  }

  async listBaselines(scheduleId: string): Promise<Baseline[]> {
    requireId(scheduleId, 'schedule')
    const dir = this.baselineDir(scheduleId)
    const files = (await this.listDir(dir)).filter((f) => f.endsWith('.json'))
    const out: Baseline[] = []
    for (const f of files) {
      const b = await this.readJson<Baseline>(join(dir, f), 'baseline')
      if (b) out.push(b)
    }
    return out.sort((a, b) => (a.createdAt < b.createdAt ? 1 : a.createdAt > b.createdAt ? -1 : b.id.localeCompare(a.id)))
  }

  async putBaseline(scheduleId: string, b: Baseline): Promise<void> {
    requireId(scheduleId, 'schedule')
    requireId(b.id, 'baseline')
    await this.writeJson(join(this.baselineDir(scheduleId), `${b.id}.json`), b, 'baseline', true)
  }

  async deleteBaseline(scheduleId: string, baselineId: string): Promise<void> {
    requireId(scheduleId, 'schedule')
    requireId(baselineId, 'baseline')
    try {
      await rm(join(this.baselineDir(scheduleId), `${baselineId}.json`), { force: true })
    } catch (e) {
      throw new StoreError('Streamline could not delete the baseline. Check that the data folder is writable.', e)
    }
  }
}
