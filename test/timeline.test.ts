import { readFileSync } from 'node:fs'
import { describe, expect, it } from 'vitest'
import { computeRange, createLayout } from '../src/core/layout'
import type { Zoom } from '../src/core/layout'
import { emptyPresentation } from '../src/core/model'
import type { Baseline, Criticality, Dependency, Presentation, Schedule, Task } from '../src/core/model'
import { parseThemeTokens, tokenResolver } from '../src/core/render/theme'
import { renderTimeline } from '../src/core/render/timeline'
import type { ColorToken, TimelineOptions } from '../src/core/render/timeline'
import { resolveView } from '../src/core/views'
import type { ResolvedView, ViewId } from '../src/core/views'

function task(id: string, o: Partial<Task> = {}): Task {
  return {
    id,
    name: `Task ${id}`,
    workstream: 'a',
    owner: null,
    start: '2026-10-05',
    finish: '2026-10-09',
    milestone: false,
    critical: 'unavailable',
    floatDays: null,
    parentId: null,
    isSummary: false,
    deps: [],
    ...o,
  }
}

const dep = (predecessorId: string, type: Dependency['type'] = 'FS', lagDays: number | null = null): Dependency => ({ predecessorId, type, lagDays })

function schedule(tasks: Task[]): Schedule {
  return {
    id: 's1',
    name: 'S',
    createdAt: '2026-10-01T00:00:00Z',
    source: { kind: 'xlsx', filename: 's.xlsx', importedAt: '2026-10-01T00:00:00Z' },
    workstreams: [
      { id: 'a', name: 'Alpha' },
      { id: 'b', name: 'Beta' },
    ],
    tasks,
  }
}

interface Setup {
  s: Schedule
  p?: Presentation
  view?: ViewId
  zoom?: Zoom
  extra?: Partial<TimelineOptions>
  mutate?: (r: ResolvedView) => ResolvedView
  baselineDates?: Array<string | null>
}

function build({ s, p, view = 'detailed', zoom = 'week', extra, mutate, baselineDates = [] }: Setup) {
  const pres = p ?? emptyPresentation('s1')
  let resolved = resolveView(s, pres, { view })
  if (mutate) resolved = mutate(resolved)
  const range = computeRange([...s.tasks.flatMap((t) => [t.start, t.finish]), ...baselineDates], zoom)!
  const layout = createLayout(range, zoom, resolved.rows.length)
  const options: TimelineOptions = {
    schedule: s,
    presentation: pres,
    resolved,
    layout,
    showDeps: true,
    today: '2026-10-14',
    color: (t: ColorToken) => `var(--${t})`,
    font: { text: 'var(--font)', mono: 'var(--mono)' },
    ...extra,
  }
  return { svg: renderTimeline(options), layout, resolved, options }
}

const barOf = (svg: string, id: string) => svg.match(new RegExp(`<g class="bar" data-task-id="${id}"[^>]*>[\\s\\S]*?</g>`))![0]
const nums = (d: string) => (d.match(/-?\d+(?:\.\d+)?/g) ?? []).map(Number)
const pathFor = (svg: string, succ: string, pred?: string) => {
  const re = new RegExp(`<path class="dep" data-pred-id="${pred ?? '[^"]*'}" data-succ-id="${succ}"[^>]*>[\\s\\S]*?</path>`)
  return svg.match(re)![0]
}
const dOf = (path: string) => path.match(/ d="([^"]+)"/)![1]

describe('renderTimeline', () => {
  const base = [
    task('p', { name: 'Board fab', start: '2026-10-05', finish: '2026-10-09', critical: 'critical' }),
    task('fs', { name: 'EVT build', start: '2026-10-13', finish: '2026-10-16', critical: 'critical', deps: [dep('p', 'FS', 2)] }),
    task('ss', { start: '2026-10-12', finish: '2026-10-15', deps: [dep('p', 'SS', null)] }),
    task('ff', { start: '2026-10-19', finish: '2026-10-23', deps: [dep('p', 'FF', 0)] }),
    task('sf', { start: '2026-10-20', finish: '2026-10-26', deps: [dep('p', 'SF', -1)] }),
  ]

  it('is deterministic', () => {
    const a = build({ s: schedule(base) }).svg
    const b = build({ s: schedule(base) }).svg
    expect(a).toBe(b)
    expect(a.startsWith('<svg xmlns="http://www.w3.org/2000/svg" class="tl-svg"')).toBe(true)
    expect(a).toContain('aria-label="Timeline, Detailed view"')
  })

  it('draws bars at layout.span', () => {
    const { svg, layout } = build({ s: schedule(base) })
    const sp = layout.span('2026-10-13', '2026-10-16')
    expect(barOf(svg, 'fs')).toContain(`<rect x="${sp.x}" y="`)
    expect(barOf(svg, 'fs')).toContain(`width="${sp.width}" height="18" rx="3"`)
  })

  it('fills critical only for source critical', () => {
    const { svg } = build({ s: schedule(base) })
    expect(barOf(svg, 'p')).toContain('fill="var(--critical)"')
    expect(barOf(svg, 'ss')).toContain('fill="var(--task)"')
    expect(barOf(svg, 'ss')).not.toContain('--critical')
    const unavail = build({ s: schedule([task('u', { critical: 'unavailable' as Criticality })]) }).svg
    expect(barOf(unavail, 'u')).not.toContain('--critical')
  })

  it('uses lane colors in workstreams view', () => {
    const s = schedule([task('x', { workstream: 'a' }), task('y', { workstream: 'b' })])
    const { svg } = build({ s, view: 'workstreams' })
    expect(barOf(svg, 'x')).toContain('var(--lane-1)')
    expect(barOf(svg, 'y')).toContain('var(--lane-2)')
  })

  describe('dependency anchors', () => {
    const s = schedule(base)
    const { svg, layout, resolved } = build({ s })
    const rowY = (id: string) => layout.y(resolved.rows.findIndex((r) => r.kind === 'task' && r.taskId === id)) + 18
    const right = (a: string, b: string) => {
      const sp = layout.span(a, b)
      return sp.x + sp.width
    }
    const left = (a: string) => layout.x(a)

    it('FS: predecessor finish to successor start', () => {
      const d = nums(dOf(pathFor(svg, 'fs')))
      expect(d[0]).toBe(right('2026-10-05', '2026-10-09'))
      expect(d[1]).toBe(rowY('p'))
      expect(d[d.length - 1]).toBe(left('2026-10-13'))
      expect(d[d.length - 2]).toBe(rowY('fs'))
      expect(d[2]).toBe(d[0] + 8)
      expect(pathFor(svg, 'fs')).toContain('marker-end="url(#tl-detailed-ah-crit)"')
    })

    it('SS: predecessor start to successor start', () => {
      const d = nums(dOf(pathFor(svg, 'ss')))
      expect(d[0]).toBe(left('2026-10-05'))
      expect(d[2]).toBe(d[0] - 8)
      expect(d[d.length - 1]).toBe(left('2026-10-12'))
    })

    it('FF: predecessor finish to successor finish, entering the right edge', () => {
      const d = nums(dOf(pathFor(svg, 'ff')))
      expect(d[0]).toBe(right('2026-10-05', '2026-10-09'))
      const tx = right('2026-10-19', '2026-10-23')
      expect(d[d.length - 1]).toBe(tx)
      expect(d[d.length - 3]).toBe(tx + 8)
    })

    it('SF: predecessor start to successor finish', () => {
      const d = nums(dOf(pathFor(svg, 'sf')))
      expect(d[0]).toBe(left('2026-10-05'))
      expect(d[d.length - 1]).toBe(right('2026-10-20', '2026-10-26'))
    })

    it('routes through the row boundary next to the target row and never starts under the source bar', () => {
      const d = nums(dOf(pathFor(svg, 'fs')))
      const gy = d[3]
      expect(gy).toBe(layout.y(resolved.rows.findIndex((r) => r.kind === 'task' && r.taskId === 'fs')))
      expect(d[2]).toBeGreaterThan(d[0])
    })

    it('puts lag in the title', () => {
      expect(pathFor(svg, 'fs')).toContain('<title>Board fab → EVT build · Finish-to-start · lag +2d</title>')
      expect(pathFor(svg, 'ss')).toContain('Start-to-start · lag not stated')
      expect(pathFor(svg, 'sf')).toContain('Start-to-finish · lag -1d')
    })

    it('draws nothing in the executive view and hides when showDeps is off', () => {
      expect(build({ s, view: 'executive' }).svg).not.toContain('class="dep"')
      expect(build({ s, extra: { showDeps: false } }).svg).not.toContain('class="dep"')
      expect(build({ s, view: 'dependencies', extra: { showDeps: false } }).svg).toContain('class="dep"')
    })
  })

  it('draws unknown dependencies dotted, without an arrowhead, labelled unknown', () => {
    const s = schedule([task('p'), task('u', { start: '2026-10-13', finish: '2026-10-16', deps: [dep('p', 'unknown', 3)] })])
    const { svg, layout } = build({ s })
    const path = pathFor(svg, 'u')
    expect(path).toContain('stroke-dasharray="2 3"')
    expect(path).toContain('stroke="var(--dep-unknown)"')
    expect(path).toContain('marker-end="url(#tl-detailed-oc-unk)"')
    expect(path).not.toContain('-ah-')
    expect(path).toContain('type unknown (not stated in source)')
    expect(path).not.toContain('Finish-to-start')
    const d = nums(dOf(path))
    const sp = layout.span('2026-10-05', '2026-10-09')
    expect(d[0]).toBe(sp.x + sp.width)
    expect(d[d.length - 1]).toBe(layout.x('2026-10-13'))
  })

  it('renders missing dates without inventing them', () => {
    const s = schedule([
      task('n', { start: null, finish: null }),
      task('a', { start: '2026-10-07', finish: null }),
      task('b', { start: null, finish: '2026-10-08' }),
      task('ok'),
    ])
    const { svg, layout } = build({ s })
    const n = barOf(svg, 'n')
    expect(n).toContain('>Dates unavailable<')
    expect(n).not.toContain('<rect')
    expect(n).toContain(`x="${layout.x(layout.range.start) + 8}"`)
    expect(n).toContain('Unavailable')
    const a = barOf(svg, 'a')
    expect(a).toContain('>Finish unavailable<')
    expect(a).toContain(`x1="${layout.x('2026-10-07')}"`)
    expect(a).toContain('stroke-width="2"')
    expect(a).not.toContain('<rect')
    const b = barOf(svg, 'b')
    expect(b).toContain('>Start unavailable<')
    expect(b).toContain(`x1="${layout.x('2026-10-08') + layout.ppd}"`)
  })

  it('wraps tasks with source name, dates and criticality wording in the title', () => {
    const s = schedule([task('x', { name: 'Raw name', critical: 'not-critical' }), task('y', { critical: 'critical' }), task('z')])
    const p = { ...emptyPresentation('s1'), tasks: { x: { displayLabel: 'Friendly' } } }
    const { svg } = build({ s, p })
    expect(barOf(svg, 'x')).toContain('<title>Friendly\nSource name: Raw name\nOct 5 – Oct 9\nNot critical (source)</title>')
    expect(barOf(svg, 'y')).toContain('Critical (source)')
    expect(barOf(svg, 'z')).toContain('Criticality unavailable')
  })

  it('applies selection and emphasis', () => {
    const s = schedule([task('a'), task('b'), task('c', { milestone: true, start: '2026-10-12', finish: '2026-10-12' })])
    const p = { ...emptyPresentation('s1'), tasks: { b: { emphasis: 'muted' as const }, a: { emphasis: 'highlight' as const } } }
    const { svg } = build({ s, p, extra: { selectedId: 'c' } })
    expect(barOf(svg, 'a')).toContain('stroke="var(--ink-2)" stroke-width="1.5"')
    expect(barOf(svg, 'b')).toContain('opacity="0.45"')
    expect(barOf(svg, 'c')).toContain('<polygon')
    expect(barOf(svg, 'c')).toContain('stroke="var(--ink)" stroke-width="3"')
  })

  it('draws summary tasks as a thin bracket', () => {
    const s = schedule([task('sum', { isSummary: true })])
    expect(barOf(build({ s }).svg, 'sum')).toContain('height="6"')
  })

  it('draws the today line only inside the range', () => {
    const { svg, layout } = build({ s: schedule(base) })
    expect(svg).toContain(`<line class="today" x1="${layout.cx('2026-10-14')}"`)
    expect(svg).toContain('<title>Today, Oct 14</title>')
    expect(build({ s: schedule(base), extra: { today: '2030-01-01' } }).svg).not.toContain('class="today"')
  })

  it('draws header ticks per zoom', () => {
    const s = schedule(base)
    const mono = (svg: string) => (svg.match(/font-size="11"/g) ?? []).length
    const day = mono(build({ s, zoom: 'day' }).svg)
    const week = mono(build({ s, zoom: 'week' }).svg)
    const month = mono(build({ s, zoom: 'month' }).svg)
    expect(day).toBeGreaterThan(week)
    expect(week).toBeGreaterThan(month)
    expect(build({ s }).svg).toContain('>Oct 2026<')
  })

  it('adds hidden dependency pills with merged counts', () => {
    const s = schedule([task('a'), task('b', { start: '2026-10-12', finish: '2026-10-14' })])
    const { svg } = build({
      s,
      mutate: (r) => ({
        ...r,
        hiddenDeps: [
          { taskId: 'a', direction: 'successor', count: 1, reason: 'filtered' },
          { taskId: 'a', direction: 'successor', count: 1, reason: 'other-page' },
          { taskId: 'b', direction: 'predecessor', count: 2, reason: 'collapsed' },
        ],
      }),
    })
    expect(svg).toContain('>2 hidden<')
    expect((svg.match(/>2 hidden</g) ?? []).length).toBe(2)
    expect(svg).toContain('1 successor hidden (filtered out)\n1 successor hidden (on another page)')
    expect(svg).toContain('2 predecessors hidden (in a collapsed group)')
  })

  it('places pills outside the bar', () => {
    const s = schedule([task('z', { start: '2026-09-01', finish: '2026-09-02' }), task('a')])
    const { svg, layout } = build({
      s,
      mutate: (r) => ({
        ...r,
        hiddenDeps: [
          { taskId: 'a', direction: 'predecessor' as const, count: 3, reason: 'view' as const },
          { taskId: 'a', direction: 'successor' as const, count: 3, reason: 'view' as const },
        ],
      }),
    })
    const sp = layout.span('2026-10-05', '2026-10-09')
    const pred = svg.match(/data-direction="predecessor"><title>[^<]*<\/title><rect x="([\d.]+)" y="[\d.]+" width="([\d.]+)"/)!
    const succ = svg.match(/data-direction="successor"><title>[^<]*<\/title><rect x="([\d.]+)"/)!
    expect(Number(pred[1]) + Number(pred[2])).toBeLessThan(sp.x)
    expect(Number(succ[1])).toBeGreaterThan(sp.x + sp.width)
    expect(svg).toContain('3 predecessors hidden (not in this view)')
  })

  it('highlights downstream tasks and dims the rest in the dependencies view', () => {
    const s = schedule([
      task('a', { start: '2026-10-05', finish: '2026-10-06' }),
      task('b', { start: '2026-10-07', finish: '2026-10-08', deps: [dep('a')] }),
      task('c', { start: '2026-10-09', finish: '2026-10-12', deps: [dep('b')] }),
      task('d', { start: '2026-10-09', finish: '2026-10-12', deps: [dep('a')] }),
      task('e', { start: '2026-10-13', finish: '2026-10-14' }),
    ])
    const { svg } = build({ s, view: 'dependencies', extra: { selectedId: 'b' } })
    expect(barOf(svg, 'c')).toContain('fill="var(--highlight)"')
    expect(barOf(svg, 'd')).toContain('opacity="0.35"')
    expect(barOf(svg, 'e')).toContain('opacity="0.35"')
    expect(barOf(svg, 'b')).not.toContain('opacity')
    expect(pathFor(svg, 'c', 'b')).toContain('stroke="var(--highlight)"')
    expect(pathFor(svg, 'b', 'a')).toContain('stroke="var(--dim)"')
    expect(pathFor(svg, 'd', 'a')).toContain('marker-end="url(#tl-dependencies-ah-dim)"')
  })

  it('draws baseline ghosts under the bars', () => {
    const s = schedule([task('a'), task('m', { milestone: true, start: '2026-10-12', finish: '2026-10-12' }), task('n')])
    const baseline: Baseline = {
      id: 'b1',
      scheduleId: 's1',
      name: 'B',
      createdAt: '2026-10-01T00:00:00Z',
      tasks: [
        { id: 'a', name: 'a', start: '2026-10-01', finish: '2026-10-03', milestone: false },
        { id: 'm', name: 'm', start: '2026-10-10', finish: '2026-10-10', milestone: true },
      ],
    }
    const { svg, layout } = build({ s, extra: { baseline }, baselineDates: ['2026-10-01', '2026-10-10'] })
    const sp = layout.span('2026-10-01', '2026-10-03')
    expect(svg).toContain(`<rect class="ghost" x="${sp.x}"`)
    expect(svg).toContain('height="4"')
    expect((svg.match(/class="ghost"/g) ?? []).length).toBe(2)
    expect(svg).toMatch(/<polygon class="ghost"[^>]*fill="none"/)
    expect(svg.indexOf('class="ghost"')).toBeLessThan(svg.indexOf('class="bar"'))
    expect(build({ s }).svg).not.toContain('class="ghost"')
  })

  it('renders the executive summary and milestones row', () => {
    const s = schedule([
      task('a', { critical: 'critical' }),
      task('b', { workstream: 'b', critical: 'not-critical', start: '2026-10-12', finish: '2026-10-20' }),
      task('m', { name: 'Gate', milestone: true, start: '2026-10-08', finish: '2026-10-08' }),
    ])
    const { svg, layout } = build({ s, view: 'executive' })
    expect(svg).toContain('<g class="summary" data-group-key="a">')
    expect(svg).toContain('<g class="summary" data-group-key="b">')
    expect(svg).toContain('>Oct 20<')
    expect(svg).toContain('>Gate · Oct 8<')
    expect(svg).toContain(`<polygon points="${layout.cx('2026-10-08')},`)
    expect(svg.match(/<g class="summary" data-group-key="a">[\s\S]*?<\/g>/)![0]).toContain('fill="var(--critical)"')
    expect(svg.match(/<g class="summary" data-group-key="b">[\s\S]*?<\/g>/)![0]).toContain('fill="var(--task)"')
  })

  it('escapes names', () => {
    const s = schedule([task('x', { name: 'A <b> & "C"' })])
    const { svg } = build({ s })
    expect(svg).toContain('A &lt;b&gt; &amp; &quot;C&quot;')
    expect(svg).not.toContain('<b>')
    const withLabels = build({ s, extra: { labels: { width: 260, title: 'S <&> "t"' } } }).svg
    expect(withLabels).toContain('S &lt;&amp;&gt; &quot;t&quot;')
  })

  it('draws an export label column and shifts the chart', () => {
    const long = 'A very long task name that must be truncated for the column'
    const s = schedule([task('x', { name: long })])
    const plain = build({ s })
    const { svg, layout } = build({ s, extra: { labels: { width: 260, title: 'S · Detailed' } } })
    expect(svg).toContain('>S · Detailed<')
    expect(svg).toContain(`width="${layout.width + 260}"`)
    expect(svg).toContain('<g transform="translate(260,0)">')
    expect(svg).toContain(`${long.slice(0, 31)}…<title>${long}</title>`)
    expect(svg).toContain('font-size="12.5"')
    expect(svg).toContain('font-size="14" font-weight="600"')
    expect(plain.svg).not.toContain('class="labels"')
    expect(barOf(svg, 'x')).toBe(barOf(plain.svg, 'x'))
  })

  it('keeps marker ids unique per view', () => {
    const s = schedule(base)
    const a = build({ s, view: 'detailed' }).svg
    const b = build({ s, view: 'dependencies' }).svg
    expect(a).toContain('id="tl-detailed-ah-norm"')
    expect(b).toContain('id="tl-dependencies-ah-norm"')
  })

  describe('theme', () => {
    const css = readFileSync(new URL('../src/web/theme.css', import.meta.url), 'utf8')
    const tokens = parseThemeTokens(css)

    it('parses :root tokens', () => {
      expect(tokens.ink).toBe('#1a1d21')
      expect(tokens['lane-6']).toBe('#8a6d9e')
      expect(parseThemeTokens('a{--x:1;} :root { --y: 2; /* --z: 3; */ } :root { --w: 4; }')).toEqual({ y: '2' })
    })

    it('screen output uses only var() colors', () => {
      const s = schedule(base)
      const svg = build({ s, extra: { baseline: null } }).svg
      expect(svg).not.toMatch(/#[0-9a-fA-F]{3,8}\b/)
      expect(svg).not.toMatch(/rgba?\(/)
      expect(svg).toContain('var(--critical)')
    })

    it('resolves every token for export', () => {
      const s = schedule([...base, task('u', { deps: [dep('p', 'unknown')] }), task('m', { milestone: true, start: '2026-10-12', finish: '2026-10-12' })])
      const color = tokenResolver(tokens)
      const svg = build({
        s,
        view: 'dependencies',
        extra: { color, font: { text: 'Segoe UI, sans-serif', mono: 'Consolas, monospace' }, selectedId: 'p', labels: { width: 240, title: 'T' } },
      }).svg
      expect(svg).not.toContain('var(')
      expect(svg).toContain('#c2410c')
      expect(() => tokenResolver({})('ink')).toThrow()
    })
  })
})
