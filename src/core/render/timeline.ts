// The one SVG timeline renderer. The screen and every export (SVG, PNG, PDF) call this
// function, so they cannot drift. Positions come only from Layout.

import { dayNum, fmtShort, isoOf, weekday } from '../dates'
import type { Baseline, Criticality, DependencyType, Presentation, Schedule, Task } from '../model'
import type { Layout } from '../layout'
import { downstreamOf } from '../reachability'
import type { HiddenDependency, ResolvedView, Row, ViewId } from '../views'

export const COLOR_TOKENS = [
  'canvas', 'surface', 'surface-2', 'band', 'select', 'ink', 'ink-2', 'muted', 'border', 'border-2',
  'critical', 'task', 'highlight', 'dim', 'baseline', 'dep-line', 'dep-unknown', 'summary-bar',
  'milestone', 'today', 'unavailable-ink', 'lane-1', 'lane-2', 'lane-3', 'lane-4', 'lane-5', 'lane-6',
] as const

export type ColorToken = (typeof COLOR_TOKENS)[number]

export interface TimelineOptions {
  schedule: Schedule
  presentation: Presentation
  resolved: ResolvedView
  layout: Layout
  selectedId?: string | null
  showDeps: boolean
  baseline?: Baseline | null
  /** ISO date. The renderer never reads the clock. */
  today: string
  color: (token: ColorToken) => string
  font: { text: string; mono: string }
  /** Export only: draws a left label column. */
  labels?: { width: number; title: string }
}

const VIEW_LABELS: Record<ViewId, string> = {
  executive: 'Executive',
  detailed: 'Detailed',
  dependencies: 'Dependencies',
  workstreams: 'Workstreams',
  critical: 'Critical',
}

const MONTHS = ['Jan', 'Feb', 'Mar', 'Apr', 'May', 'Jun', 'Jul', 'Aug', 'Sep', 'Oct', 'Nov', 'Dec']

const DEP_NAMES: Record<Exclude<DependencyType, 'unknown'>, string> = {
  FS: 'Finish-to-start',
  SS: 'Start-to-start',
  FF: 'Finish-to-finish',
  SF: 'Start-to-finish',
}

const REASON_TEXT: Record<HiddenDependency['reason'], string> = {
  filtered: 'filtered out',
  collapsed: 'in a collapsed group',
  'other-page': 'on another page',
  view: 'not in this view',
}

const MILESTONE_HALF = 8
const ARROW_GAP = 8
const LABEL_MAX = 32

function esc(s: string): string {
  return s.replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;')
}

function n(v: number): string {
  return String(Number(v.toFixed(2)))
}

function textWidth(s: string, size: number): number {
  return s.length * size * 0.55
}

function critWording(c: Criticality): string {
  if (c === 'critical') return 'Critical (source)'
  if (c === 'not-critical') return 'Not critical (source)'
  return 'Criticality unavailable'
}

function lagText(lag: number | null): string {
  if (lag === null) return 'lag not stated'
  return `lag ${lag < 0 ? '-' : '+'}${Math.abs(lag)}d`
}

function plural(count: number, word: string): string {
  return `${count} ${word}${count === 1 ? '' : 's'}`
}

interface Geometry {
  /** x of the start edge, when the start is known. */
  left: number | null
  /** x of the finish edge, when the finish is known. */
  right: number | null
  /** Outermost x on each side, for placing markers. */
  minX: number
  maxX: number
}

export function renderTimeline(o: TimelineOptions): string {
  const { schedule, presentation, resolved, layout, color, font } = o
  const view = resolved.view
  const lx = o.labels ? o.labels.width : 0
  const head = layout.y(0)
  const rowH = layout.y(1) - layout.y(0)
  const totalW = layout.width + lx
  const H = layout.height
  const W = layout.width
  const pid = `tl-${view}${o.labels ? '-x' : ''}`
  const r0 = layout.range.start
  const r1 = layout.range.end
  const x0 = layout.x(r0)

  const taskById = new Map(schedule.tasks.map((t) => [t.id, t]))
  const labelOf = (t: Task) => presentation.tasks[t.id]?.displayLabel?.trim() || t.name
  const groupIndex = new Map(resolved.groups.map((g, i) => [g.key, i]))
  const rowIndexOf = new Map<string, number>()
  resolved.rows.forEach((r, i) => {
    if (r.kind === 'task') rowIndexOf.set(r.taskId, i)
  })

  const hl =
    view === 'dependencies' && o.selectedId && taskById.has(o.selectedId)
      ? new Set<string>([o.selectedId, ...downstreamOf(schedule, o.selectedId)])
      : null

  const txt = (x: number, y: number, size: number, fill: ColorToken, s: string, extra = '') =>
    `<text x="${n(x)}" y="${n(y)}" font-size="${size}" fill="${color(fill)}"${extra}>${esc(s)}</text>`
  const op = (v: number) => (v === 1 ? '' : ` opacity="${v}"`)

  const bands: string[] = []
  const header: string[] = []
  const todayLine: string[] = []
  const ghosts: string[] = []
  const deps: string[] = []
  const bars: string[] = []
  const pills: string[] = []
  const geo = new Map<string, Geometry>()
  const rowOfGeo = new Map<string, number>()

  // Row bands and separators.
  resolved.rows.forEach((r, i) => {
    const y = layout.y(i)
    if (r.kind === 'group' || r.kind === 'milestones') bands.push(`<rect x="0" y="${n(y)}" width="${n(W)}" height="${n(rowH)}" fill="${color('band')}"/>`)
    if (r.kind === 'task' && r.taskId === o.selectedId) bands.push(`<rect x="0" y="${n(y)}" width="${n(W)}" height="${n(rowH)}" fill="${color('select')}"/>`)
    bands.push(`<line x1="0" x2="${n(W)}" y1="${n(y + rowH - 0.5)}" y2="${n(y + rowH - 0.5)}" stroke="${color('border-2')}"/>`)
  })

  // Header, month labels and gridlines.
  header.push(`<rect x="0" y="0" width="${n(W)}" height="${head}" fill="${color('surface-2')}"/>`)
  let nextFirst = Infinity
  for (let d = r0 + 1; d <= r1; d++) {
    if (isoOf(d).slice(8, 10) === '01') {
      nextFirst = d
      break
    }
  }
  const startLabelRoom = Math.ceil(70 / layout.ppd)
  for (let d = r0; d <= r1; d++) {
    const iso = isoOf(d)
    const dom = Number(iso.slice(8, 10))
    const monday = weekday(d) === 1
    const x = layout.x(d)
    if (dom === 1 || (d === r0 && nextFirst - r0 >= startLabelRoom)) {
      header.push(txt(x + 6, 17, 12, 'ink-2', `${MONTHS[Number(iso.slice(5, 7)) - 1]} ${iso.slice(0, 4)}`, ' font-weight="600"'))
    }
    const mark = layout.zoom === 'day' || (layout.zoom === 'week' && monday) || (layout.zoom === 'month' && dom === 1)
    if (!mark) continue
    const strong = layout.zoom === 'day' ? monday : true
    header.push(`<line x1="${n(x + 0.5)}" x2="${n(x + 0.5)}" y1="${head - 14}" y2="${n(H)}" stroke="${color(strong ? 'border' : 'border-2')}"/>`)
    header.push(txt(x + 4, head - 9, 11, 'muted', String(dom), ` font-family="${esc(font.mono)}"`))
  }
  header.push(`<line x1="0" x2="${n(W)}" y1="${head - 0.5}" y2="${head - 0.5}" stroke="${color('border')}"/>`)

  // Today.
  const td = dayNum(o.today)
  if (td >= r0 && td <= r1) {
    const x = layout.cx(o.today)
    todayLine.push(
      `<line class="today" x1="${n(x)}" x2="${n(x)}" y1="${head}" y2="${n(H)}" stroke="${color('today')}" stroke-width="1.5" stroke-dasharray="4 3"><title>Today, ${esc(fmtShort(o.today))}</title></line>`,
    )
  }

  // Per-row drawing.
  const unavailable = (x: number, y: number, s: string) => txt(x, y + 22, 12, 'muted', s, ' font-style="italic"')

  resolved.rows.forEach((r: Row, i) => {
    const y = layout.y(i)
    if (r.kind === 'task') {
      const t = taskById.get(r.taskId)
      if (t) drawTask(t, r.groupKey, i, y)
    } else if (r.kind === 'group') {
      if (r.start && r.finish) {
        const s = layout.span(r.start, r.finish)
        bars.push(`<rect x="${n(s.x)}" y="${n(y + 15)}" width="${n(s.width)}" height="6" rx="1.5" fill="${color('summary-bar')}"/>`)
      }
    } else if (r.kind === 'summary') {
      drawSummary(r, y)
    } else {
      drawMilestonesRow(r, i, y)
    }
  })

  function drawTask(t: Task, groupKey: string, i: number, y: number) {
    const emphasis = presentation.tasks[t.id]?.emphasis ?? 'none'
    const selected = t.id === o.selectedId
    const dimmed = hl !== null && !hl.has(t.id)
    const opacity = dimmed ? 0.35 : emphasis === 'muted' ? 0.45 : 1
    const stroke = selected ? { c: 'ink' as const, w: 2 } : emphasis === 'highlight' ? { c: 'ink-2' as const, w: 1.5 } : null
    const strokeAttr = () => (stroke ? ` stroke="${color(stroke.c)}" stroke-width="${selected && t.milestone ? 3 : stroke.w}"` : '')
    const lane = `lane-${((groupIndex.get(groupKey) ?? 0) % 6) + 1}` as ColorToken
    const base: ColorToken = view === 'workstreams' ? lane : t.critical === 'critical' ? 'critical' : t.milestone ? 'milestone' : 'task'
    const fill = color(hl !== null && !dimmed && !selected ? 'highlight' : base)

    const label = labelOf(t)
    const hasOverride = label !== t.name
    const known = t.start && t.finish ? `${fmtShort(t.start)} – ${fmtShort(t.finish)}` : null
    let dates: string
    if (t.milestone && (t.start ?? t.finish)) dates = fmtShort((t.start ?? t.finish)!)
    else if (known) dates = known
    else if (t.start) dates = `Start ${fmtShort(t.start)}, finish unavailable`
    else if (t.finish) dates = `Start unavailable, finish ${fmtShort(t.finish)}`
    else dates = 'Unavailable'
    const title = [label, ...(hasOverride ? [`Source name: ${t.name}`] : []), dates, critWording(t.critical)].join('\n')
    const open = `<g class="bar" data-task-id="${esc(t.id)}"${op(opacity)}><title>${esc(title)}</title>`
    const cy = y + 18
    const parts: string[] = [open]
    rowOfGeo.set(t.id, i)

    const milestoneDate = t.milestone ? (t.start ?? t.finish) : null
    if (milestoneDate) {
      const cx = layout.cx(milestoneDate)
      const h = MILESTONE_HALF
      parts.push(
        `<polygon points="${n(cx)},${cy - h} ${n(cx + h)},${cy} ${n(cx)},${cy + h} ${n(cx - h)},${cy}" fill="${fill}"${strokeAttr()}/>`,
      )
      geo.set(t.id, { left: cx - h, right: cx + h, minX: cx - h, maxX: cx + h })
    } else if (t.start && t.finish) {
      const s = layout.span(t.start, t.finish)
      if (t.isSummary) {
        parts.push(`<rect x="${n(s.x)}" y="${n(y + 15)}" width="${n(s.width)}" height="6" rx="1.5" fill="${color('summary-bar')}"${strokeAttr()}/>`)
      } else {
        parts.push(`<rect x="${n(s.x)}" y="${n(y + 9)}" width="${n(s.width)}" height="18" rx="3" fill="${fill}"${strokeAttr()}/>`)
      }
      geo.set(t.id, { left: s.x, right: s.x + s.width, minX: s.x, maxX: s.x + s.width })
    } else if (t.start || t.finish) {
      const startKnown = t.start !== null
      const tx = startKnown ? layout.x(t.start!) : layout.x(t.finish!) + layout.ppd
      const note = startKnown ? 'Finish unavailable' : 'Start unavailable'
      parts.push(`<line x1="${n(tx)}" x2="${n(tx)}" y1="${n(y + 9)}" y2="${n(y + 27)}" stroke="${color('muted')}" stroke-width="2"/>`)
      parts.push(unavailable(tx + 8, y, note))
      geo.set(t.id, {
        left: startKnown ? tx : null,
        right: startKnown ? null : tx,
        minX: tx,
        maxX: tx + 8 + textWidth(note, 12),
      })
    } else {
      const tx = x0 + 8
      const note = 'Dates unavailable'
      parts.push(unavailable(tx, y, note))
      geo.set(t.id, { left: null, right: null, minX: tx, maxX: tx + textWidth(note, 12) })
    }
    parts.push('</g>')
    bars.push(parts.join(''))
  }

  function drawSummary(r: Extract<Row, { kind: 'summary' }>, y: number) {
    const parts = [`<g class="summary" data-group-key="${esc(r.key)}">`]
    const dates = r.start && r.finish ? `${fmtShort(r.start)} – ${fmtShort(r.finish)}` : 'Unavailable'
    const crit =
      r.critical === 'critical' ? 'Contains critical tasks (source)' : r.critical === 'not-critical' ? 'No critical tasks (source)' : 'Criticality unavailable'
    parts.push(`<title>${esc([r.label, dates, crit].join('\n'))}</title>`)
    if (r.start && r.finish) {
      const s = layout.span(r.start, r.finish)
      parts.push(`<rect x="${n(s.x)}" y="${n(y + 11)}" width="${n(s.width)}" height="14" rx="3" fill="${color(r.critical === 'critical' ? 'critical' : 'task')}"/>`)
      parts.push(txt(s.x + s.width + 8, y + 22, 12, 'muted', fmtShort(r.finish)))
    } else if (!r.start && !r.finish) {
      parts.push(unavailable(x0 + 8, y, 'Dates unavailable'))
    }
    parts.push('</g>')
    bars.push(parts.join(''))
  }

  function drawMilestonesRow(r: Extract<Row, { kind: 'milestones' }>, i: number, y: number) {
    for (const id of r.taskIds) {
      const t = taskById.get(id)
      const date = t ? (t.start ?? t.finish) : null
      if (!t || !date) continue
      const cx = layout.cx(date)
      const cy = y + 18
      const h = 7
      const label = `${labelOf(t)} · ${fmtShort(date)}`
      const title = [labelOf(t), ...(labelOf(t) !== t.name ? [`Source name: ${t.name}`] : []), fmtShort(date), critWording(t.critical)].join('\n')
      bars.push(
        `<g class="bar" data-task-id="${esc(id)}"><title>${esc(title)}</title>` +
          `<polygon points="${n(cx)},${cy - h} ${n(cx + h)},${cy} ${n(cx)},${cy + h} ${n(cx - h)},${cy}" fill="${color('milestone')}"/>` +
          txt(cx + 12, cy + 4, 12, 'ink', label) +
          '</g>',
      )
      geo.set(id, { left: cx - h, right: cx + h, minX: cx - h, maxX: cx + 12 + textWidth(label, 12) })
      rowOfGeo.set(id, i)
    }
  }

  // Baseline ghosts, drawn under the bars.
  if (o.baseline) {
    const bmap = new Map(o.baseline.tasks.map((b) => [b.id, b]))
    resolved.rows.forEach((r, i) => {
      if (r.kind !== 'task') return
      const b = bmap.get(r.taskId)
      if (!b) return
      const y = layout.y(i)
      const date = b.milestone ? (b.start ?? b.finish) : null
      if (date) {
        const cx = layout.cx(date)
        ghosts.push(
          `<polygon class="ghost" points="${n(cx)},${y + 25} ${n(cx + 4)},${y + 29} ${n(cx)},${y + 33} ${n(cx - 4)},${y + 29}" fill="none" stroke="${color('baseline')}" stroke-width="1.5"/>`,
        )
      } else if (!b.milestone && b.start && b.finish) {
        const s = layout.span(b.start, b.finish)
        ghosts.push(`<rect class="ghost" x="${n(s.x)}" y="${y + 29}" width="${n(s.width)}" height="4" rx="1" fill="${color('baseline')}"/>`)
      }
    })
  }

  // Dependencies.
  const showDeps = view !== 'executive' && (o.showDeps || view === 'dependencies')
  const markerIds = { crit: `${pid}-ah-crit`, norm: `${pid}-ah-norm`, hl: `${pid}-ah-hl`, dim: `${pid}-ah-dim`, ocUnk: `${pid}-oc-unk`, ocHl: `${pid}-oc-hl`, ocDim: `${pid}-oc-dim` }
  if (showDeps) {
    for (const [i, r] of resolved.rows.entries()) {
      if (r.kind !== 'task') continue
      const succ = taskById.get(r.taskId)
      const sg = geo.get(r.taskId)
      if (!succ || !sg) continue
      for (const d of succ.deps) {
        const pi = rowIndexOf.get(d.predecessorId)
        const pred = taskById.get(d.predecessorId)
        const pg = geo.get(d.predecessorId)
        if (pi === undefined || !pred || !pg) continue
        const unknown = d.type === 'unknown'
        const fromFinish = d.type === 'FS' || d.type === 'FF' || unknown
        const toFinish = d.type === 'FF' || d.type === 'SF'
        const sx = fromFinish ? pg.right : pg.left
        const tx = toFinish ? sg.right : sg.left
        if (sx === null || tx === null) continue
        const sy = layout.y(pi) + 18
        const ty = layout.y(i) + 18
        const exitX = sx + (fromFinish ? ARROW_GAP : -ARROW_GAP)
        const enterX = tx + (toFinish ? ARROW_GAP : -ARROW_GAP)
        const gy = i > pi ? layout.y(i) : layout.y(i + 1)
        const path = `M${n(sx)} ${n(sy)} H${n(exitX)} V${n(gy)} H${n(enterX)} V${n(ty)} H${n(tx)}`

        const both = hl !== null && hl.has(d.predecessorId) && hl.has(succ.id)
        const crit = pred.critical === 'critical' && succ.critical === 'critical'
        let stroke: ColorToken = unknown ? 'dep-unknown' : crit ? 'critical' : 'dep-line'
        let width = crit ? 1.5 : 1.25
        let marker = unknown ? markerIds.ocUnk : crit ? markerIds.crit : markerIds.norm
        if (hl) {
          if (both) {
            stroke = 'highlight'
            width = 2
            marker = unknown ? markerIds.ocHl : markerIds.hl
          } else {
            stroke = 'dim'
            marker = unknown ? markerIds.ocDim : markerIds.dim
          }
        }
        const pn = labelOf(pred)
        const sn = labelOf(succ)
        const detail = unknown ? 'type unknown (not stated in source)' : `${DEP_NAMES[d.type as Exclude<DependencyType, 'unknown'>]} · ${lagText(d.lagDays)}`
        deps.push(
          `<path class="dep" data-pred-id="${esc(pred.id)}" data-succ-id="${esc(succ.id)}" data-dep-type="${d.type}" d="${path}" fill="none" stroke="${color(stroke)}" stroke-width="${width}"` +
            `${unknown ? ' stroke-dasharray="2 3"' : ''} marker-end="url(#${marker})"><title>${esc(`${pn} → ${sn} · ${detail}`)}</title></path>`,
        )
      }
    }
  }

  // Hidden dependency markers.
  const hiddenByKey = new Map<string, HiddenDependency[]>()
  for (const h of resolved.hiddenDeps) {
    if (!geo.has(h.taskId)) continue
    const key = `${h.taskId}\u0000${h.direction}`
    const list = hiddenByKey.get(key)
    if (list) list.push(h)
    else hiddenByKey.set(key, [h])
  }
  for (const list of hiddenByKey.values()) {
    const first = list[0]
    const g = geo.get(first.taskId)!
    const row = rowOfGeo.get(first.taskId)!
    const total = list.reduce((a, h) => a + h.count, 0)
    const word = first.direction === 'predecessor' ? 'predecessor' : 'successor'
    const lines = list.map((h) => `${plural(h.count, word)} hidden (${REASON_TEXT[h.reason]})`)
    const label = `${total} hidden`
    const w = textWidth(label, 12) + 14
    const y = layout.y(row) + (rowH - 18) / 2
    const x = first.direction === 'predecessor' ? g.minX - 6 - w : g.maxX + 6
    pills.push(
      `<g class="hidden-deps" data-task-id="${esc(first.taskId)}" data-direction="${first.direction}"><title>${esc(lines.join('\n'))}</title>` +
        `<rect x="${n(x)}" y="${n(y)}" width="${n(w)}" height="18" rx="9" fill="${color('surface-2')}" stroke="${color('border')}"/>` +
        txt(x + w / 2, y + 13, 12, 'ink-2', label, ' text-anchor="middle"') +
        '</g>',
    )
  }

  const marker = (id: string, tok: ColorToken) =>
    `<marker id="${id}" viewBox="0 0 6 6" refX="5.5" refY="3" markerWidth="6" markerHeight="6" orient="auto"><path d="M0 0L6 3L0 6z" fill="${color(tok)}"/></marker>`
  const circle = (id: string, tok: ColorToken) =>
    `<marker id="${id}" viewBox="0 0 6 6" refX="5.5" refY="3" markerWidth="6" markerHeight="6" orient="auto"><circle cx="3" cy="3" r="2.2" fill="${color('surface')}" stroke="${color(tok)}" stroke-width="1"/></marker>`
  const defs = showDeps
    ? `<defs>${marker(markerIds.crit, 'critical')}${marker(markerIds.norm, 'dep-line')}${marker(markerIds.hl, 'highlight')}${marker(markerIds.dim, 'dim')}` +
      `${circle(markerIds.ocUnk, 'dep-unknown')}${circle(markerIds.ocHl, 'highlight')}${circle(markerIds.ocDim, 'dim')}</defs>`
    : ''

  const out: string[] = []
  out.push(
    `<svg xmlns="http://www.w3.org/2000/svg" class="tl-svg" width="${n(totalW)}" height="${n(H)}" viewBox="0 0 ${n(totalW)} ${n(H)}" font-family="${esc(font.text)}" role="img" aria-label="Timeline, ${esc(VIEW_LABELS[view])} view">`,
  )
  out.push(defs)
  out.push(`<rect width="${n(totalW)}" height="${n(H)}" fill="${color('surface')}"/>`)
  out.push(`<g transform="translate(${n(lx)},0)">`)
  out.push(...bands, ...header, ...todayLine, ...ghosts, ...deps, ...bars, ...pills)
  out.push('</g>')

  if (o.labels) {
    const w = lx
    out.push(`<g class="labels"><rect x="0" y="0" width="${n(w)}" height="${n(H)}" fill="${color('surface')}"/><rect x="0" y="0" width="${n(w)}" height="${head}" fill="${color('surface-2')}"/>`)
    out.push(txt(12, head - 14, 12, 'muted', o.labels.title))
    resolved.rows.forEach((r, i) => {
      const y = layout.y(i)
      let full: string
      let indent = 12
      let bold = true
      if (r.kind === 'task') {
        const t = taskById.get(r.taskId)
        full = t ? labelOf(t) : r.taskId
        indent = 24 + r.depth * 12
        bold = false
      } else full = r.label
      const source = r.kind === 'task' ? (taskById.get(r.taskId)?.name ?? full) : full
      const short = full.length > LABEL_MAX ? `${full.slice(0, LABEL_MAX - 1)}…` : full
      if (r.kind !== 'task') out.push(`<rect x="0" y="${n(y)}" width="${n(w)}" height="${n(rowH)}" fill="${color('band')}"/>`)
      out.push(
        `<text x="${indent}" y="${n(y + 22)}" font-size="${bold ? 14 : 12.5}"${bold ? ' font-weight="600"' : ''} fill="${color('ink')}">${esc(short)}<title>${esc(source)}</title></text>`,
      )
      out.push(`<line x1="0" x2="${n(w)}" y1="${n(y + rowH - 0.5)}" y2="${n(y + rowH - 0.5)}" stroke="${color('border-2')}"/>`)
    })
    out.push(`<line x1="${n(w - 0.5)}" x2="${n(w - 0.5)}" y1="0" y2="${n(H)}" stroke="${color('border')}"/></g>`)
  }
  out.push('</svg>')
  return out.join('')
}
