import { useCallback, useEffect, useRef, useState } from 'react'
import type { Baseline, Presentation, Schedule, ScheduleMeta } from '../core/model'
import { VIEW_IDS, type ViewId } from '../core/views'
import { api, type ExportFormat } from './api'
import { Command, Palette } from './components/Palette'
import { ImportDialog } from './components/ImportDialog'
import { Landing } from './components/Landing'
import { ApiPage, BaselinesPage, EXPORT_FORMATS, ExportsPage, ViewsPage } from './components/Pages'
import { BrandMark, Loading, Modal, SearchIcon, Toasts, useToasts } from './components/ui'
import { INITIAL_WORKSPACE, Workspace, type WorkspaceState } from './components/Workspace'
import { download, localTodayISO, longDateTime, shortDate } from './format'
import { PAGE_IDS, readRoute, routeHash, type PageId, type Route } from './route'
import { S } from '../core/strings'

interface OpenSchedule {
  schedule: Schedule
  presentation: Presentation
  baselines: Baseline[]
}

const PAGE_LABEL: Record<PageId, string> = { schedules: S.nav.schedules, views: S.nav.views, baselines: S.nav.baselines, exports: S.nav.exports, api: S.nav.api }

export function App() {
  const [route, setRouteState] = useState<Route>(() => readRoute(location.hash))
  const [schedules, setSchedules] = useState<ScheduleMeta[] | null>(null)
  const [listLoading, setListLoading] = useState(true)
  const [open, setOpen] = useState<OpenSchedule | null>(null)
  const [openError, setOpenError] = useState<string | null>(null)
  const [ws, setWs] = useState<WorkspaceState>(INITIAL_WORKSPACE)
  const [apiDown, setApiDown] = useState(false)
  const [importing, setImporting] = useState<{ replace: { id: string; name: string } | null } | null>(null)
  const [paletteOpen, setPaletteOpen] = useState(false)
  const [baselineOpen, setBaselineOpen] = useState(false)
  const [baselineName, setBaselineName] = useState('')
  const [exportMenu, setExportMenu] = useState(false)
  const [todayTick, setTodayTick] = useState(0)
  const exportBtn = useRef<HTMLButtonElement>(null)
  const { toasts, push: toast } = useToasts()

  const setRoute = useCallback((next: Route) => {
    setRouteState(next)
    const hash = routeHash(next)
    try {
      if (location.hash !== hash) history.pushState(null, '', hash)
    } catch {
      // History can be unavailable in sandboxed frames; the in-memory route still works.
    }
  }, [])

  useEffect(() => {
    const onHash = () => setRouteState(readRoute(location.hash))
    window.addEventListener('popstate', onHash)
    window.addEventListener('hashchange', onHash)
    return () => {
      window.removeEventListener('popstate', onHash)
      window.removeEventListener('hashchange', onHash)
    }
  }, [])

  const loadList = useCallback(async () => {
    setListLoading(true)
    try {
      setSchedules(await api.listSchedules())
      setApiDown(false)
    } catch (e) {
      setApiDown(true)
      toast((e as Error).message, true)
    } finally {
      setListLoading(false)
    }
  }, [toast])

  useEffect(() => {
    void loadList()
  }, [loadList])

  const loadSchedule = useCallback(
    async (id: string) => {
      setOpenError(null)
      try {
        const [schedule, presentation, baselines] = await Promise.all([api.getSchedule(id), api.getPresentation(id), api.listBaselines(id)])
        setOpen({ schedule, presentation, baselines })
      } catch (e) {
        setOpen(null)
        setOpenError((e as Error).message)
      }
    },
    [],
  )

  // Open the schedule named in the route; reset workspace state when switching schedules.
  useEffect(() => {
    if (!route.schedule) {
      setOpen(null)
      return
    }
    if (open?.schedule.id === route.schedule) return
    setOpen(null)
    setWs(INITIAL_WORKSPACE)
    void loadSchedule(route.schedule)
  }, [route.schedule]) // eslint-disable-line react-hooks/exhaustive-deps

  const go = (page: PageId) => setRoute({ ...route, page, group: undefined })
  const openView = (view: ViewId, group?: string) => {
    if (view === 'workstreams' && group !== undefined) setWs((s) => ({ ...s, collapsed: [], page: 1 }))
    else setWs((s) => ({ ...s, page: 1 }))
    setRoute({ ...route, page: 'schedules', view, group })
  }
  const openSchedule = (id: string) => setRoute({ page: 'schedules', schedule: id, view: route.view })
  const toList = () => setRoute({ page: 'schedules', schedule: null, view: route.view })

  const saveBaseline = async () => {
    if (!open || !baselineName.trim()) return
    try {
      const b = await api.saveBaseline(open.schedule.id, baselineName.trim())
      setOpen({ ...open, baselines: [b, ...open.baselines] })
      setBaselineOpen(false)
      toast(S.baselines.saved(b.name))
    } catch (e) {
      toast((e as Error).message, true)
    }
  }
  const openBaselineDialog = () => {
    if (!open) return
    setBaselineName(S.baselines.defaultName(open.baselines.length + 1, shortDate(localTodayISO())))
    setBaselineOpen(true)
  }
  const deleteBaseline = async (b: Baseline) => {
    if (!open || !window.confirm(S.baselines.deleteConfirm(b.name))) return
    try {
      await api.deleteBaseline(open.schedule.id, b.id)
      setOpen({ ...open, baselines: open.baselines.filter((x) => x.id !== b.id) })
      if (ws.compareId === b.id) setWs((s) => ({ ...s, compareId: '' }))
      toast(S.baselines.deleted(b.name))
    } catch (e) {
      toast((e as Error).message, true)
    }
  }

  const exportAs = async (format: ExportFormat) => {
    setExportMenu(false)
    if (!open) return
    try {
      const { blob, filename } = await api.exportFile(open.schedule.id, {
        format,
        view: route.view,
        zoom: ws.zoom,
        filters: ws.filters,
        collapsed: ws.collapsed,
        focusGroup: route.view === 'workstreams' ? route.group : undefined,
        page: ws.page,
        compare: ws.compareId || undefined,
        showDeps: ws.showDeps,
        selectedId: ws.selectedId ?? undefined,
        today: localTodayISO(),
      })
      download(filename, blob)
      toast(S.exportsPage.done(S.views[route.view].label, format.toUpperCase()))
    } catch (e) {
      toast(S.exportsPage.failed((e as Error).message), true)
    }
  }

  const afterImport = (s: Schedule, replaced: boolean) => {
    setImporting(null)
    toast(replaced ? S.importer.reimported(s.name) : S.importer.accepted(s.name))
    void loadList()
    if (replaced) void loadSchedule(s.id)
    setRoute({ page: 'schedules', schedule: s.id, view: route.view })
  }

  // Global keys: command menu, Escape closes the drawer or menus.
  useEffect(() => {
    const onKey = (e: KeyboardEvent) => {
      if ((e.ctrlKey || e.metaKey) && e.key.toLowerCase() === 'k') {
        e.preventDefault()
        setPaletteOpen((o) => !o)
      } else if (e.key === 'Escape' && !document.querySelector('dialog[open]')) {
        if (exportMenu) {
          setExportMenu(false)
          exportBtn.current?.focus()
        } else if (ws.selectedId) setWs((s) => ({ ...s, selectedId: null }))
      }
    }
    window.addEventListener('keydown', onKey)
    return () => window.removeEventListener('keydown', onKey)
  }, [exportMenu, ws.selectedId])

  useEffect(() => {
    if (!exportMenu) return
    const onDoc = (e: MouseEvent) => {
      if (!(e.target as Element).closest('.menu-wrap')) setExportMenu(false)
    }
    document.addEventListener('click', onDoc)
    return () => document.removeEventListener('click', onDoc)
  }, [exportMenu])

  const commands = (): Command[] => {
    const list: Command[] = []
    const h = S.palette.hint
    if (open) {
      VIEW_IDS.forEach((v) => list.push({ label: S.palette.showView(S.views[v].label), hint: h.view, run: () => openView(v) }))
      EXPORT_FORMATS.forEach((f) => list.push({ label: S.palette.exportAs(f.toUpperCase()), hint: h.export, run: () => void exportAs(f) }))
      list.push({ label: S.baselines.save, hint: h.baseline, run: openBaselineDialog })
      open.baselines.forEach((b) => list.push({ label: S.palette.compareTo(b.name), hint: h.baseline, run: () => setWs((s) => ({ ...s, compareId: b.id })) }))
      if (ws.compareId) list.push({ label: S.palette.stopCompare, hint: h.baseline, run: () => setWs((s) => ({ ...s, compareId: '' })) })
      list.push({ label: ws.showDeps ? S.palette.hideDeps : S.palette.showDeps, hint: h.display, run: () => setWs((s) => ({ ...s, showDeps: !s.showDeps })) })
      ;(['day', 'week', 'month'] as const).forEach((z) =>
        list.push({ label: S.palette.zoomTo(z === 'day' ? S.toolbar.zoomDay : z === 'week' ? S.toolbar.zoomWeek : S.toolbar.zoomMonth), hint: h.display, run: () => setWs((s) => ({ ...s, zoom: z })) }),
      )
      list.push({ label: S.palette.jumpToday, hint: h.display, run: () => { go('schedules'); setTodayTick((n) => n + 1) } }) // prettier-ignore
      list.push({ label: S.palette.reimport, hint: h.source, run: () => setImporting({ replace: { id: open.schedule.id, name: open.schedule.name } }) })
    }
    list.push({ label: S.landing.import, hint: h.source, run: () => setImporting({ replace: null }) })
    list.push({ label: S.palette.goTo(S.palette.allSchedules), hint: h.page, run: toList })
    if (open) PAGE_IDS.forEach((p) => list.push({ label: S.palette.goTo(PAGE_LABEL[p]), hint: h.page, run: () => go(p) }))
    schedules?.forEach((m) => list.push({ label: S.palette.openSchedule(m.name), hint: h.schedule, run: () => openSchedule(m.id) }))
    open?.schedule.tasks.forEach((t) =>
      list.push({
        label: open.presentation.tasks[t.id]?.displayLabel || t.name,
        hint: `${h.task} · ${shortDate(t.start)}`,
        run: () => {
          setWs((s) => ({ ...s, selectedId: t.id }))
          setRoute({ ...route, page: 'schedules', view: route.view === 'executive' ? 'detailed' : route.view })
        },
      }),
    )
    return list
  }

  const inWorkspace = !!route.schedule
  const page: PageId = route.page

  const main = (() => {
    if (!inWorkspace) {
      if (page === 'api') return <ApiPage />
      return (
        <Landing
          schedules={schedules}
          loading={listLoading}
          onOpen={openSchedule}
          onImport={() => setImporting({ replace: null })}
          onChanged={loadList}
          toast={toast}
        />
      )
    }
    if (openError)
      return (
        <section className="page">
          <div className="content">
            <div className="panel empty-state" role="alert">
              <h2>{S.states.notFound}</h2>
              <p>{openError}</p>
              <button className="btn primary" onClick={toList}>
                {S.states.backToList}
              </button>
            </div>
          </div>
        </section>
      )
    if (!open)
      return (
        <section className="page">
          <div className="content">
            <div className="panel">
              <Loading />
            </div>
          </div>
        </section>
      )
    switch (page) {
      case 'views':
        return <ViewsPage onOpen={(v) => openView(v)} />
      case 'baselines':
        return (
          <BaselinesPage
            baselines={open.baselines}
            compareId={ws.compareId}
            onSave={openBaselineDialog}
            onCompare={(id) => {
              setWs((s) => ({ ...s, compareId: s.compareId === id ? '' : id }))
              if (ws.compareId !== id) go('schedules')
            }}
            onDelete={deleteBaseline}
          />
        )
      case 'exports':
        return <ExportsPage onExport={exportAs} />
      case 'api':
        return <ApiPage />
      default:
        return (
          <Workspace
            schedule={open.schedule}
            presentation={open.presentation}
            baselines={open.baselines}
            view={route.view}
            focusGroup={route.group}
            state={ws}
            setState={setWs}
            onView={openView}
            onPresentation={(p) => setOpen((o) => (o ? { ...o, presentation: p } : o))}
            todayTick={todayTick}
            toast={toast}
          />
        )
    }
  })()

  const navPages: PageId[] = ['schedules', 'views', 'baselines', 'exports', 'api']

  return (
    <div className="app">
      <nav className="sidebar" aria-label={S.nav.label}>
        <div className="brand">
          <BrandMark />
          {S.appName}
        </div>
        {navPages.map((p) => (
          <button
            key={p}
            className="nav-btn"
            aria-current={(p === 'schedules' ? page === 'schedules' : page === p) ? 'page' : undefined}
            disabled={!inWorkspace && p !== 'schedules' && p !== 'api'}
            onClick={() => (p === 'schedules' && !inWorkspace ? toList() : go(p))}
          >
            {PAGE_LABEL[p]}
          </button>
        ))}
        <div className="side-label">{S.source.label}</div>
        <div className="source-card">
          {open ? (
            <>
              <div className="mono" style={{ color: 'var(--ink)' }}>
                {open.schedule.source.filename}
              </div>
              <div className="muted">{S.source.imported(longDateTime(open.schedule.source.importedAt))}</div>
              <div className="muted">{S.source.readOnly}</div>
              <button className="btn small" style={{ marginTop: 6 }} onClick={() => setImporting({ replace: { id: open.schedule.id, name: open.schedule.name } })}>
                {S.source.reimport}
              </button>
            </>
          ) : (
            <div className="muted">{S.source.none}</div>
          )}
        </div>
      </nav>

      <main>
        <header className="topbar">
          <nav className="crumb" aria-label={S.topbar.breadcrumbLabel}>
            {inWorkspace ? (
              <>
                <button className="crumb-link" onClick={toList}>
                  {S.topbar.breadcrumbRoot}
                </button>
                <span aria-hidden="true">/</span>
                <b aria-current="page">{open?.schedule.name ?? '…'}</b>
              </>
            ) : (
              <b aria-current="page">{page === 'api' ? S.nav.api : S.topbar.breadcrumbRoot}</b>
            )}
          </nav>
          <div className="spacer" />
          <button className="cmd-btn" aria-label={S.topbar.commandMenuLabel} onClick={() => setPaletteOpen(true)}>
            <SearchIcon />
            <span style={{ flex: 1 }}>{S.topbar.commandMenu}</span>
            <span className="kbd">{S.topbar.commandShortcut}</span>
          </button>
          {inWorkspace && open && (
            <>
              <button className="btn" onClick={() => setTodayTick((n) => n + 1)}>
                {S.toolbar.today}
              </button>
              <button className="btn" onClick={openBaselineDialog}>
                {S.topbar.saveBaseline}
              </button>
              <div className="menu-wrap">
                <button ref={exportBtn} className="btn primary" aria-haspopup="menu" aria-expanded={exportMenu} onClick={() => setExportMenu((m) => !m)}>
                  {S.topbar.export}
                </button>
                {exportMenu && (
                  <div className="menu" role="menu">
                    {(
                      [
                        ['png', S.exportsPage.menuPng, true],
                        ['svg', S.exportsPage.menuSvg, true],
                        ['pdf', S.exportsPage.menuPdf, true],
                        ['csv', S.exportsPage.menuCsv, false],
                        ['json', S.exportsPage.menuJson, false],
                      ] as const
                    ).map(([f, label, view]) => (
                      <button key={f} role="menuitem" onClick={() => void exportAs(f)}>
                        {label} {view && <span className="muted">{S.exportsPage.currentView}</span>}
                      </button>
                    ))}
                  </div>
                )}
              </div>
            </>
          )}
          {!inWorkspace && (
            <button className="btn primary" onClick={() => setImporting({ replace: null })}>
              {S.landing.import}
            </button>
          )}
        </header>

        {apiDown && (
          <div className="banner" role="alert">
            <b>{S.states.apiDown}</b>
          </div>
        )}

        {main}
      </main>

      <Palette open={paletteOpen} commands={paletteOpen ? commands() : []} onClose={() => setPaletteOpen(false)} />

      <Modal open={baselineOpen} onClose={() => setBaselineOpen(false)} label={S.baselines.dialogTitle}>
        <form
          className="dlg-body"
          onSubmit={(e) => {
            e.preventDefault()
            void saveBaseline()
          }}
        >
          <h2>{S.baselines.dialogTitle}</h2>
          <label className="field">
            {S.baselines.nameLabel}
            <input value={baselineName} required maxLength={80} autoFocus onFocus={(e) => e.target.select()} onChange={(e) => setBaselineName(e.target.value)} />
          </label>
          <div className="dlg-actions">
            <button className="btn" type="button" onClick={() => setBaselineOpen(false)}>
              {S.baselines.cancel}
            </button>
            <button className="btn primary" type="submit">
              {S.baselines.save}
            </button>
          </div>
        </form>
      </Modal>

      <ImportDialog open={!!importing} replace={importing?.replace ?? null} onClose={() => setImporting(null)} onAccepted={afterImport} />

      <Toasts toasts={toasts} />
    </div>
  )
}
