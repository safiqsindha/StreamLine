import { useState } from 'react'
import type { ScheduleMeta } from '../../core/model'
import { api } from '../api'
import { longDateTime } from '../format'
import { S } from '../strings'
import { Loading, Modal } from './ui'

export function Landing({
  schedules,
  loading,
  onOpen,
  onImport,
  onChanged,
  toast,
}: {
  schedules: ScheduleMeta[] | null
  loading: boolean
  onOpen: (id: string) => void
  onImport: () => void
  onChanged: () => void
  toast: (msg: string, err?: boolean) => void
}) {
  const [renaming, setRenaming] = useState<ScheduleMeta | null>(null)
  const [newName, setNewName] = useState('')
  const [deleting, setDeleting] = useState<{ meta: ScheduleMeta; baselines: number } | null>(null)

  const startDelete = async (m: ScheduleMeta) => {
    try {
      const b = await api.listBaselines(m.id)
      setDeleting({ meta: m, baselines: b.length })
    } catch (e) {
      toast((e as Error).message, true)
    }
  }

  const doRename = async () => {
    if (!renaming || !newName.trim()) return
    try {
      const m = await api.renameSchedule(renaming.id, newName.trim())
      toast(S.landing.renamed(m.name))
      setRenaming(null)
      onChanged()
    } catch (e) {
      toast((e as Error).message, true)
    }
  }

  const doDelete = async () => {
    if (!deleting) return
    try {
      await api.deleteSchedule(deleting.meta.id)
      toast(S.landing.deleted(deleting.meta.name))
      setDeleting(null)
      onChanged()
    } catch (e) {
      toast((e as Error).message, true)
    }
  }

  return (
    <section className="page" aria-label={S.landing.title}>
      <div className="content">
        <div className="panel">
          <div className="panel-head" style={{ display: 'flex', flexWrap: 'wrap', alignItems: 'center', gap: 10 }}>
            <div style={{ flex: '1 1 260px' }}>
              <h2>{S.landing.title}</h2>
              <p>{S.landing.subtitle}</p>
            </div>
            <button className="btn primary" onClick={onImport}>
              {S.landing.import}
            </button>
          </div>

          {loading || !schedules ? (
            <Loading />
          ) : schedules.length === 0 ? (
            <div className="empty-state">
              <h2>{S.landing.emptyTitle}</h2>
              <p>{S.landing.emptyBody}</p>
              <button className="btn primary" onClick={onImport}>
                {S.landing.import}
              </button>
            </div>
          ) : (
            <div className="table-box">
              <table>
                <thead>
                  <tr>
                    <th>{S.landing.colName}</th>
                    <th>{S.landing.colFile}</th>
                    <th>{S.landing.colImported}</th>
                    <th>{S.landing.colTasks}</th>
                    <th>
                      <span className="sr-only-text">{S.landing.open}</span>
                    </th>
                  </tr>
                </thead>
                <tbody>
                  {schedules.map((m) => (
                    <tr key={m.id}>
                      <td>
                        <button className="card-open" onClick={() => onOpen(m.id)}>
                          {m.name}
                        </button>
                      </td>
                      <td className="mono">{m.source.filename}</td>
                      <td>{longDateTime(m.source.importedAt)}</td>
                      <td className="mono">{m.taskCount}</td>
                      <td style={{ whiteSpace: 'nowrap' }}>
                        <div className="sim-row">
                          <button className="btn small" onClick={() => onOpen(m.id)}>
                            {S.landing.open}
                          </button>
                          <button
                            className="btn small"
                            onClick={() => {
                              setRenaming(m)
                              setNewName(m.name)
                            }}
                          >
                            {S.landing.rename}
                          </button>
                          <button className="btn small" onClick={() => startDelete(m)}>
                            {S.landing.delete}
                          </button>
                        </div>
                      </td>
                    </tr>
                  ))}
                </tbody>
              </table>
            </div>
          )}
        </div>
      </div>

      <Modal open={!!renaming} onClose={() => setRenaming(null)} label={S.landing.renameTitle}>
        <form
          className="dlg-body"
          onSubmit={(e) => {
            e.preventDefault()
            void doRename()
          }}
        >
          <h2>{S.landing.renameTitle}</h2>
          <label className="field">
            {S.landing.renameLabel}
            <input value={newName} required maxLength={120} autoFocus onChange={(e) => setNewName(e.target.value)} />
          </label>
          <div className="dlg-actions">
            <button className="btn" type="button" onClick={() => setRenaming(null)}>
              {S.baselines.cancel}
            </button>
            <button className="btn primary" type="submit">
              {S.landing.renameSave}
            </button>
          </div>
        </form>
      </Modal>

      <Modal open={!!deleting} onClose={() => setDeleting(null)} label={deleting ? S.landing.deleteTitle(deleting.meta.name) : ''}>
        {deleting && (
          <div className="dlg-body">
            <h2>{S.landing.deleteTitle(deleting.meta.name)}</h2>
            <p>{S.landing.deleteBody(deleting.meta.name, deleting.meta.source.filename, deleting.baselines)}</p>
            <div className="dlg-actions">
              <button className="btn" type="button" autoFocus onClick={() => setDeleting(null)}>
                {S.baselines.cancel}
              </button>
              <button className="btn danger" type="button" onClick={doDelete}>
                {S.landing.deleteConfirm}
              </button>
            </div>
          </div>
        )}
      </Modal>
    </section>
  )
}
