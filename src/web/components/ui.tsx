import { useCallback, useEffect, useRef, useState, type ReactNode } from 'react'
import { S } from '../strings'

export function BrandMark() {
  return (
    <span className="brand-mark" aria-hidden="true">
      <svg width="14" height="14" viewBox="0 0 14 14" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round">
        <path d="M2 4h6M5 7h7M2 10h5" />
      </svg>
    </span>
  )
}

export function SearchIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 16 16" fill="none" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" aria-hidden="true">
      <circle cx="7" cy="7" r="4.5" />
      <path d="M10.5 10.5L14 14" />
    </svg>
  )
}

export function CloseIcon() {
  return (
    <svg width="14" height="14" viewBox="0 0 14 14" stroke="currentColor" strokeWidth="1.6" strokeLinecap="round" aria-hidden="true">
      <path d="M3 3l8 8M11 3l-8 8" />
    </svg>
  )
}

export function Chevron({ closed }: { closed: boolean }) {
  return (
    <svg className={`chev ${closed ? 'closed' : ''}`} width="12" height="12" viewBox="0 0 12 12" fill="none" stroke="currentColor" strokeWidth="1.6" aria-hidden="true">
      <path d="M3 4.5l3 3 3-3" />
    </svg>
  )
}

export function Loading({ label = S.states.loading }: { label?: string }) {
  return (
    <div className="loading" role="status" aria-live="polite">
      <span className="spinner" aria-hidden="true" />
      {label}
    </div>
  )
}

export function Unavailable() {
  return <span className="na">{S.unavailable}</span>
}

/** Modal dialog with native dialog semantics; Escape closes via the browser's cancel event. */
export function Modal({
  open,
  onClose,
  label,
  wide,
  children,
}: {
  open: boolean
  onClose: () => void
  label: string
  wide?: boolean
  children: ReactNode
}) {
  const ref = useRef<HTMLDialogElement>(null)
  useEffect(() => {
    const d = ref.current
    if (!d) return
    if (open && !d.open) d.showModal()
    if (!open && d.open) d.close()
  }, [open])
  return (
    <dialog
      ref={ref}
      aria-label={label}
      className={wide ? 'wide' : undefined}
      onCancel={(e) => {
        e.preventDefault()
        onClose()
      }}
      onClick={(e) => {
        if (e.target === ref.current) onClose()
      }}
    >
      {open ? children : null}
    </dialog>
  )
}

export interface Toast {
  id: number
  message: string
  error?: boolean
}

export function useToasts() {
  const [toasts, setToasts] = useState<Toast[]>([])
  const next = useRef(1)
  const push = useCallback((message: string, error?: boolean) => {
    const id = next.current++
    setToasts((t) => [...t.slice(-2), { id, message, error }])
    setTimeout(() => setToasts((t) => t.filter((x) => x.id !== id)), error ? 6000 : 3200)
  }, [])
  return { toasts, push }
}

export function Toasts({ toasts }: { toasts: Toast[] }) {
  return (
    <div className="toast-box" aria-live="polite">
      {toasts.map((t) => (
        <div key={t.id} className={`toast ${t.error ? 'err' : ''}`} role={t.error ? 'alert' : undefined}>
          {t.message}
        </div>
      ))}
    </div>
  )
}
