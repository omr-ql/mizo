import { useEffect, useState } from 'react'
import { openButton } from '../format'
import type { DeskAlert } from '../types'

export function AlertApp() {
  const [alert, setAlert] = useState<DeskAlert | null>(null)

  useEffect(() => window.office.onAlert(setAlert), [])

  if (!alert) return <main className="alert-app" />

  return (
    <main className="alert-app">
      <section className="alert-card">
        <div className="alert-top">
          <strong>{alert.title}</strong>
          <button className="x" onClick={() => void window.office.dismissAlert(alert.id)} aria-label="إغلاق">×</button>
        </div>
        <p>{alert.message}</p>
        <div className="alert-actions">
          <button className="btn" onClick={() => void window.office.dismissAlert(alert.id)}>تم</button>
          <button className="btn primary" onClick={() => void window.office.openAlert(alert)}>{openButton(alert.kind)}</button>
        </div>
      </section>
    </main>
  )
}
