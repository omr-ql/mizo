import type { BootStatus, DeskAlert } from './types'

export interface OfficeBridge {
  request<T = unknown>(method: string, path: string, body?: unknown): Promise<T>
  getBoot(): Promise<BootStatus>
  setAutostart(on: boolean): Promise<boolean>
  onNavigate(cb: (screen: string) => void): () => void
  onAlert(cb: (alert: DeskAlert) => void): () => void
  dismissAlert(id: number): Promise<void>
  openAlert(alert: DeskAlert): Promise<void>
}

declare global {
  interface Window {
    office: OfficeBridge
  }
}

export {}
