import { contextBridge, ipcRenderer } from 'electron'

contextBridge.exposeInMainWorld('office', {
  request: (method: string, path: string, body?: unknown) =>
    ipcRenderer.invoke('api-request', { method, path, body }),
  getBoot: () => ipcRenderer.invoke('boot-status'),
  setAutostart: (on: boolean) => ipcRenderer.invoke('set-autostart', on),
  onNavigate: (cb: (screen: string) => void) => {
    const listener = (_event: unknown, screen: string) => cb(screen)
    ipcRenderer.on('navigate', listener)
    return () => ipcRenderer.removeListener('navigate', listener)
  },
  onAlert: (cb: (alert: unknown) => void) => {
    const listener = (_event: unknown, alert: unknown) => cb(alert)
    ipcRenderer.on('alert-data', listener)
    return () => ipcRenderer.removeListener('alert-data', listener)
  },
  dismissAlert: (id: number) => ipcRenderer.invoke('alert-dismiss', id),
  openAlert: (alert: unknown) => ipcRenderer.invoke('alert-open', alert)
})
