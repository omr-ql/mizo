import { app, BrowserWindow, Tray, Menu, ipcMain, nativeImage, screen } from 'electron'
import { spawn, ChildProcessWithoutNullStreams } from 'child_process'
import fs from 'fs'
import path from 'path'
import { deflateSync } from 'zlib'

const API = 'http://127.0.0.1:5080'
const root = path.join(__dirname, '..')
const dev = process.env.ELECTRON_DEV === '1'

type DeskAlert = {
  id: number
  title: string
  message: string
  source: string
  kind: string
  entryId?: string | null
  storeId?: string | null
}

let mainWindow: BrowserWindow | null = null
let alertWindow: BrowserWindow | null = null
let tray: Tray | null = null
let host: ChildProcessWithoutNullStreams | null = null
let quitting = false
let suppressDismiss = false
let currentAlert: DeskAlert | null = null
const alertQueue: DeskAlert[] = []
const seenAlerts = new Set<number>()
let lastBriefingSlot = ''

const boot = {
  api: false,
  dockerError: null as string | null,
  hostStarted: false
}

function iconBuffer() {
  const size = 32
  const raw = Buffer.alloc((size * 4 + 1) * size)
  for (let y = 0; y < size; y++) {
    const row = y * (size * 4 + 1)
    raw[row] = 0
    for (let x = 0; x < size; x++) {
      const i = row + 1 + x * 4
      const cx = x - size / 2 + 0.5
      const cy = y - size / 2 + 0.5
      const inside = cx * cx + cy * cy < (size / 2 - 1) ** 2
      raw[i] = inside ? 0xc4 : 0x14
      raw[i + 1] = inside ? 0xa4 : 0x12
      raw[i + 2] = inside ? 0x6a : 0x0e
      raw[i + 3] = 255
    }
  }
  const ihdr = Buffer.alloc(13)
  ihdr.writeUInt32BE(size, 0)
  ihdr.writeUInt32BE(size, 4)
  ihdr[8] = 8
  ihdr[9] = 6
  return Buffer.concat([
    Buffer.from([137, 80, 78, 71, 13, 10, 26, 10]),
    chunk('IHDR', ihdr),
    chunk('IDAT', deflateSync(raw)),
    chunk('IEND', Buffer.alloc(0))
  ])
}

function chunk(type: string, data: Buffer) {
  const name = Buffer.from(type)
  const len = Buffer.alloc(4)
  len.writeUInt32BE(data.length)
  const crc = Buffer.alloc(4)
  crc.writeUInt32BE(crc32(Buffer.concat([name, data])))
  return Buffer.concat([len, name, data, crc])
}

function crc32(buf: Buffer) {
  let c = ~0
  for (const byte of buf) {
    c ^= byte
    for (let k = 0; k < 8; k++) c = (c >>> 1) ^ (0xedb88320 & -(c & 1))
  }
  return ~c >>> 0
}

function run(command: string, args: string[]) {
  return new Promise<{ code: number; stdout: string; stderr: string }>((resolve) => {
    const child = spawn(command, args, { cwd: root, windowsHide: true })
    let stdout = ''
    let stderr = ''
    child.stdout?.on('data', (chunk) => { stdout += chunk.toString() })
    child.stderr?.on('data', (chunk) => { stderr += chunk.toString() })
    child.on('error', (error) => resolve({ code: 1, stdout, stderr: error.message }))
    child.on('close', (code) => resolve({ code: code ?? 1, stdout, stderr }))
  })
}

async function api(method: string, apiPath: string, body?: unknown) {
  const response = await fetch(API + apiPath, {
    method,
    headers: body === undefined ? undefined : { 'Content-Type': 'application/json' },
    body: body === undefined ? undefined : JSON.stringify(body)
  })
  const text = await response.text()
  const data = text ? JSON.parse(text) : null
  if (!response.ok) {
    throw new Error(data?.error || `HTTP ${response.status}`)
  }
  return data
}

function pageUrl(file: string) {
  if (dev) return `http://127.0.0.1:5173/${file}`
  return `file:///${path.join(root, 'dist', file).replace(/\\/g, '/')}`
}

function createMainWindow() {
  mainWindow = new BrowserWindow({
    width: 1180,
    height: 760,
    minWidth: 900,
    minHeight: 620,
    title: 'mizo',
    backgroundColor: '#14120e',
    show: false,
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  })
  mainWindow.once('ready-to-show', () => mainWindow?.show())
  mainWindow.on('close', (event) => {
    if (!quitting) {
      event.preventDefault()
      mainWindow?.hide()
    }
  })
  void mainWindow.loadURL(pageUrl('index.html'))
}

function createAlertWindow(): BrowserWindow {
  const area = screen.getPrimaryDisplay().workArea
  const width = 420
  const height = 250
  alertWindow = new BrowserWindow({
    width,
    height,
    x: Math.round(area.x + 16),
    y: Math.round(area.y + area.height - height - 16),
    frame: false,
    resizable: false,
    alwaysOnTop: true,
    skipTaskbar: false,
    show: false,
    backgroundColor: '#1e1a14',
    webPreferences: {
      preload: path.join(__dirname, 'preload.js'),
      contextIsolation: true,
      nodeIntegration: false,
      sandbox: false
    }
  })
  alertWindow.on('closed', () => {
    alertWindow = null
    if (!suppressDismiss && currentAlert) void dismissAlert(currentAlert.id)
    currentAlert = null
    suppressDismiss = false
    showNextAlert()
  })
  void alertWindow.loadURL(pageUrl('alert.html'))
  return alertWindow
}

function enqueueAlert(alert: DeskAlert) {
  if (!alert?.id || seenAlerts.has(alert.id)) return
  seenAlerts.add(alert.id)
  alertQueue.push(alert)
  showNextAlert()
}

function showNextAlert() {
  if (currentAlert || alertQueue.length === 0) return
  currentAlert = alertQueue.shift() ?? null
  if (!currentAlert) return
  const alert = currentAlert
  const deliver = () => {
    alertWindow?.webContents.send('alert-data', alert)
    alertWindow?.show()
  }
  if (!alertWindow) {
    const created = createAlertWindow()
    created.webContents.once('did-finish-load', deliver)
    return
  }
  deliver()
}

async function dismissAlert(id: number) {
  try {
    await api('POST', `/api/alerts/${id}/dismiss`)
  } catch {
    seenAlerts.delete(id)
  }
}

function finishAlert() {
  suppressDismiss = true
  const win = alertWindow
  currentAlert = null
  if (win) win.close()
  else showNextAlert()
}

async function waitForApi() {
  for (let attempt = 0; attempt < 90; attempt++) {
    try {
      const response = await fetch(`${API}/health`)
      if (response.ok) return
    } catch {
      // Docker is still starting.
    }
    await new Promise((resolve) => setTimeout(resolve, 1000))
  }
  throw new Error('تعذر الوصول إلى الخدمة على 127.0.0.1:5080')
}

function startHost() {
  if (quitting || host) return
  const dll = path.join(root, 'outlook-host', 'OfficeDesk.OutlookHost', 'bin', 'Release', 'net10.0-windows', 'OfficeDesk.OutlookHost.dll')
  if (!fs.existsSync(dll)) {
    boot.hostStarted = false
    boot.dockerError = boot.dockerError ?? 'ملف Outlook host غير مبني.'
    return
  }
  const logPath = path.join(app.getPath('userData'), 'outlook-host.log')
  const log = fs.createWriteStream(logPath, { flags: 'a' })
  host = spawn('dotnet', ['exec', dll], {
    cwd: root,
    windowsHide: true,
    env: { ...process.env, OFFICEDESK_API: API }
  })
  boot.hostStarted = true
  host.stdout.pipe(log)
  host.stderr.pipe(log)
  host.on('exit', () => {
    host = null
    boot.hostStarted = false
    if (!quitting) setTimeout(startHost, 5000)
  })
}

async function ensureDockerEngine() {
  const info = await run('docker', ['info'])
  if (info.code === 0) return
  const desktop = 'C:\\Program Files\\Docker\\Docker\\Docker Desktop.exe'
  if (!fs.existsSync(desktop)) {
    throw new Error('Docker Desktop غير مثبت. الخدمة تعمل عبر Docker فقط.')
  }
  spawn(desktop, [], { detached: true, stdio: 'ignore', windowsHide: true }).unref()
  for (let attempt = 0; attempt < 45; attempt++) {
    await new Promise((resolve) => setTimeout(resolve, 2000))
    const again = await run('docker', ['info'])
    if (again.code === 0) return
  }
  throw new Error('Docker لم يعمل بعد. الخدمة تعمل عبر Docker فقط.')
}

async function startServices() {
  try {
    await ensureDockerEngine()
  } catch (error) {
    boot.dockerError = error instanceof Error ? error.message : 'تعذر تشغيل Docker.'
    publishBoot()
    return
  }
  const compose = await run('docker', ['compose', '-p', 'office-desk', '-f', path.join(root, 'docker-compose.yml'), 'up', '-d', '--build'])
  if (compose.code !== 0) {
    boot.dockerError = compose.stderr || compose.stdout || 'تعذر تشغيل Docker.'
    publishBoot()
    return
  }
  try {
    await waitForApi()
    boot.api = true
    boot.dockerError = null
  } catch (error) {
    boot.dockerError = error instanceof Error ? error.message : 'الخدمة المحلية لا تستجيب.'
    publishBoot()
    return
  }
  startHost()
  try {
    const settings = await api('GET', '/api/settings')
    app.setLoginItemSettings({ openAtLogin: Boolean(settings.startWithWindows) })
  } catch {
    // Settings appear after the first successful call.
  }
  publishBoot()
  setInterval(() => { void pollAlerts() }, 10000)
  setInterval(() => { void maybeBriefing() }, 20000)
  void pollAlerts()
}

async function pollAlerts() {
  if (!boot.api) return
  try {
    const alerts = await api('GET', '/api/alerts/pending') as DeskAlert[]
    for (const alert of alerts) enqueueAlert(alert)
  } catch {
    // The next poll retries.
  }
}

async function maybeBriefing() {
  if (!boot.api) return
  const parts = Object.fromEntries(new Intl.DateTimeFormat('en-GB', {
    timeZone: 'Asia/Riyadh',
    weekday: 'long',
    year: 'numeric',
    month: '2-digit',
    day: '2-digit',
    hour: '2-digit',
    minute: '2-digit',
    hourCycle: 'h23'
  }).formatToParts(new Date()).map((part) => [part.type, part.value]))
  const minute = Number(parts.minute)
  const slot = parts.hour === '08' && minute < 5
    ? 'morning'
    : parts.hour === '15' && minute < 5
      ? 'afternoon'
      : ''
  if (!slot) return
  const key = `${parts.year}-${parts.month}-${parts.day}-${slot}`
  if (key === lastBriefingSlot) return
  lastBriefingSlot = key
  try {
    const result = await api('POST', '/api/briefing/generate', { quietIfEmpty: true })
    if (result?.alert) enqueueAlert(result.alert)
  } catch {
    lastBriefingSlot = ''
  }
}

function publishBoot() {
  mainWindow?.webContents.send('boot-status', boot)
}

function createTray() {
  tray = new Tray(nativeImage.createFromBuffer(iconBuffer()))
  tray.setToolTip('mizo')
  tray.setContextMenu(Menu.buildFromTemplate([
    { label: 'فتح', click: () => { mainWindow?.show(); mainWindow?.focus() } },
    { label: 'إنهاء', click: () => { quitting = true; app.quit() } }
  ]))
  tray.on('click', () => { mainWindow?.show(); mainWindow?.focus() })
}

const gotLock = app.requestSingleInstanceLock()
if (!gotLock) {
  app.quit()
} else {
  app.on('second-instance', () => { mainWindow?.show(); mainWindow?.focus() })
  app.whenReady().then(() => {
    app.setAppUserModelId('com.ejadah.office-desk')
    createMainWindow()
    createTray()
    void startServices()
  })
}

app.on('before-quit', () => {
  quitting = true
  if (host && !host.killed) host.kill()
})

ipcMain.handle('boot-status', () => boot)

ipcMain.handle('api-request', async (_event, payload: { method: string; path: string; body?: unknown }) => {
  return api(payload.method, payload.path, payload.body)
})

ipcMain.handle('set-autostart', (_event, on: boolean) => {
  app.setLoginItemSettings({ openAtLogin: on })
  return app.getLoginItemSettings().openAtLogin
})

ipcMain.handle('alert-dismiss', async (_event, id: number) => {
  await dismissAlert(id)
  if (currentAlert?.id === id) finishAlert()
})

ipcMain.handle('alert-open', async (_event, alert: DeskAlert) => {
  if (alert.kind === 'briefing') {
    mainWindow?.show()
    mainWindow?.focus()
    mainWindow?.webContents.send('navigate', 'briefing')
  } else {
    await api('POST', `/api/alerts/${alert.id}/open`)
  }
  await dismissAlert(alert.id)
  if (currentAlert?.id === alert.id) finishAlert()
})
