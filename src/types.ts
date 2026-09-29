export type ScreenName = 'briefing' | 'mail' | 'threads' | 'draft' | 'pending' | 'settings'

export type BootStatus = {
  api: boolean
  dockerError: string | null
  hostStarted: boolean
}

export type HostStatus = {
  ok: boolean
  message: string
  updatedAt: string
}

export type Settings = {
  displayName: string
  accountSmtp: string
  workDays: string
  workStart: string
  workEnd: string
  urgentWords: string
  pollIntervalMinutes: number
  startWithWindows: boolean
}

export type MailRow = {
  id: number
  entryId: string
  storeId: string
  folder: string
  subject: string
  sender: string
  senderAddress: string
  receivedAt: string
  preview: string
  unread: boolean
  classification: string
}

export type ThreadRow = {
  id: number
  subject: string
  state: string
  recommendedStep: string
  entryId: string
  storeId: string
  counterpartName: string
  counterpartAddress: string
  lastActivityAt: string
}

export type BriefingItem = ThreadRow & { threadId: number }

export type Briefing = {
  needsYou: BriefingItem[]
  waiting: BriefingItem[]
  info: BriefingItem[]
  generatedAt: string
}

export type DraftRow = {
  id: number
  threadId: number | null
  entryId: string | null
  storeId: string | null
  to: string
  subject: string
  body: string
  status: string
  updatedAt: string
}

export type PendingRow = {
  id: number
  title: string
  notes: string
  done: boolean
  createdAt: string
}

export type DeskAlert = {
  id: number
  title: string
  message: string
  source: string
  kind: string
  entryId?: string | null
  storeId?: string | null
}

export type DraftSeed = {
  threadId?: number
  entryId?: string
  storeId?: string
  to: string
  subject: string
  body: string
}
