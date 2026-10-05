import { useEffect, useState } from 'react'
import { draftName, stateName, when } from './format'
import type { Briefing, DraftRow, DraftSeed, MailRow, PendingRow, Settings, ThreadRow } from './types'

async function load<T>(path: string): Promise<T> {
  return window.office.request<T>('GET', path)
}

export function BriefingScreen({ onDraft }: { onDraft: (thread: ThreadRow) => void }) {
  const [data, setData] = useState<Briefing | null>(null)
  const [error, setError] = useState('')

  useEffect(() => {
    load<Briefing>('/api/briefing').then(setData).catch((err: Error) => setError(err.message))
  }, [])

  if (error) return <p className="banner">{error}</p>
  if (!data) return <p className="muted">جارٍ التحميل…</p>

  const columns = [
    ['يحتاجك', data.needsYou],
    ['بانتظارهم', data.waiting],
    ['للعلم', data.info]
  ] as const

  const rows = columns.flatMap(([title, items]) => items.map((item) => ({ title, item })))

  return (
    <>
      <div className="stats">
        {columns.map(([title, items], index) => (
          <article className={index === 0 ? 'stat hot' : 'stat'} key={title}>
            <span>{title}</span>
            <b>{items.length}</b>
          </article>
        ))}
      </div>
      {rows.length === 0 && <p className="empty">لا شيء هنا.</p>}
      {rows.length > 0 && (
        <div className="table-wrap">
          <table>
            <colgroup>
              <col className="w-kind" /><col /><col className="w-step" /><col className="w-when" /><col className="w-btn" />
            </colgroup>
            <thead>
              <tr><th>النوع</th><th>الموضوع</th><th>الخطوة</th><th>النشاط</th><th className="fit"></th></tr>
            </thead>
            <tbody>
              {rows.map(({ title, item }) => (
                <tr key={item.threadId}>
                  <td><span className="badge">{title}</span></td>
                  <td className="subject" title={item.subject}><bdi>{item.subject}</bdi></td>
                  <td className="step" title={item.recommendedStep}>{item.recommendedStep}</td>
                  <td className="when">{when(item.lastActivityAt)}</td>
                  <td className="fit">
                    {item.state === 'awaiting_me' && (
                      <button className="btn primary" onClick={() => onDraft({ ...item, id: item.threadId })}>مسودة</button>
                    )}
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>
  )
}

export function MailScreen() {
  const [folder, setFolder] = useState<'inbox' | 'sent'>('inbox')
  const [rows, setRows] = useState<MailRow[]>([])
  const [error, setError] = useState('')

  useEffect(() => {
    load<MailRow[]>(`/api/mail?folder=${folder}`).then(setRows).catch((err: Error) => setError(err.message))
  }, [folder])

  async function open(row: MailRow) {
    await window.office.request('POST', '/api/mail/open', { entryId: row.entryId, storeId: row.storeId })
  }

  return (
    <>
      <div className="tabs">
        <button className="btn" aria-pressed={folder === 'inbox'} onClick={() => setFolder('inbox')}>الوارد</button>
        <button className="btn" aria-pressed={folder === 'sent'} onClick={() => setFolder('sent')}>الصادر</button>
      </div>
      {error && <p className="banner">{error}</p>}
      {rows.length === 0 && <p className="empty">لا رسائل في هذا المجلد بعد.</p>}
      {rows.length > 0 && (
        <div className="table-wrap">
          <table>
            <colgroup>
              <col /><col className="w-who" /><col className="w-when" /><col className="w-open" />
            </colgroup>
            <thead>
              <tr><th>الموضوع</th><th>المرسل</th><th>الوقت</th><th className="fit"></th></tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.id}>
                  <td className="stack" title={row.subject}><strong><bdi>{row.subject || '(بدون عنوان)'}</bdi></strong><div className="muted"><bdi>{row.preview}</bdi></div></td>
                  <td title={row.sender}>{row.sender}{row.unread ? ' · غير مقروء' : ''}</td>
                  <td className="when">{when(row.receivedAt)}</td>
                  <td className="fit"><button className="btn" onClick={() => void open(row)}>افتح البريد</button></td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>
  )
}

export function ThreadsScreen({ onDraft }: { onDraft: (thread: ThreadRow) => void }) {
  const [rows, setRows] = useState<ThreadRow[]>([])
  const [error, setError] = useState('')

  function refresh() {
    load<ThreadRow[]>('/api/threads').then(setRows).catch((err: Error) => setError(err.message))
  }

  useEffect(() => { refresh() }, [])

  async function closeThread(id: number) {
    await window.office.request('PATCH', `/api/threads/${id}`, { state: 'closed' })
    refresh()
  }

  async function open(row: ThreadRow) {
    await window.office.request('POST', '/api/mail/open', { entryId: row.entryId, storeId: row.storeId })
  }

  return (
    <>
      {error && <p className="banner">{error}</p>}
      {rows.length === 0 && <p className="empty">لا خيوط بعد. ستظهر بعد أول مزامنة من Outlook.</p>}
      {rows.length > 0 && (
        <div className="table-wrap">
          <table>
            <colgroup>
              <col /><col className="w-state" /><col className="w-step" /><col className="w-when" /><col className="w-actions" />
            </colgroup>
            <thead>
              <tr><th>الموضوع</th><th>الحالة</th><th>الخطوة</th><th>النشاط</th><th className="fit"></th></tr>
            </thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.id}>
                  <td className="subject" title={row.subject}><bdi>{row.subject}</bdi></td>
                  <td><span className="badge">{stateName(row.state)}</span></td>
                  <td className="step" title={row.recommendedStep}>{row.recommendedStep}</td>
                  <td className="when">{when(row.lastActivityAt)}</td>
                  <td className="fit">
                    <div className="actions">
                      <button className="btn" onClick={() => void open(row)}>افتح البريد</button>
                      {row.state === 'awaiting_me' && <button className="btn primary" onClick={() => onDraft(row)}>مسودة</button>}
                      {row.state !== 'closed' && <button className="btn ghost" onClick={() => void closeThread(row.id)}>إغلاق</button>}
                    </div>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>
  )
}

export function DraftScreen({ seed }: { seed: DraftSeed | null }) {
  const [rows, setRows] = useState<DraftRow[]>([])
  const [currentId, setCurrentId] = useState<number | null>(null)
  const [to, setTo] = useState(seed?.to ?? '')
  const [subject, setSubject] = useState(seed?.subject ?? '')
  const [body, setBody] = useState(seed?.body ?? '')
  const [entryId, setEntryId] = useState(seed?.entryId)
  const [storeId, setStoreId] = useState(seed?.storeId)
  const [threadId, setThreadId] = useState(seed?.threadId)
  const [note, setNote] = useState('')

  function refresh() {
    load<DraftRow[]>('/api/drafts').then(setRows).catch((err: Error) => setNote(err.message))
  }

  useEffect(() => { refresh() }, [])
  useEffect(() => {
    setTo(seed?.to ?? '')
    setSubject(seed?.subject ?? '')
    setBody(seed?.body ?? '')
    setEntryId(seed?.entryId)
    setStoreId(seed?.storeId)
    setThreadId(seed?.threadId)
    setCurrentId(null)
  }, [seed])

  function payload() {
    return { threadId, entryId, storeId, to, subject, body }
  }

  async function saveLocal() {
    if (currentId) {
      await window.office.request('PUT', `/api/drafts/${currentId}`, payload())
    } else {
      const created = await window.office.request<DraftRow>('POST', '/api/drafts', payload())
      setCurrentId(created.id)
    }
    setNote('حُفظ النص في التطبيق.')
    refresh()
  }

  async function send() {
    try {
      let id = currentId
      if (!id) {
        const created = await window.office.request<DraftRow>('POST', '/api/drafts', payload())
        id = created.id
        setCurrentId(id)
      } else {
        await window.office.request('PUT', `/api/drafts/${id}`, payload())
      }
      await window.office.request('POST', `/api/drafts/${id}/send`)
      setNote('طُلب الإرسال. لن يخرج البريد إلا عبر Outlook على هذا الجهاز.')
      refresh()
    } catch (err) {
      setNote(err instanceof Error ? err.message : 'تعذر الإرسال')
    }
  }

  async function saveOutlook() {
    let id = currentId
    if (!id) {
      const created = await window.office.request<DraftRow>('POST', '/api/drafts', payload())
      id = created.id
      setCurrentId(id)
    } else {
      await window.office.request('PUT', `/api/drafts/${id}`, payload())
    }
    await window.office.request('POST', `/api/drafts/${id}/save`)
    setNote('طُلب حفظ المسودة داخل Outlook.')
    refresh()
  }

  function select(row: DraftRow) {
    setCurrentId(row.id)
    setTo(row.to)
    setSubject(row.subject)
    setBody(row.body)
    setEntryId(row.entryId ?? undefined)
    setStoreId(row.storeId ?? undefined)
    setThreadId(row.threadId ?? undefined)
    setNote('')
  }

  return (
    <div className="split">
      <aside className="draft-list">
        {rows.length === 0 && <p className="empty">لا مسودات محفوظة.</p>}
        {rows.map((row) => (
          <button className="btn" key={row.id} onClick={() => select(row)}>
            {row.subject || '(بدون عنوان)'}
            <div className="muted">{draftName(row.status)}</div>
          </button>
        ))}
      </aside>
      <form className="composer" onSubmit={(event) => event.preventDefault()}>
        <label>إلى<input value={to} onChange={(event) => setTo(event.target.value)} /></label>
        <label>الموضوع<input value={subject} onChange={(event) => setSubject(event.target.value)} /></label>
        <label>النص<textarea value={body} onChange={(event) => setBody(event.target.value)} /></label>
        <div className="actions">
          <button className="btn primary" onClick={() => void send()}>أرسل</button>
          <button className="btn" onClick={() => void saveLocal()}>احفظ هنا</button>
          <button className="btn ghost" onClick={() => void saveOutlook()}>احفظ في مسودات Outlook</button>
        </div>
        {note && <p className="muted">{note}</p>}
      </form>
    </div>
  )
}

export function PendingScreen() {
  const [rows, setRows] = useState<PendingRow[]>([])
  const [title, setTitle] = useState('')
  const [error, setError] = useState('')

  function refresh() {
    load<PendingRow[]>('/api/pending').then(setRows).catch((err: Error) => setError(err.message))
  }

  useEffect(() => { refresh() }, [])

  async function add() {
    if (!title.trim()) return
    await window.office.request('POST', '/api/pending', { title })
    setTitle('')
    refresh()
  }

  async function toggle(row: PendingRow) {
    await window.office.request('PATCH', `/api/pending/${row.id}`, { done: !row.done })
    refresh()
  }

  return (
    <>
      <form className="actions" onSubmit={(event) => { event.preventDefault(); void add() }}>
        <input value={title} onChange={(event) => setTitle(event.target.value)} placeholder="مهمة معلّقة" />
        <button className="btn primary" type="submit">إضافة</button>
      </form>
      {error && <p className="banner">{error}</p>}
      {rows.length === 0 && <p className="empty">لا مهام معلّقة.</p>}
      {rows.length > 0 && (
        <div className="table-wrap">
          <table>
            <thead><tr><th>المهمة</th><th>الحالة</th></tr></thead>
            <tbody>
              {rows.map((row) => (
                <tr key={row.id}>
                  <td><strong dir="auto">{row.title}</strong></td>
                  <td>
                    <label className="checks">
                      <input type="checkbox" checked={row.done} onChange={() => void toggle(row)} />
                      <span>{row.done ? 'تمت' : 'معلّقة'}</span>
                    </label>
                  </td>
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      )}
    </>
  )
}

const dayOptions = [
  ['Sunday', 'الأحد'],
  ['Monday', 'الاثنين'],
  ['Tuesday', 'الثلاثاء'],
  ['Wednesday', 'الأربعاء'],
  ['Thursday', 'الخميس'],
  ['Friday', 'الجمعة'],
  ['Saturday', 'السبت']
] as const

export function SettingsScreen() {
  const [form, setForm] = useState<Settings | null>(null)
  const [note, setNote] = useState('')

  useEffect(() => {
    load<Settings>('/api/settings').then(setForm).catch((err: Error) => setNote(err.message))
  }, [])

  if (!form) return <p className="muted">{note || 'جارٍ التحميل…'}</p>

  const current = form
  const selected = new Set(current.workDays.split(',').map((day) => day.trim()).filter(Boolean))

  function toggleDay(day: string) {
    const next = new Set(selected)
    if (next.has(day)) next.delete(day)
    else next.add(day)
    setForm({ ...current, workDays: dayOptions.map(([key]) => key).filter((key) => next.has(key)).join(',') })
  }

  async function save() {
    await window.office.request('PUT', '/api/settings', current)
    await window.office.setAutostart(current.startWithWindows)
    setNote('حُفظت الإعدادات.')
  }

  return (
    <form className="form" onSubmit={(event) => { event.preventDefault(); void save() }}>
      <label>الاسم<input value={form.displayName} onChange={(event) => setForm({ ...form, displayName: event.target.value })} /></label>
      <label>حساب Outlook<input value={form.accountSmtp} onChange={(event) => setForm({ ...form, accountSmtp: event.target.value })} /></label>
      <label>بداية الدوام<input value={form.workStart} onChange={(event) => setForm({ ...form, workStart: event.target.value })} /></label>
      <label>نهاية الدوام<input value={form.workEnd} onChange={(event) => setForm({ ...form, workEnd: event.target.value })} /></label>
      <label>كلمات عاجلة<input value={form.urgentWords} onChange={(event) => setForm({ ...form, urgentWords: event.target.value })} /></label>
      <div className="checks">
        {dayOptions.map(([key, label]) => (
          <label key={key}><input type="checkbox" checked={selected.has(key)} onChange={() => toggleDay(key)} />{label}</label>
        ))}
      </div>
      <label className="checks">
        <input type="checkbox" checked={form.startWithWindows} onChange={(event) => setForm({ ...form, startWithWindows: event.target.checked })} />
        التشغيل مع ويندوز
      </label>
      <div className="actions"><button className="btn primary" type="submit">حفظ</button></div>
      {note && <p className="muted">{note}</p>}
    </form>
  )
}
