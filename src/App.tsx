import { useEffect, useState } from 'react'
import { BriefingScreen, DraftScreen, MailScreen, PendingScreen, SettingsScreen, SummaryScreen } from './screens'
import logo from './assets/logo.png'
import type { BootStatus, DraftSeed, HostStatus, ScreenName, Settings, ThreadRow } from './types'

const nav: { id: ScreenName; label: string; icon: string }[] = [
  { id: 'briefing', label: 'إحاطة', icon: 'grid' },
  { id: 'mail', label: 'البريد', icon: 'mail' },
  { id: 'summary', label: 'ملخص', icon: 'note' },
  { id: 'draft', label: 'مسودة', icon: 'pen' },
  { id: 'pending', label: 'معلّق', icon: 'clock' },
  { id: 'settings', label: 'إعدادات', icon: 'gear' }
]

function Icon({ name }: { name: string }) {
  const common = { width: 18, height: 18, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.8 }
  if (name === 'mail') return <svg {...common}><rect x="3" y="5" width="18" height="14" rx="2" /><path d="m4 7 8 6 8-6" /></svg>
  if (name === 'note') return <svg {...common}><path d="M7 3.5h7.5L19 8v12.5H7V3.5Z" /><path d="M14.5 3.5V8H19" /><path d="M10 12h5M10 16h5" /></svg>
  if (name === 'pen') return <svg {...common}><path d="M4 20h4l10-10-4-4L4 16v4Z" /><path d="m13 7 4 4" /></svg>
  if (name === 'clock') return <svg {...common}><circle cx="12" cy="12" r="8" /><path d="M12 8v5l3 2" /></svg>
  if (name === 'gear') return <svg {...common}><circle cx="12" cy="12" r="3" /><path d="M12 3v2m0 14v2M3 12h2m14 0h2M5.6 5.6l1.4 1.4m10 10 1.4 1.4m0-12.8-1.4 1.4m-10 10-1.4 1.4" /></svg>
  return <svg {...common}><rect x="4" y="4" width="6" height="6" rx="1.5" /><rect x="14" y="4" width="6" height="6" rx="1.5" /><rect x="4" y="14" width="6" height="6" rx="1.5" /><rect x="14" y="14" width="6" height="6" rx="1.5" /></svg>
}

const titles: Record<ScreenName, string> = {
  briefing: 'إحاطة',
  mail: 'البريد',
  summary: 'ملخص',
  draft: 'مسودة',
  pending: 'معلّق',
  settings: 'إعدادات'
}

export function App() {
  const [screen, setScreen] = useState<ScreenName>('briefing')
  const [boot, setBoot] = useState<BootStatus | null>(null)
  const [host, setHost] = useState<HostStatus | null>(null)
  const [seed, setSeed] = useState<DraftSeed | null>(null)
  const [account, setAccount] = useState('')
  const [knownAccount, setKnownAccount] = useState(false)
  const [setupEmail, setSetupEmail] = useState('')
  const [setupError, setSetupError] = useState('')

  useEffect(() => {
    if (!window.office) return
    void window.office.getBoot().then(setBoot)
    const timer = window.setInterval(() => {
      window.office.getBoot().then(setBoot).catch(() => undefined)
      window.office.request<HostStatus>('GET', '/api/host/status').then(setHost).catch(() => undefined)
    }, 4000)
    return () => window.clearInterval(timer)
  }, [])

  useEffect(() => {
    if (!window.office) return
    return window.office.onNavigate((next) => {
      if (next in titles) setScreen(next as ScreenName)
    })
  }, [])

  useEffect(() => {
    if (!window.office || !boot?.api) return
    window.office.request<Settings>('GET', '/api/settings').then((settings) => {
      setAccount(settings.accountSmtp.trim())
      setKnownAccount(true)
    }).catch(() => undefined)
  }, [boot?.api])

  async function saveAccount() {
    const email = setupEmail.trim()
    if (!email.includes('@') || email.includes(' ')) {
      setSetupError('اكتب بريد Outlook كما يظهر في الحساب.')
      return
    }
    try {
      const current = await window.office.request<Settings>('GET', '/api/settings')
      await window.office.request('PUT', '/api/settings', { ...current, accountSmtp: email })
      setAccount(email)
      setSetupError('')
    } catch (err) {
      setSetupError(err instanceof Error ? err.message : 'تعذر حفظ البريد.')
    }
  }

  function draftFrom(thread: ThreadRow) {
    const subject = thread.subject.startsWith('Re:') ? thread.subject : `Re: ${thread.subject}`
    setSeed({
      threadId: thread.id,
      entryId: thread.entryId,
      storeId: thread.storeId,
      to: thread.counterpartAddress,
      subject,
      body: ''
    })
    setScreen('draft')
  }

  if (!window.office) {
    return <main className="main"><p>افتح mizo من تطبيق سطح المكتب.</p></main>
  }

  if (!knownAccount || !account) {
    return (
      <main className="setup">
        <form onSubmit={(event) => { event.preventDefault(); void saveAccount() }}>
          <img src={logo} alt="" />
          <h1>mizo</h1>
          <p>اكتب بريد Outlook على هذا الجهاز. التطبيق يقرأ هذا الحساب فقط.</p>
          {boot?.dockerError && <div className="banner">{boot.dockerError}</div>}
          <label>البريد
            <input value={setupEmail} onChange={(event) => setSetupEmail(event.target.value)} placeholder="name@company.com" autoFocus />
          </label>
          <button className="btn primary" type="submit" disabled={!boot?.api}>ابدأ</button>
          {setupError && <p className="banner">{setupError}</p>}
          {!boot?.api && !boot?.dockerError && <p className="muted">جارٍ تجهيز الخدمة…</p>}
        </form>
      </main>
    )
  }

  return (
    <div className="shell">
      <aside className="sidebar">
        <div className="logo"><img src={logo} alt="" /> mizo</div>
        {nav.map((item) => (
          <button key={item.id} className={screen === item.id ? 'nav active' : 'nav'} onClick={() => setScreen(item.id)}>
            <Icon name={item.icon} />
            {item.label}
          </button>
        ))}
        <div className="side-foot">
          <div className="avatar">{account.slice(0, 1).toUpperCase()}</div>
          <div>
            <strong>{account}</strong>
            <span>{host?.ok ? 'Outlook' : boot?.api ? 'Docker' : 'بانتظار الاتصال'}</span>
          </div>
        </div>
      </aside>
      <div className="canvas">
        <main className="panel">
          <header className="top">
            <div>
              <h1>{titles[screen]}</h1>
              <p>{host?.message || 'حساب العمل فقط، والإرسال بعد «أرسل».'}</p>
            </div>
          </header>
          {boot?.dockerError && <div className="banner">{boot.dockerError}</div>}
          {screen === 'briefing' && <BriefingScreen onDraft={draftFrom} />}
          {screen === 'mail' && <MailScreen />}
          {screen === 'summary' && <SummaryScreen onDraft={draftFrom} />}
          {screen === 'draft' && <DraftScreen seed={seed} />}
          {screen === 'pending' && <PendingScreen />}
          {screen === 'settings' && <SettingsScreen />}
        </main>
      </div>
    </div>
  )
}
