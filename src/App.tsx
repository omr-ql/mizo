import { useEffect, useState } from 'react'
import { BriefingScreen, DraftScreen, MailScreen, PendingScreen, SettingsScreen, ThreadsScreen } from './screens'
import logo from './assets/logo.png'
import type { BootStatus, DraftSeed, HostStatus, ScreenName, ThreadRow } from './types'

const nav: { id: ScreenName; label: string; icon: string }[] = [
  { id: 'briefing', label: 'إحاطة', icon: 'grid' },
  { id: 'mail', label: 'البريد', icon: 'mail' },
  { id: 'threads', label: 'الخيوط', icon: 'box' },
  { id: 'draft', label: 'مسودة', icon: 'pen' },
  { id: 'pending', label: 'معلّق', icon: 'clock' },
  { id: 'settings', label: 'إعدادات', icon: 'gear' }
]

function Icon({ name }: { name: string }) {
  const common = { width: 18, height: 18, viewBox: '0 0 24 24', fill: 'none', stroke: 'currentColor', strokeWidth: 1.8 }
  if (name === 'mail') return <svg {...common}><rect x="3" y="5" width="18" height="14" rx="2" /><path d="m4 7 8 6 8-6" /></svg>
  if (name === 'box') return <svg {...common}><path d="M3 7.5 12 3l9 4.5-9 4.5L3 7.5Z" /><path d="M3 7.5V16.5L12 21l9-4.5V7.5" /><path d="M12 12v9" /></svg>
  if (name === 'pen') return <svg {...common}><path d="M4 20h4l10-10-4-4L4 16v4Z" /><path d="m13 7 4 4" /></svg>
  if (name === 'clock') return <svg {...common}><circle cx="12" cy="12" r="8" /><path d="M12 8v5l3 2" /></svg>
  if (name === 'gear') return <svg {...common}><circle cx="12" cy="12" r="3" /><path d="M12 3v2m0 14v2M3 12h2m14 0h2M5.6 5.6l1.4 1.4m10 10 1.4 1.4m0-12.8-1.4 1.4m-10 10-1.4 1.4" /></svg>
  return <svg {...common}><rect x="4" y="4" width="6" height="6" rx="1.5" /><rect x="14" y="4" width="6" height="6" rx="1.5" /><rect x="4" y="14" width="6" height="6" rx="1.5" /><rect x="14" y="14" width="6" height="6" rx="1.5" /></svg>
}

const titles: Record<ScreenName, string> = {
  briefing: 'إحاطة',
  mail: 'البريد',
  threads: 'الخيوط',
  draft: 'مسودة',
  pending: 'معلّق',
  settings: 'إعدادات'
}

export function App() {
  const [screen, setScreen] = useState<ScreenName>('briefing')
  const [boot, setBoot] = useState<BootStatus | null>(null)
  const [host, setHost] = useState<HostStatus | null>(null)
  const [seed, setSeed] = useState<DraftSeed | null>(null)

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

  const account = 'o.bamusa@ejadah-sa.com'

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
          {screen === 'threads' && <ThreadsScreen onDraft={draftFrom} />}
          {screen === 'draft' && <DraftScreen seed={seed} />}
          {screen === 'pending' && <PendingScreen />}
          {screen === 'settings' && <SettingsScreen />}
        </main>
      </div>
    </div>
  )
}
