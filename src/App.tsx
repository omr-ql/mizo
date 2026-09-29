import { useEffect, useState } from 'react'
import { BriefingScreen, DraftScreen, MailScreen, PendingScreen, SettingsScreen, ThreadsScreen } from './screens'
import type { BootStatus, DraftSeed, HostStatus, ScreenName, ThreadRow } from './types'

const nav: { id: ScreenName; label: string }[] = [
  { id: 'briefing', label: 'إحاطة' },
  { id: 'mail', label: 'البريد' },
  { id: 'threads', label: 'الخيوط' },
  { id: 'draft', label: 'مسودة' },
  { id: 'pending', label: 'معلّق' },
  { id: 'settings', label: 'إعدادات' }
]

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

  return (
    <div className="app">
      <aside className="side">
        <div className="brand">
          <strong>mizo</strong>
          <span>o.bamusa@ejadah-sa.com</span>
        </div>
        {nav.map((item) => (
          <button key={item.id} className={screen === item.id ? 'nav active' : 'nav'} onClick={() => setScreen(item.id)}>
            {item.label}
          </button>
        ))}
        <div className="side-foot">
          <div><span className={boot?.api ? 'dot on' : 'dot'} />Docker</div>
          <div><span className={host?.ok ? 'dot on' : 'dot'} />Outlook</div>
          <div>{host?.message || 'بانتظار الاتصال'}</div>
        </div>
      </aside>
      <main className="main">
        <header className="top">
          <div>
            <h1>{titles[screen]}</h1>
            <p>حساب العمل فقط، والإرسال بعد «أرسل».</p>
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
  )
}
