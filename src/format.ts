const stateLabel: Record<string, string> = {
  awaiting_me: 'يحتاجك',
  awaiting_them: 'بانتظارهم',
  info: 'للعلم',
  closed: 'مغلق'
}

const draftLabel: Record<string, string> = {
  editing: 'مسودة',
  send_requested: 'بانتظار الإرسال',
  sent: 'أُرسل',
  save_requested: 'جارٍ الحفظ',
  saved: 'حُفظت في Outlook'
}

export function stateName(state: string) {
  return stateLabel[state] ?? state
}

export function draftName(status: string) {
  return draftLabel[status] ?? status
}

export function when(value: string) {
  if (!value) return ''
  return new Intl.DateTimeFormat('ar-SA', {
    timeZone: 'Asia/Riyadh',
    month: 'short',
    day: 'numeric',
    hour: '2-digit',
    minute: '2-digit'
  }).format(new Date(value))
}

export function openButton(kind: string) {
  if (kind === 'teams') return 'افتح Teams'
  if (kind === 'briefing') return 'افتح الموجز'
  return 'افتح البريد'
}
