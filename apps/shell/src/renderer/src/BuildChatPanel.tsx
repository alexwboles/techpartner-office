import { useEffect, useRef, useState, type ReactNode } from 'react'
import type { BuildChatTranscriptEntry, BuildChatTurn } from '../../shared/build-chat'
import { useI18n } from './locale'
import './build-chat.css'

type ChatEntry = BuildChatTranscriptEntry

const APP_BADGE: Record<string, string> = {
  docs: 'DOC',
  sheets: 'SHEET',
  slides: 'DECK',
}

/** One-tap starting points on a fresh conversation: each chip sends its
 *  prompt as the first user message, so the clarify-then-build flow
 *  starts without typing. The prompts stay inside what the chat can
 *  actually do — a template-backed file in each of the three apps. */
const STARTER_KEYS = [
  'buildChatStarterInvoice',
  'buildChatStarterBusinessPlan',
  'buildChatStarterBudget',
  'buildChatStarterPresentation',
] as const

/**
 * Build chat panel. Two placements share this one component and the
 * same main-process conversation: the rail's Build entry renders it
 * full-height in place of the Home file list, and the Home landing
 * embeds it as a card (`embedded`) under the create row.
 *
 * The user describes what they need in plain words; the main-process
 * service holds the conversation and answers one turn at a time: a
 * questions card when facts are missing (nothing builds until they are
 * answered), a plan card before files are created, an explicit
 * confirmation wait before an existing file is edited, and created-file
 * cards with Open buttons once a build lands. When no AI is configured
 * the header says so plainly and points at the AI settings — in that
 * mode the service only matches templates and fills their details.
 */
export function BuildChatPanel(props: {
  onOpenPath: (path: string) => void
  onOpenModelSettings: () => void
  /** landing card placement: fixed height, header chrome shed (see build-chat.css) */
  embedded?: boolean
}) {
  const { t } = useI18n()
  const [entries, setEntries] = useState<ChatEntry[]>([])
  const [draft, setDraft] = useState('')
  const [busy, setBusy] = useState(false)
  const [hydrated, setHydrated] = useState(false)
  const [modelAvailable, setModelAvailable] = useState<boolean | null>(null)
  const listRef = useRef<HTMLDivElement | null>(null)

  // The service owns the transcript (persisted across restarts): hydrate
  // from it on mount, so leaving and returning — or relaunching the app —
  // shows the same conversation instead of a blank panel.
  useEffect(() => {
    let live = true
    void window.aiOffice
      .buildChatState()
      .then((state) => {
        if (!live) return
        setModelAvailable(state.modelAvailable)
        if (Array.isArray(state.entries) && state.entries.length > 0) {
          setEntries(state.entries)
        }
        setHydrated(true)
      })
      .catch(() => {
        if (!live) return
        setModelAvailable(null)
        setHydrated(true)
      })
    return () => {
      live = false
    }
  }, [])

  useEffect(() => {
    const list = listRef.current
    if (list) list.scrollTop = list.scrollHeight
  }, [entries, busy])

  const sendText = async (rawText: string): Promise<void> => {
    const text = rawText.trim()
    if (!text || busy) return
    setBusy(true)
    setEntries((current) => [...current, { role: 'user', text }])
    try {
      const result = await window.aiOffice.buildChatSend(text)
      setModelAvailable(result.modelAvailable)
      setEntries((current) => [
        ...current,
        { role: 'assistant', text: result.message, turn: result },
      ])
    } catch (error) {
      const message = error instanceof Error ? error.message : String(error)
      setEntries((current) => [
        ...current,
        {
          role: 'assistant',
          text: '',
          turn: {
            kind: 'error',
            message: '',
            questions: [],
            planLines: [],
            files: [],
            errors: [message],
            modelAvailable: modelAvailable ?? true,
          },
        },
      ])
    } finally {
      setBusy(false)
    }
  }

  const send = async (): Promise<void> => {
    const text = draft.trim()
    if (!text) return
    setDraft('')
    await sendText(text)
  }

  const newChat = async (): Promise<void> => {
    if (busy) return
    try {
      await window.aiOffice.buildChatReset()
    } catch {
      // a reset failure leaves the local transcript cleared anyway
    }
    setEntries([])
    setDraft('')
  }

  const renderTurn = (entry: ChatEntry): ReactNode => {
    const result = entry.turn
    if (!result) return entry.text ? <p className="build-chat-text">{entry.text}</p> : null
    return (
      <>
        {result.message && <p className="build-chat-text">{result.message}</p>}
        {result.questions.length > 0 && (
          <div className="build-chat-card">
            <h3 className="build-chat-card-title">{t('buildChatQuestionsTitle')}</h3>
            <ol className="build-chat-questions">
              {result.questions.map((question, index) => (
                <li key={index}>{question}</li>
              ))}
            </ol>
          </div>
        )}
        {result.planLines.length > 0 && (
          <div className="build-chat-card">
            <h3 className="build-chat-card-title">{t('buildChatPlanTitle')}</h3>
            <ul className="build-chat-plan">
              {result.planLines.map((line, index) => (
                <li key={index}>{line}</li>
              ))}
            </ul>
            {result.kind === 'confirm' && (
              <p className="build-chat-confirm-hint">{t('buildChatConfirmHint')}</p>
            )}
          </div>
        )}
        {result.files.length > 0 && (
          <div className="build-chat-card">
            <h3 className="build-chat-card-title">{t('buildChatCreatedTitle')}</h3>
            <ul className="build-chat-files">
              {result.files.map((file, index) => (
                <li key={`${file.title}-${index}`} className="build-chat-file">
                  <span className="build-chat-badge">{APP_BADGE[file.app] ?? 'FILE'}</span>
                  <span className="build-chat-file-title">{file.title}</span>
                  {file.path ? (
                    <button className="build-chat-btn" onClick={() => props.onOpenPath(file.path!)}>
                      {t('buildChatOpenFile')}
                    </button>
                  ) : (
                    <span className="build-chat-file-note">{t('buildChatOpenedTab')}</span>
                  )}
                </li>
              ))}
            </ul>
          </div>
        )}
        {result.errors.length > 0 && (
          <div className="build-chat-card build-chat-card-error">
            <h3 className="build-chat-card-title">{t('buildChatFailedTitle')}</h3>
            <ul className="build-chat-plan">
              {result.errors.map((error, index) => (
                <li key={index}>{error}</li>
              ))}
            </ul>
          </div>
        )}
      </>
    )
  }

  const content = (
    <>
      <section className="build-chat-head" aria-label={t('buildChatTitle')}>
        <div className="build-chat-head-row">
          <h1 className="build-chat-title">{t('buildChatTitle')}</h1>
          <button className="build-chat-btn" onClick={() => void newChat()} disabled={busy}>
            {t('buildChatNew')}
          </button>
        </div>
        <p className="build-chat-sub">{t('buildChatSub')}</p>
        {modelAvailable === false && (
          <div className="build-chat-nomodel">
            <strong>{t('buildChatNoModelTitle')}</strong>
            <p>{t('buildChatNoModelBody')}</p>
            <button className="build-chat-btn" onClick={props.onOpenModelSettings}>
              {t('buildChatOpenSettings')}
            </button>
          </div>
        )}
      </section>

      <div className="build-chat-list" ref={listRef}>
        {entries.length === 0 && <p className="build-chat-empty">{t('buildChatEmpty')}</p>}
        {entries.map((entry, index) => (
          <div key={index} className={`build-chat-msg build-chat-msg-${entry.role}`}>
            {entry.role === 'user' ? (
              <p className="build-chat-text">{entry.text}</p>
            ) : (
              renderTurn(entry)
            )}
          </div>
        ))}
        {busy && <p className="build-chat-working">{t('buildChatSending')}</p>}
      </div>

      {hydrated && entries.length === 0 && !busy && (
        <div className="build-chat-starters">
          {STARTER_KEYS.map((key) => (
            <button key={key} className="build-chat-chip" onClick={() => void sendText(t(key))}>
              {t(key)}
            </button>
          ))}
        </div>
      )}

      <div className="build-chat-composer">
        <textarea
          className="build-chat-input"
          value={draft}
          placeholder={t('buildChatPlaceholder')}
          rows={3}
          onChange={(event) => setDraft(event.target.value)}
          onKeyDown={(event) => {
            if (event.key === 'Enter' && !event.shiftKey) {
              event.preventDefault()
              void send()
            }
          }}
        />
        <button
          className="build-chat-btn build-chat-btn-primary"
          onClick={() => void send()}
          disabled={busy || !draft.trim()}
        >
          {busy ? t('buildChatSending') : t('buildChatSend')}
        </button>
      </div>
    </>
  )

  if (props.embedded) {
    return (
      <section className="build-chat build-chat-embedded" aria-label={t('buildChatTitle')}>
        {content}
      </section>
    )
  }
  return <main className="content build-chat">{content}</main>
}
