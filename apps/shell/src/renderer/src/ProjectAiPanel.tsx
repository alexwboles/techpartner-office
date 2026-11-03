import { useCallback, useEffect, useRef, useState } from 'react'
import {
  PROJECT_AI_PROJECT_TYPES,
  normalizeProjectType,
  type ProjectAiProjectType,
  type ProjectAiState,
} from '../../shared/project-ai'
import { useI18n, type StringKey } from './locale'
import './project-ai.css'

const EMPTY_STATE: ProjectAiState = {
  folder: null,
  files: [],
  projectType: 'business-plan',
  status: 'idle',
  statusMessage: '',
  applied: [],
  queue: [],
  lastReviewAt: null,
  lastError: null,
}

/** string key naming each project type, for the type picker */
const TYPE_LABEL_KEYS: Record<ProjectAiProjectType, StringKey> = {
  'business-plan': 'projectAiTypeBusinessPlan',
  'grant-application': 'projectAiTypeGrantApplication',
  'event-plan': 'projectAiTypeEventPlan',
  'marketing-plan': 'projectAiTypeMarketingPlan',
}

/**
 * Project AI panel, rendered in place of the Home file list when the
 * sidebar's Project AI entry is active.
 *
 * Flow: pick a saved project folder, pick a project type (business plan,
 * grant application, event plan or marketing plan), type a short brief,
 * build. The service generates a document, a budget spreadsheet and a
 * deck into the folder and opens them as tabs, then watches the folder
 * and reviews the files as they change. The panel polls the service
 * state (the watcher itself is main-side and debounced).
 *
 * Safety, stated in the panel so it is never folklore: only typo fixes
 * and number syncs (spreadsheet = source of truth) are applied
 * automatically, and every one lands in the change log below. Everything
 * else waits in the approval queue until the user clicks Accept.
 */
export function ProjectAiPanel() {
  const { t } = useI18n()
  const [state, setState] = useState<ProjectAiState>(EMPTY_STATE)
  const [folder, setFolder] = useState('')
  const [brief, setBrief] = useState('')
  const [projectType, setProjectType] = useState<ProjectAiProjectType>('business-plan')
  /** true once the user picks a type themselves: the polled service state
   *  (which defaults to business plan until a build) must not overwrite a
   *  pick the user is about to build with */
  const typePickedRef = useRef(false)
  const [busy, setBusy] = useState(false)
  const [notice, setNotice] = useState('')

  const refresh = useCallback(async () => {
    try {
      const next = await window.aiOffice.projectAiState()
      setState(next)
      if (next.folder) setFolder(next.folder)
      // the type persists in the service state like the rest of the panel
      // state; state saved before project types existed has no type and
      // reads as a business plan. A type the user just picked is theirs
      // until they build with it, so polling does not overwrite it.
      if (!typePickedRef.current) setProjectType(normalizeProjectType(next.projectType))
    } catch {
      // the panel keeps its last state when the service is unreachable
    }
  }, [])

  useEffect(() => {
    void refresh()
    const timer = setInterval(() => {
      void refresh()
    }, 4000)
    return () => clearInterval(timer)
  }, [refresh])

  const pickFolder = async (): Promise<void> => {
    try {
      const picked = await window.aiOffice.projectAiPickFolder()
      if (picked) {
        setFolder(picked)
        setNotice('')
      }
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error))
    }
  }

  const build = async (): Promise<void> => {
    if (!folder || !brief.trim()) return
    setBusy(true)
    setNotice('')
    try {
      await window.aiOffice.projectAiBuild(folder, brief, projectType)
      // the built type is now the service state's type, so polling may
      // adopt it again from here on
      typePickedRef.current = false
      await refresh()
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error))
    } finally {
      setBusy(false)
    }
  }

  const reviewNow = async (): Promise<void> => {
    setBusy(true)
    try {
      await window.aiOffice.projectAiReview()
      await refresh()
    } catch (error) {
      setNotice(error instanceof Error ? error.message : String(error))
    } finally {
      setBusy(false)
    }
  }

  const accept = async (id: string): Promise<void> => {
    const ok = await window.aiOffice.projectAiAccept(id)
    if (!ok) setNotice('That suggestion could not be applied exactly; it stays in the queue.')
    await refresh()
  }

  const reject = async (id: string): Promise<void> => {
    await window.aiOffice.projectAiReject(id)
    await refresh()
  }

  const openFile = (path: string): void => {
    void window.aiOffice.openPath(path)
  }

  return (
    <main className="content project-ai">
      <section className="project-ai-card" aria-label="Project AI">
        <h1 className="project-ai-title">Project AI</h1>
        <p className="project-ai-sub">{t('projectAiSub')}</p>

        <div className="project-ai-row">
          <label className="project-ai-label" htmlFor="project-ai-folder">
            Project folder
          </label>
          <div className="project-ai-folder-row">
            <input
              id="project-ai-folder"
              className="project-ai-input"
              value={folder}
              placeholder="No folder picked yet"
              readOnly
            />
            <button className="project-ai-btn" onClick={() => void pickFolder()}>
              Choose folder
            </button>
          </div>
        </div>

        <div className="project-ai-row">
          <span className="project-ai-label" id="project-ai-type-label">
            {t('projectAiTypeLabel')}
          </span>
          <div className="project-ai-types" role="group" aria-labelledby="project-ai-type-label">
            {PROJECT_AI_PROJECT_TYPES.map((type) => (
              <button
                key={type}
                type="button"
                className={`project-ai-type${projectType === type ? ' project-ai-type-active' : ''}`}
                aria-pressed={projectType === type}
                onClick={() => {
                  typePickedRef.current = true
                  setProjectType(type)
                }}
              >
                {t(TYPE_LABEL_KEYS[type])}
              </button>
            ))}
          </div>
        </div>

        <div className="project-ai-row">
          <label className="project-ai-label" htmlFor="project-ai-brief">
            Brief
          </label>
          <textarea
            id="project-ai-brief"
            className="project-ai-textarea"
            value={brief}
            onChange={(event) => setBrief(event.target.value)}
            placeholder="Example: A bakery in Palatka selling fresh bread and pastries to local customers."
            rows={3}
          />
        </div>

        <div className="project-ai-actions">
          <button
            className="project-ai-btn project-ai-btn-primary"
            disabled={!folder || !brief.trim() || busy || state.status === 'building'}
            onClick={() => void build()}
          >
            {state.status === 'building'
              ? 'Building…'
              : `${t('projectAiBuild')}: ${t(TYPE_LABEL_KEYS[projectType])}`}
          </button>
          <button
            className="project-ai-btn"
            disabled={busy || !state.folder}
            onClick={() => void reviewNow()}
          >
            Review now
          </button>
        </div>

        {state.statusMessage && <p className="project-ai-status">{state.statusMessage}</p>}
        {state.lastError && <p className="project-ai-error">{state.lastError}</p>}
        {notice && <p className="project-ai-error">{notice}</p>}

        {state.files.length > 0 && (
          <div className="project-ai-files">
            <h2 className="project-ai-heading">Project files</h2>
            <ul>
              {state.files.map((file) => (
                <li key={file}>
                  <button className="project-ai-link" onClick={() => openFile(file)}>
                    {file.split(/[/\\]/).pop()}
                  </button>
                </li>
              ))}
            </ul>
          </div>
        )}
      </section>

      <section className="project-ai-card" aria-label="Applied automatically">
        <h2 className="project-ai-heading">Applied automatically</h2>
        <p className="project-ai-note">
          Only two things are ever applied without asking: typo fixes, and number syncs where a
          figure in the plan or deck disagrees with the budget spreadsheet (the spreadsheet is the
          source of truth). If a file is open in a tab, reopen it to see an applied change.
        </p>
        {state.applied.length === 0 ? (
          <p className="project-ai-empty">Nothing applied automatically yet.</p>
        ) : (
          <ul className="project-ai-list">
            {state.applied.map((entry, index) => (
              <li key={`${entry.suggestionId}-${index}`}>
                <span className="project-ai-kind">{entry.kind}</span> {entry.fileName}:{' '}
                {entry.before} → {entry.after}
                <span className="project-ai-via"> ({entry.via})</span>
              </li>
            ))}
          </ul>
        )}
      </section>

      <section className="project-ai-card" aria-label="Needs your approval">
        <h2 className="project-ai-heading">Needs your approval</h2>
        <p className="project-ai-note">
          Everything else — wording, structure, new content, and any figure where the spreadsheet is
          not clearly the source — waits here until you accept it.
        </p>
        {state.queue.length === 0 ? (
          <p className="project-ai-empty">
            No suggestions waiting.
            {state.lastReviewAt
              ? ` Last review: ${new Date(state.lastReviewAt).toLocaleString()}.`
              : ' Reviews run automatically a moment after you save a project file.'}
          </p>
        ) : (
          <ul className="project-ai-list">
            {state.queue.map((suggestion) => (
              <li key={suggestion.id} className="project-ai-queue-item">
                <div>
                  <strong>{suggestion.title}</strong>
                  <div className="project-ai-detail">{suggestion.detail}</div>
                  {suggestion.before && suggestion.before !== suggestion.after && (
                    <div className="project-ai-diff">
                      {suggestion.before} → {suggestion.after}
                    </div>
                  )}
                </div>
                <div className="project-ai-queue-actions">
                  <button
                    className="project-ai-btn project-ai-btn-primary"
                    onClick={() => void accept(suggestion.id)}
                  >
                    Accept
                  </button>
                  <button className="project-ai-btn" onClick={() => void reject(suggestion.id)}>
                    Dismiss
                  </button>
                </div>
              </li>
            ))}
          </ul>
        )}
      </section>
    </main>
  )
}
