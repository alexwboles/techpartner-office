import { memo, useMemo, useState } from 'react'
import { FileBadge } from './file-badge'
import { useI18n } from './locale'
import type { StringKey } from './locale'
import type { NewFileOpts } from '../../shared/home-api'
import {
  TEMPLATE_CATEGORIES,
  TEMPLATES,
  templateCategoryKey,
  templateExtOf,
  templatePreview,
  templateVisualPreview,
  type TemplateCategory,
  type TemplateDef,
} from '../../shared/template-defs'
import {
  loadTemplateGalleryState,
  recordTemplateUse,
  saveTemplateGalleryState,
  toggleTemplateFavorite,
  type TemplateGalleryState,
} from './template-gallery-store'
import './template-gallery.css'

/**
 * Miniature live preview of a template's real first screenful, rendered
 * from the template definition (doc blocks, sheet grid, title slide) as
 * plain React elements — no injected HTML, no images, nothing async.
 * Memoized per template: the definitions are static, so each preview is
 * computed once and reused across gallery re-renders.
 */
const TemplateVisual = memo(function TemplateVisual({ template }: { template: TemplateDef }) {
  const { t } = useI18n()
  const visual = useMemo(() => templateVisualPreview(template), [template])
  if (visual.app === 'docs') {
    return (
      <div className="template-visual template-visual-doc" aria-hidden="true">
        <div className="template-visual-page">
          {visual.blocks.map((block, i) => {
            switch (block.kind) {
              case 'heading':
                return (
                  <div key={i} className={`tv-h${block.level}`}>
                    {block.text}
                  </div>
                )
              case 'paragraph':
                return (
                  <p key={i} className="tv-p">
                    {block.text}
                  </p>
                )
              case 'list':
                return (
                  <ul key={i} className="tv-list">
                    {block.items.map((item, j) => (
                      <li key={j}>{item}</li>
                    ))}
                  </ul>
                )
              case 'table':
                return (
                  <table key={i} className="tv-table">
                    <tbody>
                      {block.rows.map((row, r) => (
                        <tr key={r}>
                          {row.map((cell, c) =>
                            cell.bold ? <th key={c}>{cell.text}</th> : <td key={c}>{cell.text}</td>,
                          )}
                        </tr>
                      ))}
                    </tbody>
                  </table>
                )
            }
          })}
        </div>
      </div>
    )
  }
  if (visual.app === 'sheets') {
    return (
      <div className="template-visual template-visual-sheet" aria-hidden="true">
        <div className="template-visual-sheet-tab">{visual.sheetName}</div>
        <div className="template-visual-grid-wrap">
          <table className="template-visual-grid">
            <tbody>
              {visual.rows.map((row, r) => (
                <tr key={r} className={visual.headerRow && r === 0 ? 'tv-header' : undefined}>
                  {row.map((cell, c) => (
                    <td
                      key={c}
                      className={
                        [cell.numeric ? 'tv-num' : '', cell.bold ? 'tv-bold' : '']
                          .filter(Boolean)
                          .join(' ') || undefined
                      }
                    >
                      {cell.text}
                    </td>
                  ))}
                </tr>
              ))}
            </tbody>
          </table>
        </div>
      </div>
    )
  }
  return (
    <div className="template-visual template-visual-slide" aria-hidden="true">
      <div className="template-visual-slide-box">
        <div className="tv-slide-title">{visual.title}</div>
        {visual.subtitle && <div className="tv-slide-sub">{visual.subtitle}</div>}
        {visual.lines.map((line, i) => (
          <div key={i} className="tv-slide-line">
            • {line}
          </div>
        ))}
      </div>
      {visual.extraSlides > 0 && (
        <span className="template-visual-more">
          {t('templatePreviewMoreSlides', { count: visual.extraSlides })}
        </span>
      )}
    </div>
  )
})

/**
 * Template gallery modal: category chips + search over the template
 * library, one card per template with a live visual preview and a
 * structure preview, both derived from the real template content
 * (headings, sheet names, slide titles — never a fake screenshot).
 * "Use template" builds a real document and opens it; "Open local file"
 * keeps the old file-picker behavior reachable.
 *
 * Cards can be starred; favorites and the most recently used templates
 * (persisted renderer-locally, see template-gallery-store.ts) are pinned
 * as shortcut sections above the main grid in the unfiltered view only.
 */
export function TemplateGallery({
  createOpts,
  onClose,
}: {
  createOpts?: NewFileOpts
  onClose: () => void
}) {
  const { t } = useI18n()
  const [query, setQuery] = useState('')
  const [category, setCategory] = useState<'all' | TemplateCategory>('all')
  const [busyId, setBusyId] = useState<string | null>(null)
  const [error, setError] = useState<string | null>(null)
  const [galleryState, setGalleryState] = useState<TemplateGalleryState>(() =>
    loadTemplateGalleryState(),
  )

  const favoriteIds = useMemo(() => new Set(galleryState.favorites), [galleryState.favorites])
  /** starred templates, in library (TEMPLATES) order */
  const favoriteTemplates = useMemo(
    () => TEMPLATES.filter((template) => favoriteIds.has(template.id)),
    [favoriteIds],
  )
  /** recently used templates, most recent first; unknown ids are skipped */
  const recentTemplates = useMemo(() => {
    const byId = new Map(TEMPLATES.map((template) => [template.id, template]))
    return galleryState.recent
      .map((entry) => byId.get(entry.id))
      .filter((template): template is TemplateDef => template !== undefined)
  }, [galleryState.recent])
  // the pinned sections are shortcuts over the unfiltered library only:
  // searching or picking a category returns to the single flat grid
  const showSections = category === 'all' && query.trim() === ''

  const toggleFavorite = (template: TemplateDef) => {
    const next = toggleTemplateFavorite(galleryState, template.id)
    setGalleryState(next)
    saveTemplateGalleryState(next)
  }

  const visible = useMemo(() => {
    const needle = query.trim().toLowerCase()
    return TEMPLATES.filter((template) => {
      if (category !== 'all' && template.category !== category) return false
      if (!needle) return true
      const haystack = [
        template.title,
        template.description,
        t(templateCategoryKey(template.category) as StringKey),
        ...templatePreview(template),
      ]
        .join('\n')
        .toLowerCase()
      return haystack.includes(needle)
    })
  }, [query, category, t])

  const createFromTemplate = async (template: TemplateDef) => {
    setBusyId(template.id)
    setError(null)
    try {
      const result = await window.aiOffice.createFromTemplate({
        templateId: template.id,
        opts: createOpts,
      })
      if (result.ok) {
        // only a successful creation counts as a use
        const next = recordTemplateUse(galleryState, template.id)
        setGalleryState(next)
        saveTemplateGalleryState(next)
        onClose()
        return
      }
      setError(t('templateGalleryFailed', { error: result.error ?? '' }))
    } catch (err) {
      setError(
        t('templateGalleryFailed', {
          error: err instanceof Error ? err.message : String(err),
        }),
      )
    } finally {
      setBusyId(null)
    }
  }

  const renderCard = (template: TemplateDef) => {
    const preview = templatePreview(template)
    const starred = favoriteIds.has(template.id)
    return (
      <article key={template.id} className="template-card">
        <TemplateVisual template={template} />
        <div className="template-card-head">
          <FileBadge ext={templateExtOf(template)} size={30} />
          <div className="template-card-titles">
            <div className="template-card-title">{template.title}</div>
            <div className="template-card-app">
              {t(templateCategoryKey(template.category) as StringKey)}
            </div>
          </div>
          <button
            className={`template-card-star${starred ? ' starred' : ''}`}
            aria-label={t(starred ? 'templateGalleryUnfavorite' : 'templateGalleryFavorite')}
            aria-pressed={starred}
            onClick={() => toggleFavorite(template)}
          >
            <svg width="15" height="15" viewBox="0 0 16 16" aria-hidden="true">
              <path
                d="M8 1.9l1.9 3.85 4.25.62-3.07 3 .72 4.23L8 11.6l-3.8 2 .72-4.23-3.07-3 4.25-.62z"
                fill={starred ? '#f5a623' : 'none'}
                stroke={starred ? '#f5a623' : 'currentColor'}
                strokeWidth="1.2"
                strokeLinejoin="round"
              />
            </svg>
          </button>
        </div>
        <p className="template-card-desc">{template.description}</p>
        {preview.length > 0 && (
          <ul className="template-card-preview">
            {preview.slice(0, 4).map((line, i) => (
              <li key={i}>{line}</li>
            ))}
          </ul>
        )}
        <button
          className="template-card-use"
          disabled={busyId !== null}
          onClick={() => void createFromTemplate(template)}
        >
          {busyId === template.id ? t('templateGalleryCreating') : t('templateGalleryUse')}
        </button>
      </article>
    )
  }

  return (
    <div className="modal-overlay" onClick={onClose}>
      <div
        className="template-gallery"
        role="dialog"
        aria-modal="true"
        aria-label={t('templateGalleryTitle')}
        onClick={(event) => event.stopPropagation()}
      >
        <div className="template-gallery-header">
          <h3>{t('templateGalleryTitle')}</h3>
          <button
            className="template-gallery-close"
            aria-label={t('templateGalleryClose')}
            onClick={onClose}
          >
            ×
          </button>
        </div>
        <div className="template-gallery-search-row">
          <input
            className="template-gallery-search"
            type="search"
            value={query}
            onChange={(event) => setQuery(event.target.value)}
            placeholder={t('templateGallerySearch')}
            aria-label={t('templateGallerySearch')}
            autoFocus
          />
        </div>
        <div
          className="template-gallery-chips"
          role="tablist"
          aria-label={t('templateGalleryTitle')}
        >
          <button
            className={category === 'all' ? 'template-chip active' : 'template-chip'}
            onClick={() => setCategory('all')}
          >
            {t('templateGalleryAll')}
          </button>
          {TEMPLATE_CATEGORIES.map((item) => (
            <button
              key={item}
              className={category === item ? 'template-chip active' : 'template-chip'}
              onClick={() => setCategory(item)}
            >
              {t(templateCategoryKey(item) as StringKey)}
            </button>
          ))}
        </div>
        <div className="template-gallery-scroll">
          {showSections && favoriteTemplates.length > 0 && (
            <section
              className="template-gallery-section"
              aria-label={t('templateGalleryFavorites')}
            >
              <h4 className="template-gallery-section-title">{t('templateGalleryFavorites')}</h4>
              <div className="template-gallery-grid">{favoriteTemplates.map(renderCard)}</div>
            </section>
          )}
          {showSections && recentTemplates.length > 0 && (
            <section className="template-gallery-section" aria-label={t('templateGalleryRecent')}>
              <h4 className="template-gallery-section-title">{t('templateGalleryRecent')}</h4>
              <div className="template-gallery-grid">{recentTemplates.map(renderCard)}</div>
            </section>
          )}
          <div className="template-gallery-grid template-gallery-main-grid">
            {visible.map(renderCard)}
            {visible.length === 0 && (
              <p className="template-gallery-empty">{t('templateGalleryEmpty')}</p>
            )}
          </div>
        </div>
        <div className="template-gallery-footer">
          {error ? <span className="template-gallery-error">{error}</span> : <span />}
          <button
            className="template-gallery-open-local"
            onClick={() => {
              onClose()
              void window.aiOffice.browse()
            }}
          >
            {t('templateGalleryOpenLocal')}
          </button>
        </div>
      </div>
    </div>
  )
}
