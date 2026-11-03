import { useMemo, useState } from 'react'
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
  type TemplateCategory,
  type TemplateDef,
} from '../../shared/template-defs'
import './template-gallery.css'

/**
 * Template gallery modal: category chips + search over the template
 * library, one card per template with a structure preview derived from the
 * real template content (headings, sheet names, slide titles — never a
 * fake screenshot). "Use template" builds a real document and opens it;
 * "Open local file" keeps the old file-picker behavior reachable.
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

  const useTemplate = async (template: TemplateDef) => {
    setBusyId(template.id)
    setError(null)
    try {
      const result = await window.aiOffice.createFromTemplate({
        templateId: template.id,
        opts: createOpts,
      })
      if (result.ok) {
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
        <div className="template-gallery-grid">
          {visible.map((template) => {
            const preview = templatePreview(template)
            return (
              <article key={template.id} className="template-card">
                <div className="template-card-head">
                  <FileBadge ext={templateExtOf(template)} size={30} />
                  <div className="template-card-titles">
                    <div className="template-card-title">{template.title}</div>
                    <div className="template-card-app">
                      {t(templateCategoryKey(template.category) as StringKey)}
                    </div>
                  </div>
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
                  onClick={() => void useTemplate(template)}
                >
                  {busyId === template.id ? t('templateGalleryCreating') : t('templateGalleryUse')}
                </button>
              </article>
            )
          })}
          {visible.length === 0 && (
            <p className="template-gallery-empty">{t('templateGalleryEmpty')}</p>
          )}
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
