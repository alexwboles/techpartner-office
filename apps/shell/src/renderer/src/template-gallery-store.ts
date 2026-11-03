/**
 * Renderer-local state for the template gallery: favorite template ids and
 * the most recently used templates.
 *
 * Persistence follows the shell's existing pattern for UI prefs: Home's
 * folder-tree layout (Home.tsx `readTreeState` / `writeTreeState`) lives in
 * `localStorage` under a dotted key, wrapped in try/catch so corrupt or
 * unavailable storage falls back to a default instead of breaking the UI.
 * Template ids are renderer-only state — they are not files, so the main
 * process recents/starred store behind `window.aiOffice` does not fit them,
 * and no IPC surface is needed. Every read here is defensive: missing,
 * corrupt, or wrongly-shaped data yields an empty state.
 */

export const TEMPLATE_GALLERY_STORAGE_KEY = 'templateGallery.state'

/** most-recently-used list length; older uses fall off the end */
export const TEMPLATE_RECENT_LIMIT = 8

export interface TemplateRecentEntry {
  id: string
  usedAt: number
}

export interface TemplateGalleryState {
  favorites: readonly string[]
  recent: readonly TemplateRecentEntry[]
}

export const EMPTY_TEMPLATE_GALLERY_STATE: TemplateGalleryState = { favorites: [], recent: [] }

/** the slice of the Web Storage API this store needs (testable without a DOM) */
export interface TemplateGalleryStorage {
  getItem(key: string): string | null
  setItem(key: string, value: string): void
}

function defaultStorage(): TemplateGalleryStorage | null {
  try {
    return globalThis.localStorage ?? null
  } catch {
    // storage access itself can throw (blocked cookies / private modes)
    return null
  }
}

function resolveStorage(storage?: TemplateGalleryStorage): TemplateGalleryStorage | null {
  return storage ?? defaultStorage()
}

function sanitizeState(raw: unknown): TemplateGalleryState {
  if (!raw || typeof raw !== 'object') return EMPTY_TEMPLATE_GALLERY_STATE
  const parsed = raw as { favorites?: unknown; recent?: unknown }
  const favorites: string[] = []
  if (Array.isArray(parsed.favorites)) {
    for (const id of parsed.favorites) {
      if (typeof id === 'string' && id.length > 0 && !favorites.includes(id)) favorites.push(id)
    }
  }
  const recent: TemplateRecentEntry[] = []
  if (Array.isArray(parsed.recent)) {
    for (const entry of parsed.recent) {
      if (!entry || typeof entry !== 'object') continue
      const { id, usedAt } = entry as { id?: unknown; usedAt?: unknown }
      if (typeof id !== 'string' || id.length === 0) continue
      if (recent.some((item) => item.id === id)) continue
      recent.push({ id, usedAt: typeof usedAt === 'number' && usedAt > 0 ? usedAt : 0 })
      if (recent.length >= TEMPLATE_RECENT_LIMIT) break
    }
  }
  return { favorites, recent }
}

/** read the persisted state; any problem at all falls back to empty */
export function loadTemplateGalleryState(storage?: TemplateGalleryStorage): TemplateGalleryState {
  const store = resolveStorage(storage)
  if (!store) return EMPTY_TEMPLATE_GALLERY_STATE
  try {
    const raw = store.getItem(TEMPLATE_GALLERY_STORAGE_KEY)
    if (!raw) return EMPTY_TEMPLATE_GALLERY_STATE
    return sanitizeState(JSON.parse(raw) as unknown)
  } catch {
    // corrupt or absent: no favorites, no recents
    return EMPTY_TEMPLATE_GALLERY_STATE
  }
}

/** persist the state; failures (quota, blocked storage) are non-fatal */
export function saveTemplateGalleryState(
  state: TemplateGalleryState,
  storage?: TemplateGalleryStorage,
): void {
  const store = resolveStorage(storage)
  if (!store) return
  try {
    store.setItem(
      TEMPLATE_GALLERY_STORAGE_KEY,
      JSON.stringify({ favorites: [...state.favorites], recent: [...state.recent] }),
    )
  } catch {
    // the gallery just forgets across sessions
  }
}

/** star a template, or unstar it when it is already starred */
export function toggleTemplateFavorite(
  state: TemplateGalleryState,
  id: string,
): TemplateGalleryState {
  return state.favorites.includes(id)
    ? { ...state, favorites: state.favorites.filter((favorite) => favorite !== id) }
    : { ...state, favorites: [...state.favorites, id] }
}

/**
 * Record a successful use: the template moves to the front of the recent
 * list (deduplicated) and the list is capped at TEMPLATE_RECENT_LIMIT.
 * Callers must only invoke this after a creation actually succeeded.
 */
export function recordTemplateUse(
  state: TemplateGalleryState,
  id: string,
  usedAt: number = Date.now(),
): TemplateGalleryState {
  const rest = state.recent.filter((entry) => entry.id !== id)
  return { ...state, recent: [{ id, usedAt }, ...rest].slice(0, TEMPLATE_RECENT_LIMIT) }
}
