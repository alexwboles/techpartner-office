// @vitest-environment jsdom
import { afterEach, beforeEach, describe, expect, it, vi } from 'vitest'
import { act, createElement } from 'react'
import { createRoot, type Root } from 'react-dom/client'
import { defaultAiSettings } from '@genoffice/ai-provider'
import { Home } from '../src/renderer/src/Home'
import { LocaleProvider } from '../src/renderer/src/locale'
import type { RecentEntry } from '../src/shared/home-api'

let root: Root
let container: HTMLDivElement
const files: RecentEntry[] = [
  {
    path: '/fixtures/brief.docx',
    name: 'brief.docx',
    ext: 'docx',
    mtimeMs: 1,
    sizeBytes: 10,
    starred: true,
  },
  {
    path: '/fixtures/budget.xlsx',
    name: 'budget.xlsx',
    ext: 'xlsx',
    mtimeMs: 2,
    sizeBytes: 20,
    starred: false,
  },
]
const page = (starred: boolean) =>
  vi.fn(async (query: { ext?: string; limit?: number }) => {
    const entries = files.filter(
      (file) => (!starred || file.starred) && (!query.ext || file.ext === query.ext),
    )
    return {
      entries: query.limit === 0 ? [] : entries,
      total: entries.length,
      totalAll: entries.length,
      offset: 0,
    }
  })

beforeEach(async () => {
  HTMLElement.prototype.scrollTo = vi.fn()
  vi.stubGlobal('IS_REACT_ACT_ENVIRONMENT', true)
  vi.stubGlobal(
    'fetch',
    vi.fn(() => {
      throw new Error('No network in navigation tests')
    }),
  )
  window.aiOffice = {
    folderRoots: async () => [],
    onFolderChanged: () => () => {},
    recents: page(false),
    starred: page(true),
    starredGroups: async () => [],
    accountStatus: async () => ({ loggedIn: false }),
    onOpenSettings: () => () => {},
    onAccountLogin: () => () => {},
    kbList: async () => [],
    createFromTemplate: vi.fn(async () => ({ ok: true })),
    aiGetSettings: async () => defaultAiSettings(),
    aiGskStatus: async () => ({ loggedIn: false }),
    buildChatState: async () => ({ modelAvailable: false }),
    buildChatSend: vi.fn(async () => ({
      kind: 'questions' as const,
      message: '',
      questions: [],
      planLines: [],
      files: [],
      errors: [],
      modelAvailable: false,
    })),
    buildChatReset: async () => {},
    onAiSettingsChanged: () => () => {},
  } as unknown as Window['aiOffice']
  window.localStorage.clear()
  container = document.createElement('div')
  document.body.append(container)
  root = createRoot(container)
  await act(async () =>
    root.render(createElement(LocaleProvider, { initial: 'en', children: createElement(Home) })),
  )
})
afterEach(async () => {
  await act(async () => root.unmount())
  container.remove()
  vi.unstubAllGlobals()
})
function nav(label: string): HTMLButtonElement {
  return [...container.querySelectorAll<HTMLButtonElement>('.icon-rail .rail-btn')].find(
    (button) => button.getAttribute('aria-label') === label,
  )!
}
async function click(button: HTMLButtonElement) {
  expect(button).toBeTruthy()
  await act(async () => button.click())
}
async function openGallery() {
  await click(container.querySelector<HTMLButtonElement>('.landing-template-btn')!)
  expect(container.querySelector('.template-gallery')).not.toBeNull()
}
function gallerySection(label: string): Element | null {
  return container.querySelector(`.template-gallery-section[aria-label="${label}"]`)
}
function galleryCard(scope: Element, title: string): Element {
  return [...scope.querySelectorAll('.template-card')].find(
    (card) => card.querySelector('.template-card-title')?.textContent === title,
  )!
}
function mainGalleryCard(title: string): Element {
  return galleryCard(container.querySelector('.template-gallery-main-grid')!, title)
}
async function setGallerySearch(value: string) {
  const input = container.querySelector<HTMLInputElement>('.template-gallery-search')!
  const setter = Object.getOwnPropertyDescriptor(window.HTMLInputElement.prototype, 'value')!.set!
  await act(async () => {
    setter.call(input, value)
    input.dispatchEvent(new window.Event('input', { bubbles: true }))
  })
}
function storedGalleryState(): { favorites: string[]; recent: Array<{ id: string }> } {
  return (
    JSON.parse(window.localStorage.getItem('templateGallery.state') ?? 'null') ?? {
      favorites: [],
      recent: [],
    }
  )
}
function selected(): string[] {
  return [...container.querySelectorAll('.icon-rail .rail-btn.active')].map((button) =>
    button.getAttribute('aria-label')!,
  )
}

describe('Home rail navigation', () => {
  it.each(['Recent', 'Starred'])('shows only Ask AI as selected after leaving %s', async (from) => {
    await click(nav(from))
    expect(selected()).toEqual([from])
    await click(nav('Ask AI'))
    expect(container.querySelector('.kb-ask')).not.toBeNull()
    expect(selected()).toEqual(['Ask AI'])
    expect(nav('Ask AI').getAttribute('aria-current')).toBe('page')
    expect(nav(from).hasAttribute('aria-current')).toBe(false)
    expect(fetch).not.toHaveBeenCalled()
    await click(nav(from))
    expect(selected()).toEqual([from])
    expect(nav(from).getAttribute('aria-current')).toBe('page')
    expect(container.querySelector('.kb-ask')).toBeNull()
    expect(container.querySelector('.content')?.textContent).toContain('brief.docx')
    if (from === 'Starred')
      expect(container.querySelector('.content')?.textContent).not.toContain('budget.xlsx')
  })

  it('keeps the file-type filter when returning from Ask AI', async () => {
    // the landing has no filter pills: enter the full Recent view first
    await click(nav('Recent'))
    const docs = [...container.querySelectorAll<HTMLButtonElement>('.filter-pill')].find(
      (button) => button.textContent === 'Docs',
    )!
    await click(docs)
    expect(container.querySelector('.content')?.textContent).toContain('brief.docx')
    expect(container.querySelector('.content')?.textContent).not.toContain('budget.xlsx')
    await click(nav('Ask AI'))
    expect(selected()).toEqual(['Ask AI'])
    await click(nav('Recent'))
    expect(container.querySelector('.filter-pill.active')?.textContent).toBe('Docs')
    expect(container.querySelector('.content')?.textContent).toContain('brief.docx')
    expect(container.querySelector('.content')?.textContent).not.toContain('budget.xlsx')
    expect(window.aiOffice.recents).toHaveBeenLastCalledWith(
      expect.objectContaining({ ext: 'docx' }),
    )
  })

  it('toggles the folder drawer from the rail', async () => {
    expect(container.querySelector('.folder-drawer')).toBeNull()
    await click(nav('Folders'))
    expect(container.querySelector('.folder-drawer .folder-panel')).not.toBeNull()
    expect(nav('Folders').getAttribute('aria-expanded')).toBe('true')
    await click(nav('Folders'))
    expect(container.querySelector('.folder-drawer')).toBeNull()
    expect(nav('Folders').getAttribute('aria-expanded')).toBe('false')
  })

  it('opens the template gallery from the landing and creates from a template', async () => {
    const templateBtn = container.querySelector<HTMLButtonElement>('.landing-template-btn')!
    await click(templateBtn)
    expect(container.querySelector('.template-gallery')).not.toBeNull()
    expect(container.querySelector('.template-gallery')?.textContent).toContain('Modern Resume')
    const useButton = [...container.querySelectorAll<HTMLButtonElement>('.template-card-use')].find(
      (button) =>
        button.closest('.template-card')?.querySelector('.template-card-title')?.textContent ===
        'Modern Resume',
    )!
    await click(useButton)
    expect(window.aiOffice.createFromTemplate).toHaveBeenCalledWith(
      expect.objectContaining({ templateId: 'resume-modern' }),
    )
    expect(container.querySelector('.template-gallery')).toBeNull()
  })

  it('pins starred templates in a Favorites section only in the unfiltered view', async () => {
    await openGallery()
    // empty state: no sections at all, not even headers
    expect(gallerySection('Favorite templates')).toBeNull()
    expect(gallerySection('Recently used')).toBeNull()

    const star =
      mainGalleryCard('Modern Resume').querySelector<HTMLButtonElement>('.template-card-star')!
    expect(star.getAttribute('aria-label')).toBe('Favorite this template')
    await click(star)
    expect(
      mainGalleryCard('Modern Resume')
        .querySelector('.template-card-star')!
        .getAttribute('aria-label'),
    ).toBe('Unfavorite this template')

    // the favorite is pinned above the main grid, and stays in the main grid too
    const favorites = gallerySection('Favorite templates')!
    expect(favorites).not.toBeNull()
    expect(galleryCard(favorites, 'Modern Resume')).toBeTruthy()
    expect(mainGalleryCard('Modern Resume')).toBeTruthy()
    expect(storedGalleryState().favorites).toEqual(['resume-modern'])

    // picking a category chip returns to the single flat grid
    const chip = [...container.querySelectorAll<HTMLButtonElement>('.template-chip')].find(
      (button) => button.textContent !== 'All',
    )!
    await click(chip)
    expect(gallerySection('Favorite templates')).toBeNull()
    const allChip = [...container.querySelectorAll<HTMLButtonElement>('.template-chip')].find(
      (button) => button.textContent === 'All',
    )!
    await click(allChip)
    expect(gallerySection('Favorite templates')).not.toBeNull()

    // typing a search query also hides the sections; the flat grid still matches
    await setGallerySearch('resume')
    expect(gallerySection('Favorite templates')).toBeNull()
    expect(mainGalleryCard('Modern Resume')).toBeTruthy()
    await setGallerySearch('')

    // unstarring removes the section entirely again
    await click(
      mainGalleryCard('Modern Resume').querySelector<HTMLButtonElement>('.template-card-star')!,
    )
    expect(gallerySection('Favorite templates')).toBeNull()
    expect(storedGalleryState().favorites).toEqual([])
  })

  it('adds a successfully used template to Recently used, and only on success', async () => {
    await openGallery()
    const useButton =
      mainGalleryCard('Modern Resume').querySelector<HTMLButtonElement>('.template-card-use')!
    await click(useButton)
    expect(window.aiOffice.createFromTemplate).toHaveBeenCalledWith(
      expect.objectContaining({ templateId: 'resume-modern' }),
    )
    expect(container.querySelector('.template-gallery')).toBeNull()
    expect(storedGalleryState().recent.map((entry) => entry.id)).toEqual(['resume-modern'])

    await openGallery()
    const recent = gallerySection('Recently used')!
    expect(recent).not.toBeNull()
    expect(galleryCard(recent, 'Modern Resume')).toBeTruthy()
    // the shortcut does not remove the template from the main grid
    expect(mainGalleryCard('Modern Resume')).toBeTruthy()
    expect(gallerySection('Favorite templates')).toBeNull()
  })

  it('does not record a template whose creation failed', async () => {
    const create = window.aiOffice.createFromTemplate as unknown as ReturnType<typeof vi.fn>
    create.mockResolvedValueOnce({ ok: false, error: 'boom' })
    await openGallery()
    const useButton =
      mainGalleryCard('Modern Resume').querySelector<HTMLButtonElement>('.template-card-use')!
    await click(useButton)
    // the gallery stays open with the error, and nothing was recorded
    expect(container.querySelector('.template-gallery')).not.toBeNull()
    expect(container.querySelector('.template-gallery')?.textContent).toContain('boom')
    expect(gallerySection('Recently used')).toBeNull()
    expect(storedGalleryState().recent).toEqual([])
  })

  it('opens on the search-first landing and returns to it from Home', async () => {
    expect(container.querySelector('.landing-headline')).not.toBeNull()
    expect(container.querySelector('.landing-search-hint')?.textContent).toBe('Ctrl K')
    // the landing carries the Build chat card where the Project AI strip used to be
    expect(container.querySelector('.landing-ai-strip')).toBeNull()
    expect(container.querySelector('.build-chat-embedded .build-chat-input')).not.toBeNull()
    expect(container.querySelector('.build-chat-embedded .build-chat-composer')).not.toBeNull()
    expect(container.querySelector('.landing-recent-list')?.textContent).toContain('brief.docx')
    expect(selected()).toEqual(['Home'])
    const seeAll = container.querySelector<HTMLButtonElement>('.landing-see-all')!
    await click(seeAll)
    expect(container.querySelector('.landing-headline')).toBeNull()
    expect(container.querySelector('.recent-table')).not.toBeNull()
    await click(nav('Home'))
    expect(container.querySelector('.landing-headline')).not.toBeNull()
    expect(selected()).toEqual(['Home'])
  })
})
