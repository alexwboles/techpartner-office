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
    onAiSettingsChanged: () => () => {},
  } as unknown as Window['aiOffice']
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

  it('opens on the search-first landing and returns to it from Home', async () => {
    expect(container.querySelector('.landing-headline')).not.toBeNull()
    expect(container.querySelector('.landing-search-hint')?.textContent).toBe('Ctrl K')
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
