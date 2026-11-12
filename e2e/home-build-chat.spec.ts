import { test, expect } from '@playwright/test'
import { existsSync, mkdtempSync, realpathSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo } from './helpers'

/**
 * Home Build chat in a fresh profile: no AI is configured, so the panel
 * is in template-only mode. The deterministic path must ask the matched
 * template's placeholder questions first (nothing is built yet), then
 * build the real file once the answers arrive. The same flow runs in
 * both placements of the one panel: the rail's Build mode, and the card
 * embedded on the Home landing where the Project AI strip used to be.
 */
test.describe('home build chat', () => {
  let root: string

  test.beforeEach(() => {
    root = realpathSync(mkdtempSync(join(tmpdir(), 'build-chat-e2e-')))
  })

  test.afterEach(() => {
    rmSync(root, { recursive: true, force: true })
  })

  test('asks the template questions first, then builds the file', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root, starPrompt: { resolved: true } },
      videoDir: 'home-build-chat',
    })
    const { page } = launched
    try {
      await page.getByRole('button', { name: 'Build', exact: true }).click()
      await expect(page.locator('.build-chat')).toBeVisible()
      // a fresh profile has no AI configured: the panel says so plainly
      await expect(page.locator('.build-chat-nomodel')).toBeVisible()

      await page.locator('.build-chat-input').fill('I need a payment receipt')
      await page.locator('.build-chat-composer .build-chat-btn-primary').click()

      // the clarify gate, live: questions about the template's
      // placeholders, and no file on disk yet
      const questions = page.locator('.build-chat-questions li')
      await expect(questions).toHaveCount(2, { timeout: 30_000 })
      await expect(questions.nth(0)).toContainText('[Item 1]')
      await expect(questions.nth(1)).toContainText('[Item 2]')
      expect(existsSync(join(root, 'Payment Receipt.xlsx'))).toBe(false)

      await page.locator('.build-chat-input').fill('Keyboard\nMouse')
      await page.locator('.build-chat-composer .build-chat-btn-primary').click()

      // the answers complete the build: a real spreadsheet in the save folder
      await expect
        .poll(() => existsSync(join(root, 'Payment Receipt.xlsx')), { timeout: 30_000 })
        .toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'home-build-chat')
    }
  })

  test('the landing embeds the chat where the Project AI strip used to be', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root, starPrompt: { resolved: true } },
      videoDir: 'home-build-chat-landing',
    })
    const { page } = launched
    try {
      // the search-first landing is what a fresh profile opens on
      const landing = page.locator('.landing-content')
      await expect(landing).toBeVisible()

      // the static Project AI strip is gone from the landing…
      await expect(page.locator('.landing-ai-strip')).toHaveCount(0)
      await expect(landing).not.toContainText(
        'Project AI can build a business plan across Docs, Sheets and Slides.',
      )

      // …and the Build chat card sits in its place, above the recents
      const card = page.locator('.landing-column .build-chat-embedded')
      await expect(card).toBeVisible()
      await expect(card.locator('.build-chat-input')).toBeVisible()
      const chatAboveRecents = await page.evaluate(() => {
        const cardEl = document.querySelector('.landing-column .build-chat-embedded')
        const recentsEl = document.querySelector('.landing-recents')
        if (!cardEl || !recentsEl) return false
        return Boolean(cardEl.compareDocumentPosition(recentsEl) & Node.DOCUMENT_POSITION_FOLLOWING)
      })
      expect(chatAboveRecents).toBe(true)

      // a fresh profile has no AI configured: the card says so plainly
      await expect(card.locator('.build-chat-nomodel')).toBeVisible()

      await card.locator('.build-chat-input').fill('I need a payment receipt')
      await card.locator('.build-chat-composer .build-chat-btn-primary').click()

      // the clarify gate, live on the landing: questions first, no file yet
      const questions = card.locator('.build-chat-questions li')
      await expect(questions).toHaveCount(2, { timeout: 30_000 })
      await expect(questions.nth(0)).toContainText('[Item 1]')
      await expect(questions.nth(1)).toContainText('[Item 2]')
      expect(existsSync(join(root, 'Payment Receipt.xlsx'))).toBe(false)

      await card.locator('.build-chat-input').fill('Keyboard\nMouse')
      await card.locator('.build-chat-composer .build-chat-btn-primary').click()

      // the answers complete the build from the landing embed
      await expect
        .poll(() => existsSync(join(root, 'Payment Receipt.xlsx')), { timeout: 30_000 })
        .toBe(true)
      await expect(card.locator('.build-chat-files .build-chat-file')).toHaveCount(1)
      await expect(card.locator('.build-chat-file-title')).toContainText('Payment Receipt')
    } finally {
      await closeAndSaveVideo(launched, 'home-build-chat-landing')
    }
  })
})
