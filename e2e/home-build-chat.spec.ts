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

  test('starter chips on the landing send their prompt and start the flow', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root, starPrompt: { resolved: true } },
      videoDir: 'home-build-chat-starters',
    })
    const { page } = launched
    try {
      const card = page.locator('.landing-column .build-chat-embedded')
      await expect(card).toBeVisible()

      // a fresh conversation offers the four one-tap starters
      const chips = card.locator('.build-chat-chip')
      await expect(chips).toHaveCount(4)
      await expect(chips.nth(0)).toHaveText('Make an invoice')
      await expect(chips.nth(1)).toHaveText('Build a business plan')
      await expect(chips.nth(2)).toHaveText('Plan a budget spreadsheet')
      await expect(chips.nth(3)).toHaveText('Create a presentation')

      await chips.nth(0).click()

      // the tap sent the prompt as a user message…
      await expect(card.locator('.build-chat-msg-user')).toContainText('Make an invoice')
      // …and the template question flow started: the invoice template's
      // placeholders are asked first, and nothing is on disk yet
      const questions = card.locator('.build-chat-questions li')
      await expect(questions).toHaveCount(3, { timeout: 30_000 })
      await expect(questions.nth(0)).toContainText('[Product 1]')
      expect(existsSync(join(root, 'Product Sales Invoice.xlsx'))).toBe(false)

      // the starters belong to a fresh conversation only
      await expect(chips).toHaveCount(0)
    } finally {
      await closeAndSaveVideo(launched, 'home-build-chat-starters')
    }
  })

  test('the chat history survives navigating away and back', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root, starPrompt: { resolved: true } },
      videoDir: 'home-build-chat-history',
    })
    const { page } = launched
    try {
      const card = page.locator('.landing-column .build-chat-embedded')
      await expect(card).toBeVisible()
      await card.locator('.build-chat-input').fill('I need a payment receipt')
      await card.locator('.build-chat-composer .build-chat-btn-primary').click()
      await expect(card.locator('.build-chat-questions li')).toHaveCount(2, {
        timeout: 30_000,
      })

      // over to the rail's Build view: the same conversation is there,
      // because the transcript lives in the service, not in the panel
      await page.getByRole('button', { name: 'Build', exact: true }).click()
      const rail = page.locator('main.build-chat')
      await expect(rail.locator('.build-chat-msg-user')).toContainText('I need a payment receipt')
      await expect(rail.locator('.build-chat-questions li')).toHaveCount(2)

      // and back Home: the landing card still shows the whole exchange
      await page.locator('.rail-logo').click()
      const cardAgain = page.locator('.landing-column .build-chat-embedded')
      await expect(cardAgain.locator('.build-chat-msg-user')).toContainText(
        'I need a payment receipt',
      )
      await expect(cardAgain.locator('.build-chat-questions li')).toHaveCount(2, {
        timeout: 30_000,
      })
    } finally {
      await closeAndSaveVideo(launched, 'home-build-chat-history')
    }
  })

  test('the landing highlights templates between the chat card and Recent', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      settings: { defaultSaveDir: root, starPrompt: { resolved: true } },
      videoDir: 'home-landing-templates',
    })
    const { page } = launched
    try {
      const section = page.locator('.landing-templates')
      await expect(section).toBeVisible()
      await expect(section.locator('.section-label')).toHaveText('Templates')

      // DOM order in the landing column: chat card, Templates, Recent
      const order = await page.evaluate(() => {
        const chat = document.querySelector('.landing-column .build-chat-embedded')
        const templates = document.querySelector('.landing-templates')
        const recents = document.querySelector('.landing-recents')
        if (!chat || !templates || !recents) return false
        const follows = (a: Element, b: Element) =>
          Boolean(a.compareDocumentPosition(b) & Node.DOCUMENT_POSITION_FOLLOWING)
        return follows(chat, templates) && follows(templates, recents)
      })
      expect(order).toBe(true)

      // a fresh profile gets the curated spread: eight cards, each with
      // the gallery's live visual preview
      const cards = section.locator('.template-card')
      await expect(cards).toHaveCount(8)
      await expect(section.locator('.template-visual')).toHaveCount(8)

      // See all opens the full gallery; close it again for the create step
      await section.getByRole('button', { name: 'See all templates' }).click()
      await expect(page.locator('.template-gallery')).toBeVisible()
      await page.locator('.template-gallery-close').click()
      await expect(page.locator('.template-gallery')).toHaveCount(0)

      // using a known card creates its real file in the save folder,
      // through the same path the gallery's Use button takes
      const invoice = section.locator('.template-card', { hasText: 'Service Invoice' })
      await expect(invoice).toHaveCount(1)
      await invoice.locator('.template-card-use').click()
      await expect
        .poll(() => existsSync(join(root, 'Service Invoice.xlsx')), { timeout: 30_000 })
        .toBe(true)
    } finally {
      await closeAndSaveVideo(launched, 'home-landing-templates')
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
