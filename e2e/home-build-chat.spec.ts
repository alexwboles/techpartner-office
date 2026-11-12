import { test, expect } from '@playwright/test'
import { existsSync, mkdtempSync, realpathSync, rmSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { launchShell, closeAndSaveVideo } from './helpers'

/**
 * Home Build chat in a fresh profile: no AI is configured, so the panel
 * is in template-only mode. The deterministic path must ask the matched
 * template's placeholder questions first (nothing is built yet), then
 * build the real file once the answers arrive.
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
})
