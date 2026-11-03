import { test, expect } from '@playwright/test'
import { launchShell, closeAndSaveVideo, screenshotPath } from './helpers'

test.describe('home screen', () => {
  test('shows search-first landing, create row and tab bar', async () => {
    const launched = await launchShell({ onboardingSeen: true, videoDir: 'home-basics' })
    const { page } = launched
    try {
      await expect(page.locator('.landing-headline')).toBeVisible()
      await expect(page.locator('.landing-search-hint')).toHaveText('Ctrl K')
      // the + New button plus six AI create buttons, then From template
      await expect(page.locator('.landing-new-btn')).toBeVisible()
      await expect(page.locator('.quick-card')).toHaveCount(6)
      await expect(page.locator('.quick-card').first()).toContainText('AI Docs')
      await expect(page.locator('.quick-card').nth(1)).toContainText('AI Sheets')
      await expect(page.locator('.quick-card').nth(2)).toContainText('AI Slides')
      await expect(page.locator('.quick-card').nth(3)).toContainText('AI Markdown')
      await expect(page.locator('.quick-card').nth(4)).toContainText('AI HTML')
      await expect(page.locator('.quick-card').nth(5)).toContainText('AI PDF')
      await expect(page.locator('.landing-template-btn')).toContainText('From template')
      await expect(page.locator('.tab-bar .tab-item.tab-home')).toBeVisible()
      await page.screenshot({ path: screenshotPath('home-overview') })
    } finally {
      await closeAndSaveVideo(launched, 'home-basics')
    }
  })

  test('renders localized UI when GENOFFICE_LANG=zh-CN', async () => {
    const launched = await launchShell({
      onboardingSeen: true,
      lang: 'zh-CN',
      videoDir: 'home-zh-cn',
    })
    const { page } = launched
    try {
      await expect(page.getByRole('button', { name: '最近', exact: true }).first()).toBeVisible()
      await page.screenshot({ path: screenshotPath('home-zh-cn') })
    } finally {
      await closeAndSaveVideo(launched, 'home-zh-cn')
    }
  })
})
