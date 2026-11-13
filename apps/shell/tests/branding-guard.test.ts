import { readdirSync, readFileSync, statSync } from 'node:fs'
import { dirname, join } from 'node:path'
import { fileURLToPath } from 'node:url'
import { describe, expect, it } from 'vitest'

/**
 * Branding guard (2026-11-12 rename): the old product name must never come
 * back in the surfaces a user or an assistant actually reads — the shell's
 * UI strings in every language, the bundled skill the app installs, and the
 * help topics. Internal identifiers that never render (package scopes,
 * legacy-migration constants, file-format markers) live elsewhere and are
 * deliberately out of this test's scope.
 */

const REPO_ROOT = join(dirname(fileURLToPath(import.meta.url)), '..', '..', '..')
const NEEDLE = /genoffice/i

function filesUnder(dir: string): string[] {
  const out: string[] = []
  for (const entry of readdirSync(dir)) {
    const path = join(dir, entry)
    if (statSync(path).isDirectory()) out.push(...filesUnder(path))
    else out.push(path)
  }
  return out
}

function offenders(files: string[]): string[] {
  return files.filter((f) => NEEDLE.test(readFileSync(f, 'utf-8')))
}

describe('branding guard: no trace of the old product name in user-facing copy', () => {
  it('shell renderer strings (all languages)', () => {
    const file = join(REPO_ROOT, 'apps/shell/src/renderer/src/strings.ts')
    expect(offenders([file])).toEqual([])
  })

  it('bundled skill content under skills/', () => {
    const files = filesUnder(join(REPO_ROOT, 'skills'))
    expect(files.length).toBeGreaterThan(0)
    expect(offenders(files)).toEqual([])
  })

  it('user-visible help topics', () => {
    const files = filesUnder(join(REPO_ROOT, 'apps/shell/src/renderer/src/i18n/help/topics'))
    expect(files.length).toBeGreaterThan(0)
    expect(offenders(files)).toEqual([])
  })
})

describe('branding guard: no upstream AI brand in the app AI panel titles', () => {
  it('apps aiPanelTitle / ribbonAiAssistant values contain no Genspark', () => {
    const files: string[] = []
    for (const app of ['docs', 'sheets', 'slides', 'pdf', 'markdown', 'html']) {
      files.push(
        ...filesUnder(join(REPO_ROOT, 'apps', app, 'src/renderer/i18n')).filter((f) =>
          f.endsWith('.ts'),
        ),
      )
    }
    expect(files.length).toBeGreaterThan(0)
    const bad = files.filter((f) =>
      /^\s*(?:aiPanelTitle|ribbonAiAssistant):.*genspark/im.test(readFileSync(f, 'utf-8')),
    )
    expect(bad).toEqual([])
  })
})
