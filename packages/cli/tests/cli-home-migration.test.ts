import { existsSync, mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { cliHomeDir, migrateLegacyCliHome } from '../src/install'
import { tempDir } from './helpers'

describe('cliHomeDir', () => {
  it('defaults to ~/.techpartner-office', () => {
    expect(cliHomeDir({})).toMatch(/[\\/]\.techpartner-office$/)
  })

  it('honours TECHPARTNER_AUTH_DIR first, then the legacy GENOFFICE_AUTH_DIR', () => {
    expect(cliHomeDir({ TECHPARTNER_AUTH_DIR: '/tmp/new' })).toBe('/tmp/new')
    expect(cliHomeDir({ TECHPARTNER_AUTH_DIR: '/tmp/new', GENOFFICE_AUTH_DIR: '/tmp/old' })).toBe(
      '/tmp/new',
    )
    expect(cliHomeDir({ GENOFFICE_AUTH_DIR: '/tmp/old' })).toBe('/tmp/old')
  })
})

describe('migrateLegacyCliHome', () => {
  it('moves ~/.genoffice to ~/.techpartner-office with contents intact', () => {
    const home = tempDir()
    const legacy = join(home, '.genoffice')
    mkdirSync(join(legacy, 'launcher'), { recursive: true })
    writeFileSync(join(legacy, 'auth.json'), '{"token":"x"}\n')
    writeFileSync(join(legacy, 'launcher', 'techpartner'), '/some/dir\n')

    const r = migrateLegacyCliHome({}, home)
    expect(r).toEqual({ migrated: true, reason: 'renamed' })
    expect(existsSync(legacy)).toBe(false)
    const moved = join(home, '.techpartner-office')
    expect(readFileSync(join(moved, 'auth.json'), 'utf-8')).toBe('{"token":"x"}\n')
    expect(readFileSync(join(moved, 'launcher', 'techpartner'), 'utf-8')).toBe('/some/dir\n')
  })

  it('does nothing when there is no legacy folder', () => {
    const home = tempDir()
    expect(migrateLegacyCliHome({}, home)).toEqual({ migrated: false, reason: 'no-legacy-dir' })
    expect(existsSync(join(home, '.techpartner-office'))).toBe(false)
  })

  it('never merges into an existing new folder', () => {
    const home = tempDir()
    mkdirSync(join(home, '.genoffice'), { recursive: true })
    writeFileSync(join(home, '.genoffice', 'auth.json'), 'legacy')
    mkdirSync(join(home, '.techpartner-office'), { recursive: true })
    writeFileSync(join(home, '.techpartner-office', 'auth.json'), 'current')

    expect(migrateLegacyCliHome({}, home)).toEqual({ migrated: false, reason: 'target-exists' })
    expect(readFileSync(join(home, '.genoffice', 'auth.json'), 'utf-8')).toBe('legacy')
    expect(readFileSync(join(home, '.techpartner-office', 'auth.json'), 'utf-8')).toBe('current')
  })

  it('respects a custom auth dir override and moves nothing', () => {
    const home = tempDir()
    mkdirSync(join(home, '.genoffice'), { recursive: true })
    expect(migrateLegacyCliHome({ TECHPARTNER_AUTH_DIR: '/tmp/custom' }, home)).toEqual({
      migrated: false,
      reason: 'custom-dir',
    })
    expect(migrateLegacyCliHome({ GENOFFICE_AUTH_DIR: '/tmp/custom' }, home)).toEqual({
      migrated: false,
      reason: 'custom-dir',
    })
    expect(existsSync(join(home, '.genoffice'))).toBe(true)
    expect(existsSync(join(home, '.techpartner-office'))).toBe(false)
  })

  it('is idempotent: a second run is a no-op', () => {
    const home = tempDir()
    mkdirSync(join(home, '.genoffice'), { recursive: true })
    writeFileSync(join(home, '.genoffice', 'auth.json'), 'x')
    expect(migrateLegacyCliHome({}, home).migrated).toBe(true)
    expect(migrateLegacyCliHome({}, home)).toEqual({ migrated: false, reason: 'no-legacy-dir' })
    expect(readFileSync(join(home, '.techpartner-office', 'auth.json'), 'utf-8')).toBe('x')
  })
})
