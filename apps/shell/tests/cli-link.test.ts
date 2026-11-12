import { readFileSync, statSync } from 'node:fs'
import { mkdtempSync } from 'node:fs'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it, vi } from 'vitest'

vi.mock('electron', () => ({ app: { isPackaged: false, getVersion: () => '0.0.0' } }))

import { isEphemeralInstall, launcherFilePath, writeLauncherFile } from '../src/main/cli-link'

describe('techpartner launcher file', () => {
  it('lives in the techpartner-office auth directory, overridable like auth.json (TECHPARTNER_AUTH_DIR primary, GENOFFICE_AUTH_DIR fallback)', () => {
    expect(launcherFilePath({})).toBe(
      join(process.env.HOME ?? '', '.techpartner-office', 'launcher'),
    )
    expect(launcherFilePath({ TECHPARTNER_AUTH_DIR: '/tmp/x' })).toBe(join('/tmp/x', 'launcher'))
    expect(launcherFilePath({ GENOFFICE_AUTH_DIR: '/tmp/x' })).toBe(join('/tmp/x', 'launcher'))
    // the primary override wins when both are set
    expect(
      launcherFilePath({ TECHPARTNER_AUTH_DIR: '/tmp/new', GENOFFICE_AUTH_DIR: '/tmp/old' }),
    ).toBe(join('/tmp/new', 'launcher'))
  })

  it('writes one line, creates the directory, and only rewrites on change', () => {
    const dir = mkdtempSync(join(tmpdir(), 'genoffice-launcher-'))
    const file = join(dir, 'nested', 'launcher')
    expect(
      writeLauncherFile(file, '/Applications/TechPartner Office.app/Contents/Resources/cli'),
    ).toBe(true)
    expect(readFileSync(file, 'utf-8')).toBe(
      '/Applications/TechPartner Office.app/Contents/Resources/cli\n',
    )
    const before = statSync(file).mtimeMs
    expect(
      writeLauncherFile(file, '/Applications/TechPartner Office.app/Contents/Resources/cli'),
    ).toBe(false)
    expect(statSync(file).mtimeMs).toBe(before)
    expect(
      writeLauncherFile(file, 'C:\\Programs\\TechPartner Office\\resources\\techpartner'),
    ).toBe(true)
    expect(readFileSync(file, 'utf-8')).toBe(
      'C:\\Programs\\TechPartner Office\\resources\\techpartner\n',
    )
  })

  it('treats dmg and AppImage mounts as temporary', () => {
    expect(
      isEphemeralInstall(
        '/Volumes/TechPartner Office/TechPartner Office.app/Contents/Resources',
        {},
      ),
    ).toBe(true)
    expect(isEphemeralInstall('/tmp/.mount_TechPaxyz/resources', {})).toBe(true)
    expect(
      isEphemeralInstall('/opt/TechPartner Office/resources', {
        APPIMAGE: '/home/u/T.AppImage',
      }),
    ).toBe(true)
    expect(isEphemeralInstall('/Applications/TechPartner Office.app/Contents/Resources', {})).toBe(
      false,
    )
  })
})
