import { mkdirSync, readFileSync, writeFileSync } from 'node:fs'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { isTechpartnerLauncher, mcpSnippet, readMcpEntry, writeMcpEntry } from '../src/agent-mcp'
import { mcpLaunch, mcpLaunchFromLauncher, type McpLaunch } from '../src/mcp-launch'
import { tempDir } from './helpers'

const WIN_DIR = 'C:\\Users\\Jane Doe\\AppData\\Local\\Programs\\TechPartner Office\\resources\\cli'
const WIN_APP: McpLaunch = {
  command: `${WIN_DIR}\\..\\..\\TechPartner Office.exe`,
  args: [`${WIN_DIR}\\techpartner.cjs`, 'mcp'],
  env: { ELECTRON_RUN_AS_NODE: '1' },
}

// The pre-rename app shape: entries written before the rename look like this.
const LEGACY_WIN_DIR = 'C:\\Users\\Jane Doe\\AppData\\Local\\Programs\\GenOffice\\resources\\cli'
const LEGACY_WIN_APP: McpLaunch = {
  command: `${LEGACY_WIN_DIR}\\..\\..\\GenOffice.exe`,
  args: [`${LEGACY_WIN_DIR}\\genoffice.cjs`, 'mcp'],
  env: { ELECTRON_RUN_AS_NODE: '1' },
}

describe('mcp launch entry', () => {
  it('is the absolute launcher on macOS and Linux', () => {
    expect(
      mcpLaunchFromLauncher(
        '/Applications/TechPartner Office.app/Contents/Resources/cli/techpartner',
      ),
    ).toEqual({
      command: '/Applications/TechPartner Office.app/Contents/Resources/cli/techpartner',
      args: ['mcp'],
    })
  })

  it('uses the absolute launcher while the linked dir is still off the PATH (genoffice#1019)', () => {
    const launcherDir = '/Applications/TechPartner Office.app/Contents/Resources/cli'
    expect(mcpLaunch({ status: 'present', launcherDir })).toEqual({
      command: 'techpartner',
      args: ['mcp'],
    })
    expect(
      mcpLaunch({
        status: 'present',
        launcherDir,
        pathHint: 'export PATH="$HOME/.local/bin:$PATH"',
      }),
    ).toEqual({ command: `${launcherDir}/techpartner`, args: ['mcp'] })
  })

  it('runs the packaged Windows app as Node on the bundle, exactly as the app snippet does', () => {
    const exists = (p: string) => p === `${WIN_DIR}\\..\\..\\TechPartner Office.exe`
    const fromCli = mcpLaunchFromLauncher(`${WIN_DIR}\\techpartner.cmd`, { exists })
    expect(fromCli).toEqual(WIN_APP)
    expect(fromCli).toEqual(mcpLaunch({ status: 'missing', launcherDir: WIN_DIR }))
    expect(fromCli).toEqual(mcpLaunch({ status: 'present', launcherDir: WIN_DIR }))
  })

  it('runs a Windows checkout on the system node with the built bundle', () => {
    const dir = 'D:\\src\\techpartner-office\\packages\\cli\\bin'
    expect(mcpLaunchFromLauncher(`${dir}\\techpartner`, { exists: () => false })).toEqual({
      command: 'node',
      args: [`${dir}\\..\\dist\\techpartner.cjs`, 'mcp'],
    })
    expect(
      mcpLaunchFromLauncher('/mnt/src/packages/cli/bin/techpartner', {
        platform: 'win32',
        exists: (p) => p.endsWith('/bin/techpartner.cjs'),
      }),
    ).toEqual({ command: 'node', args: ['/mnt/src/packages/cli/bin/techpartner.cjs', 'mcp'] })
  })

  it('recognises every launcher shape as ours', () => {
    expect(isTechpartnerLauncher('/opt/TechPartner Office/resources/cli/techpartner')).toBe(true)
    expect(isTechpartnerLauncher('C:\\TechPartner Office\\resources\\cli\\techpartner.cmd')).toBe(
      true,
    )
    // the pre-rename launcher name still counts as ours
    expect(isTechpartnerLauncher('C:\\Program Files\\GenOffice\\GenOffice.exe')).toBe(true)
    expect(isTechpartnerLauncher(WIN_APP.command, WIN_APP.args)).toBe(true)
    expect(
      isTechpartnerLauncher('node', ['C:\\src\\packages\\cli\\dist\\techpartner.cjs', 'mcp']),
    ).toBe(true)
    expect(isTechpartnerLauncher('npx', ['-y', 'other-mcp'])).toBe(false)
    expect(isTechpartnerLauncher(null)).toBe(false)
  })

  it('carries env in every config format', () => {
    const home = tempDir()
    for (const d of ['.cursor', '.gemini', '.copilot', '.codex', 'opencode'])
      mkdirSync(join(home, d))
    const cursor = join(home, '.cursor', 'mcp.json')
    writeMcpEntry('cursor', cursor, WIN_APP)
    expect(JSON.parse(readFileSync(cursor, 'utf-8')).mcpServers.techpartner).toEqual({
      type: 'stdio',
      ...WIN_APP,
    })
    const gemini = join(home, '.gemini', 'settings.json')
    writeMcpEntry('gemini', gemini, WIN_APP)
    expect(JSON.parse(readFileSync(gemini, 'utf-8')).mcpServers.techpartner).toEqual(WIN_APP)
    const copilot = join(home, '.copilot', 'mcp-config.json')
    writeMcpEntry('copilot', copilot, WIN_APP)
    expect(JSON.parse(readFileSync(copilot, 'utf-8')).mcpServers.techpartner).toEqual({
      type: 'local',
      ...WIN_APP,
      tools: ['*'],
    })
    const opencode = join(home, 'opencode', 'opencode.json')
    writeMcpEntry('opencode', opencode, WIN_APP)
    expect(JSON.parse(readFileSync(opencode, 'utf-8')).mcp.techpartner).toEqual({
      type: 'local',
      command: [WIN_APP.command, ...WIN_APP.args],
      enabled: true,
      environment: WIN_APP.env,
    })
    const codex = join(home, '.codex', 'config.toml')
    writeFileSync(codex, '[mcp_servers.other]\ncommand = "npx"\n')
    writeMcpEntry('codex', codex, WIN_APP)
    expect(readFileSync(codex, 'utf-8')).toBe(
      '[mcp_servers.other]\ncommand = "npx"\n\n' +
        `[mcp_servers.techpartner]\ncommand = ${JSON.stringify(WIN_APP.command)}\nargs = ${JSON.stringify(WIN_APP.args)}\n\n` +
        '[mcp_servers.techpartner.env]\nELECTRON_RUN_AS_NODE = "1"\n',
    )
    for (const [id, file] of [
      ['cursor', cursor],
      ['gemini', gemini],
      ['copilot', copilot],
      ['opencode', opencode],
      ['codex', codex],
    ] as const) {
      expect(readMcpEntry(id, file, WIN_APP).status, id).toBe('registered')
      expect(
        readMcpEntry(id, file, { command: '/elsewhere/cli/techpartner', args: ['mcp'] }).status,
        id,
      ).toBe('stale')
    }
    expect(mcpSnippet('codex', WIN_APP)).toContain('[mcp_servers.techpartner.env]')
  })

  it('treats a legacy (pre-rename) app snippet entry (GenOffice.exe on genoffice.cjs) as ours, not occupied', () => {
    const home = tempDir()
    const file = join(home, 'mcp.json')
    writeFileSync(file, JSON.stringify({ mcpServers: { genoffice: LEGACY_WIN_APP } }))
    const posix = { command: '/opt/TechPartner Office/resources/cli/techpartner', args: ['mcp'] }
    expect(readMcpEntry('cursor', file, posix)).toEqual({
      status: 'stale',
      command: LEGACY_WIN_APP.command,
    })
    writeFileSync(
      file,
      JSON.stringify({
        mcpServers: { genoffice: { command: 'node', args: ['/src/dist/genoffice.cjs', 'mcp'] } },
      }),
    )
    expect(readMcpEntry('cursor', file, posix).status).toBe('stale')
    const toml = join(home, 'config.toml')
    writeFileSync(
      toml,
      `[mcp_servers.genoffice]\ncommand = ${JSON.stringify(LEGACY_WIN_APP.command)}\nargs = ${JSON.stringify(LEGACY_WIN_APP.args)}\n[mcp_servers.genoffice.env]\nELECTRON_RUN_AS_NODE = "1"\n`,
    )
    expect(readMcpEntry('codex', toml, posix).status).toBe('stale')
    // a legacy table is never 'registered' under the new name: it reads stale so install migrates it
    expect(readMcpEntry('codex', toml, LEGACY_WIN_APP).status).toBe('stale')
  })
})
