/// The default folder where new/untitled files land on their first (silent)
/// save and where AI-generated drafts go. Historically hardcoded to
/// <Documents>/GenOffice; the rebrand moved the fallback to
/// <Documents>/TechPartner Office, and a one-time startup migration
/// (migrateLegacySaveDir) renames an existing legacy folder for users who
/// never configured a custom folder. The folder is otherwise
/// user-configurable via the `defaultSaveDir` key in
/// userData/app-settings.json. Every editor main module resolves through
/// here so they all honor the same setting.
import {
  accessSync,
  constants,
  existsSync,
  mkdirSync,
  readFileSync,
  readdirSync,
  renameSync,
  statSync,
  writeFileSync,
} from 'node:fs'
import { dirname, isAbsolute, join, resolve } from 'node:path'

/** the subset of Electron's `app` needed here (kept structural: this package has no Electron dependency) */
export interface PathProvider {
  getPath(name: 'userData' | 'documents'): string
}

export const DEFAULT_SAVE_DIR_KEY = 'defaultSaveDir'

/**
 * Create `dir` and any missing parents one level at a time.
 *
 * `mkdirSync(dir, { recursive: true })` cannot be used on the startup path: when
 * the filesystem refuses the leaf with ENOENT while its parent exists, libuv's
 * recursive implementation keeps re-attempting the parent (EEXIST) and the leaf
 * (ENOENT) forever and pins the calling process at 100% CPU. Windows Defender's
 * Controlled Folder Access does exactly that for Documents, which froze the
 * whole app on launch. A plain per-level mkdir lets the refusal surface as a
 * thrown error that the callers below already handle.
 */
export function mkdirDirs(dir: string): void {
  const missing: string[] = []
  let probe = resolve(dir)
  for (;;) {
    if (existsSync(probe)) break
    missing.push(probe)
    const parent = dirname(probe)
    if (parent === probe) break
    probe = parent
  }
  for (const path of missing.reverse()) mkdirSync(path)
}

/** the configured folder from app-settings.json, or null when unset/unreadable */
export function readDefaultSaveDirSetting(settingsPath: string): string | null {
  try {
    const raw: unknown = JSON.parse(readFileSync(settingsPath, 'utf8'))
    if (!raw || typeof raw !== 'object' || Array.isArray(raw)) return null
    const value = (raw as Record<string, unknown>)[DEFAULT_SAVE_DIR_KEY]
    return typeof value === 'string' && isAbsolute(value) ? value : null
  } catch {
    return null
  }
}

/** true when the directory exists (or could be created) and is writable */
export function isUsableSaveDir(dir: string): boolean {
  try {
    mkdirDirs(dir)
    if (!statSync(dir).isDirectory()) return false
    accessSync(dir, constants.W_OK)
    return true
  } catch {
    return false
  }
}

/**
 * Pick the effective default save folder: the configured one when it is
 * usable, otherwise the fallback (created on demand). A configured folder on
 * an unplugged drive or with revoked permissions degrades to the fallback
 * instead of breaking silent saves.
 */
export function resolveDefaultSaveDir(configured: string | null, fallbackDir: string): string {
  if (configured && isUsableSaveDir(configured)) return configured
  if (!isUsableSaveDir(fallbackDir)) throw new Error(`default save dir not usable: ${fallbackDir}`)
  return fallbackDir
}

/** folder name the pre-rebrand fallback used (the left-pane folder users saw as "GenOffice") */
export const LEGACY_DEFAULT_SAVE_DIR_NAME = 'GenOffice'

/** folder name the rebrand fallback uses */
export const DEFAULT_SAVE_DIR_NAME = 'TechPartner Office'

/** convenience for the Electron mains: settings lookup + fallback in one call */
export function configuredDefaultSaveDir(app: PathProvider): string {
  const settingsPath = join(app.getPath('userData'), 'app-settings.json')
  const fallback = join(app.getPath('documents'), DEFAULT_SAVE_DIR_NAME)
  return resolveDefaultSaveDir(readDefaultSaveDirSetting(settingsPath), fallback)
}

function samePath(a: string, b: string): boolean {
  const norm = (p: string) => resolve(p)
  return process.platform === 'win32'
    ? norm(a).toLowerCase() === norm(b).toLowerCase()
    : norm(a) === norm(b)
}

function underDir(value: string, dir: string): boolean {
  if (process.platform === 'win32') {
    const v = resolve(value).toLowerCase()
    const d = resolve(dir).toLowerCase()
    return v === d || v.startsWith(d + '\\')
  }
  const v = resolve(value)
  const d = resolve(dir)
  return v === d || v.startsWith(d + '/')
}

/**
 * Rebase every string in a JSON-shaped value that names a path inside
 * `oldDir` to the same relative path inside `newDir`. Strings that merely
 * mention the old folder mid-sentence are untouched: only values that ARE
 * a path under the old folder (or the folder itself) move.
 */
export function rebaseSavedPaths(value: unknown, oldDir: string, newDir: string): unknown {
  if (typeof value === 'string' && underDir(value, oldDir)) {
    return newDir + value.slice(resolve(oldDir).length)
  }
  if (Array.isArray(value)) return value.map((entry) => rebaseSavedPaths(entry, oldDir, newDir))
  if (value && typeof value === 'object') {
    const out: Record<string, unknown> = {}
    for (const [key, entry] of Object.entries(value as Record<string, unknown>)) {
      out[key] = rebaseSavedPaths(entry, oldDir, newDir)
    }
    return out
  }
  return value
}

function rebaseJsonFile(filePath: string, oldDir: string, newDir: string): void {
  let parsed: unknown
  try {
    parsed = JSON.parse(readFileSync(filePath, 'utf8'))
  } catch {
    return
  }
  const rebased = rebaseSavedPaths(parsed, oldDir, newDir)
  if (JSON.stringify(rebased) !== JSON.stringify(parsed)) {
    writeFileSync(filePath, JSON.stringify(rebased, null, 2))
  }
}

/**
 * One-time rename of the legacy default save folder, for users who never
 * configured a custom folder. Runs at shell startup, before any folder
 * watcher or tab restore touches the save folder:
 *
 * - old folder missing → nothing to do (fresh installs already get the new fallback)
 * - new folder already exists → never merge; leave both alone
 * - user configured a different save folder → respect it, leave the legacy folder alone
 * - rename failure (locked folder, permissions) → never throw; the app keeps
 *   working with the legacy folder and retries on the next launch
 *
 * On success, saved paths pointing inside the old folder are rebased to the
 * new one in app-settings.json (starred files, last-session tabs, extra
 * folder roots, an explicit setting pointing at the old folder) and in the
 * project-store JSON files, so recents and projects keep working.
 */
export function migrateLegacySaveDir(
  documentsDir: string,
  userDataDir: string,
): { migrated: boolean; reason: string } {
  const oldDir = join(documentsDir, LEGACY_DEFAULT_SAVE_DIR_NAME)
  const newDir = join(documentsDir, DEFAULT_SAVE_DIR_NAME)
  const settingsPath = join(userDataDir, 'app-settings.json')
  try {
    const configured = readDefaultSaveDirSetting(settingsPath)
    if (configured !== null && !samePath(configured, oldDir)) {
      return { migrated: false, reason: 'custom-save-dir' }
    }
    if (!existsSync(oldDir) || !statSync(oldDir).isDirectory()) {
      return { migrated: false, reason: 'no-legacy-dir' }
    }
    if (existsSync(newDir)) {
      return { migrated: false, reason: 'target-exists' }
    }
    try {
      renameSync(oldDir, newDir)
    } catch {
      return { migrated: false, reason: 'rename-failed' }
    }
    rebaseJsonFile(settingsPath, oldDir, newDir)
    const projectsDir = join(userDataDir, 'projects')
    try {
      const walk = (dir: string): void => {
        for (const entry of readdirSync(dir, { withFileTypes: true })) {
          const entryPath = join(dir, entry.name)
          if (entry.isDirectory()) walk(entryPath)
          else if (entry.isFile() && entry.name.endsWith('.json')) {
            rebaseJsonFile(entryPath, oldDir, newDir)
          }
        }
      }
      if (existsSync(projectsDir)) walk(projectsDir)
    } catch {
      // project files are best-effort: the folder itself already moved
    }
    return { migrated: true, reason: 'renamed' }
  } catch {
    return { migrated: false, reason: 'error' }
  }
}
