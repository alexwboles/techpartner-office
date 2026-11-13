import { existsSync, readFileSync } from 'node:fs'
import { join } from 'node:path'
import {
  defaultAiSettings,
  resolveAiSettings,
  resolveCustomEndpoints,
} from '@genoffice/ai-provider'
import type { OllamaProbeResult } from '../shared/home-api'

/**
 * The Settings Ollama guide's probe (main process): read the saved AI
 * settings the same way the chat services do, and when the local-AI
 * ("custom") slot points at an Ollama endpoint on this machine, ask it
 * for its model list. Four plain outcomes the card renders directly:
 * ready (models listed), no-models (running, nothing pulled yet),
 * unreachable (Ollama not answering), not-configured (the local slot
 * points somewhere else entirely). Read-only: nothing is installed,
 * started or changed here — the guide only ever tells the user what
 * to do, and opens the download page when they ask.
 */

const PROBE_TIMEOUT_MS = 2_500

/** Ollama's own address on this machine (any path suffix allowed). */
const OLLAMA_LOCAL_ENDPOINT = /^https?:\/\/(localhost|127\.0\.0\.1):11434(?=\/|$)/i

/** The slice of fetch the probe uses — injected in tests. */
export type OllamaFetcher = (
  url: string,
  init: { signal: AbortSignal },
) => Promise<{ ok: boolean; json: () => Promise<unknown> }>

const defaultFetcher: OllamaFetcher = (url, init) => fetch(url, init)

export async function probeOllamaStatus(
  userDataPath: string,
  fetcher: OllamaFetcher = defaultFetcher,
): Promise<OllamaProbeResult> {
  const settingsPath = join(userDataPath, 'ai-settings.json')
  let stored: unknown
  if (existsSync(settingsPath)) {
    try {
      stored = JSON.parse(readFileSync(settingsPath, 'utf8'))
    } catch {
      stored = undefined
    }
  }
  // resolveCustomEndpoints mirrors the active endpoint into the
  // custom slot, so the probe sees the same address the settings UI
  // shows (endpoint-backed and legacy-slot settings alike)
  const settings = resolveCustomEndpoints(resolveAiSettings(stored, defaultAiSettings()))
  const baseUrl = (settings.providers?.custom?.baseUrl ?? '').trim()
  if (!OLLAMA_LOCAL_ENDPOINT.test(baseUrl)) return { status: 'not-configured' }

  const controller = new AbortController()
  const timer = setTimeout(() => controller.abort(), PROBE_TIMEOUT_MS)
  try {
    const response = await fetcher(`${baseUrl.replace(/\/+$/, '')}/models`, {
      signal: controller.signal,
    })
    if (!response.ok) return { status: 'unreachable' }
    const body: unknown = await response.json().catch(() => null)
    const data =
      body && typeof body === 'object' && Array.isArray((body as { data?: unknown }).data)
        ? (body as { data: unknown[] }).data
        : []
    const models = data
      .map((entry) =>
        entry && typeof entry === 'object' && typeof (entry as { id?: unknown }).id === 'string'
          ? (entry as { id: string }).id
          : '',
      )
      .filter((id) => id.length > 0)
    return models.length > 0 ? { status: 'ready', models } : { status: 'no-models' }
  } catch {
    return { status: 'unreachable' }
  } finally {
    clearTimeout(timer)
  }
}
