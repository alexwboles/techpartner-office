import { mkdtemp, writeFile } from 'node:fs/promises'
import { tmpdir } from 'node:os'
import { join } from 'node:path'
import { describe, expect, it } from 'vitest'
import { probeOllamaStatus, type OllamaFetcher } from '../src/main/ollama-probe'

/**
 * The Settings Ollama guide probe: four statuses from the configured
 * local endpoint, with the network replaced by an injected fetcher.
 * The settings file fixtures are the minimal stored shape —
 * resolveAiSettings merges them over the Ollama-prefilled defaults.
 */

async function userDataWithSettings(settings: unknown): Promise<string> {
  const dir = await mkdtemp(join(tmpdir(), 'tp-ollama-probe-'))
  if (settings !== undefined) {
    await writeFile(join(dir, 'ai-settings.json'), JSON.stringify(settings), 'utf8')
  }
  return dir
}

function storedWithBaseUrl(baseUrl: string): unknown {
  return {
    provider: 'custom',
    providers: { custom: { apiKey: '', model: '', baseUrl } },
  }
}

function jsonFetcher(body: unknown, ok = true): OllamaFetcher {
  return async () => ({ ok, json: async () => body })
}

describe('probeOllamaStatus', () => {
  it('reports ready with the model list when Ollama answers', async () => {
    const dir = await userDataWithSettings(storedWithBaseUrl('http://localhost:11434/v1'))
    const result = await probeOllamaStatus(
      dir,
      jsonFetcher({ data: [{ id: 'llama3.1:latest' }, { id: 'qwen2.5:7b' }] }),
    )
    expect(result).toEqual({ status: 'ready', models: ['llama3.1:latest', 'qwen2.5:7b'] })
  })

  it('reports no-models when Ollama runs with nothing pulled', async () => {
    const dir = await userDataWithSettings(storedWithBaseUrl('http://localhost:11434/v1'))
    const result = await probeOllamaStatus(dir, jsonFetcher({ data: [] }))
    expect(result).toEqual({ status: 'no-models' })
  })

  it('reports unreachable when the fetch throws', async () => {
    const dir = await userDataWithSettings(storedWithBaseUrl('http://localhost:11434/v1'))
    const result = await probeOllamaStatus(dir, async () => {
      throw new Error('connect ECONNREFUSED')
    })
    expect(result).toEqual({ status: 'unreachable' })
  })

  it('reports unreachable on a non-OK response', async () => {
    const dir = await userDataWithSettings(storedWithBaseUrl('http://localhost:11434/v1'))
    const result = await probeOllamaStatus(dir, jsonFetcher({}, false))
    expect(result).toEqual({ status: 'unreachable' })
  })

  it('reports not-configured without fetching when the local slot points elsewhere', async () => {
    const dir = await userDataWithSettings(storedWithBaseUrl('https://api.example.com/v1'))
    let fetched = false
    const result = await probeOllamaStatus(dir, async () => {
      fetched = true
      return { ok: true, json: async () => ({ data: [{ id: 'x' }] }) }
    })
    expect(result).toEqual({ status: 'not-configured' })
    expect(fetched).toBe(false)
  })

  it('probes the default Ollama endpoint when no settings file exists yet', async () => {
    const dir = await userDataWithSettings(undefined)
    const urls: string[] = []
    const result = await probeOllamaStatus(dir, async (url) => {
      urls.push(url)
      return { ok: true, json: async () => ({ data: [{ id: 'llama3.1:latest' }] }) }
    })
    expect(result.status).toBe('ready')
    expect(urls).toEqual(['http://localhost:11434/v1/models'])
  })

  it('accepts a base URL without the /v1 suffix', async () => {
    const dir = await userDataWithSettings(storedWithBaseUrl('http://127.0.0.1:11434'))
    const urls: string[] = []
    const result = await probeOllamaStatus(dir, async (url) => {
      urls.push(url)
      return { ok: true, json: async () => ({ data: [] }) }
    })
    expect(result).toEqual({ status: 'no-models' })
    expect(urls).toEqual(['http://127.0.0.1:11434/models'])
  })
})
