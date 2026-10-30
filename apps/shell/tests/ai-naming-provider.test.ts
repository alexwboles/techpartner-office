import { describe, expect, it } from 'vitest'
import { resolveNamingProvider } from '../src/main/ai-naming-provider'

/**
 * The naming service used to resolve its provider on its own: a hardcoded list
 * of three ids, and the key read out of the settings file. That list did not
 * know every provider the build can talk to, and it never saw the Genspark
 * key — which lives in the login state, not in the file — so for the default
 * account the call failed and the feature quietly did nothing.
 *
 * These assertions are the evidence that the provider is now resolved the way
 * the rest of the app resolves it.
 */

describe('resolveNamingProvider', () => {
  it('never uses a Genspark key from the login state', () => {
    // TechPartner Office has no Genspark account: a passed-in key is ignored,
    // and with no model configured naming declines instead of calling Genspark.
    expect(resolveNamingProvider({}, 'gsk-live-key')).toBeNull()
  })

  it('declines when no model is configured', () => {
    expect(resolveNamingProvider({}, '')).toBeNull()
  })

  it('declines when the settings file could not be read at all', () => {
    expect(resolveNamingProvider(null, '')).toBeNull()
  })

  it('honors a configured provider', () => {
    const resolved = resolveNamingProvider(
      {
        provider: 'anthropic',
        providers: { anthropic: { apiKey: 'sk-ant', model: 'claude-sonnet-4-6' } },
      },
      '',
    )
    expect(resolved?.provider).toBe('anthropic')
    expect(resolved?.config.apiKey).toBe('sk-ant')
    expect(resolved?.config.model).toBe('claude-sonnet-4-6')
  })

  it('falls back to Custom when the configured provider is half-filled', () => {
    const resolved = resolveNamingProvider(
      {
        provider: 'anthropic',
        providers: {
          anthropic: { apiKey: 'sk-ant', model: '' },
          custom: { apiKey: '', model: 'llama3', baseUrl: 'http://localhost:11434/v1' },
        },
      },
      'gsk-live-key',
    )
    expect(resolved?.provider).toBe('custom')
    expect(resolved?.config.model).toBe('llama3')
  })

  it('falls back to Custom for a provider this build does not have', () => {
    const resolved = resolveNamingProvider(
      {
        provider: 'not-a-provider',
        providers: {
          custom: { apiKey: '', model: 'llama3', baseUrl: 'http://localhost:11434/v1' },
        },
      },
      'gsk-live-key',
    )
    expect(resolved?.provider).toBe('custom')
  })

  it('does not need a key for codex', () => {
    const resolved = resolveNamingProvider({ provider: 'codex', providers: {} }, '')
    expect(resolved?.provider).toBe('codex')
  })

  it("resolves a custom endpoint's key from the environment", () => {
    process.env.GENOFFICE_NAMING_TEST_KEY = 'from-env'
    try {
      const resolved = resolveNamingProvider(
        {
          provider: 'custom',
          providers: {
            custom: {
              apiKey: '',
              model: 'local-model',
              baseUrl: 'http://127.0.0.1:11434/v1',
              profile: { apiKeyEnv: 'GENOFFICE_NAMING_TEST_KEY' },
            },
          },
        },
        '',
      )
      expect(resolved?.provider).toBe('custom')
      expect(resolved?.config.apiKey).toBe('from-env')
    } finally {
      delete process.env.GENOFFICE_NAMING_TEST_KEY
    }
  })
})
