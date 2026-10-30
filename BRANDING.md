# TechPartner Office — branding note

## What this is

A rebrand of GenOffice (Apache-2.0, https://github.com/genspark-ai/genoffice)
as **TechPartner Office**, by TechPartner Consulting
(https://www.techpartner.base44.app).

Original code and attribution are NOT removed: the Apache-2.0 LICENSE and
NOTICE files stay in place. What changes is the product identity, because
Apache-2.0 does not grant trademark permission, and GenOffice/Genspark are
trademarks of Mainfunc, Inc. (see upstream README). A fork must not ship
under those names or logos.

## TechPartner identity used here

- Product: **TechPartner Office**
- Apps: TechPartner Docs, TechPartner Sheets, TechPartner Slides,
  TechPartner PDF, plus Markdown and HTML editors
- Company: TechPartner Consulting
- Logo: `apps/shell/src/renderer/src/assets/techpartner-logo.png`
  (the TechPartner network-mark used on Alex's fundraiser-page credit,
  72x72 PNG, dark navy mark)
- Palette, from the TechPartner Business Dashboard:
  - Accent: `#1c5b4a` (deep green), deep accent `#123f34`
  - Dark rail: `#131e19`
  - Paper: `#f6f7f4`, ink `#1b241f`
  - No gradients, no rounded-edge chrome in the new shell styling
    (the logo file itself is kept as supplied)

## Done in this pilot (shell + shared chrome)

- `apps/shell/package.json`: productName TechPartner Office,
  desktopName techpartner-office.desktop, author TechPartner Consulting
- `apps/shell/electron-builder.cjs`: appId `com.techpartner.office`,
  productName TechPartner Office
- Shell window titles: TechPartner Office
- Home screen logo: TechPartner logo, alt "TechPartner Office"
- Shared UI tokens (`packages/ui/src/tokens.css`): primary button and
  brand colours moved to TechPartner green `#1c5b4a` / `#123f34`
- Upstream remote removed from this copy, so it cannot be pushed
  back to genspark-ai by accident

## Alex's rule (2026-10-28): usability first

**If a rebrand change would affect usability, do not change it.**
That rules out the full internal sweep (Phase 3, not started):

- **NOT changed, on purpose:** package scopes (`@genoffice/*`), the
  `genoffice` CLI/MCP names, internal file/user-data handling, engines,
  and file formats. Renaming those would break builds, skills, agent
  tools, and existing workflows for no user benefit.
- **Root package name stays `genoffice` — lesson learned 2026-10-28.**
  The pilot briefly renamed the root `package.json` to
  `techpartner-office`, and that broke the CLI: `packages/cli/src/resources.ts`
  finds the repo root by checking `package.json` name === `genoffice`,
  so with the new name it returned null and every Word/Markdown CLI
  operation failed with "jsdom not found", failing CI's CLI and MCP
  tests (while upstream CI at the same commit was green). The root
  name is invisible to users, so it was reverted — product name,
  appId, and window titles carry the TechPartner identity instead.

## Phase 1 + 2 (2026-10-30, Alex approved): visible branding, Genspark-free AI

This supersedes the Oct 28 note that kept the Genspark sign-in. Alex
confirmed on Oct 30 that he had opened TechPartner Office and it looks
good, and asked for a fully TechPartner-branded product with no
Genspark usage or paywall.

- **App icons:** shell (and Docs standalone) icon is now the TechPartner
  network mark in white on TechPartner green `#1c5b4a`, square (no
  rounded edges, per Alex's rule), in `apps/shell/build/icon.png` /
  `.ico` / `.icns` and `apps/shell/build/icons/`. File-type document
  icons are unchanged: they were already generic type colours, not a
  GenOffice logo.
- **Per-app display names:** TechPartner Docs / Sheets / Slides / PDF /
  Markdown / HTML in each app's `package.json` productName, window
  titles, renderer titles, and the shell's untitled-tab titles.
- **Help (English):** GenOffice → TechPartner Office and Genspark →
  TechPartner in the English help topics, with the AI, home-screen, and
  settings-integrations articles rewritten to describe no-account AI
  setup (local Ollama or your own key) instead of a hosted sign-in.
  Other languages still carry the upstream wording; they are a
  translation follow-up, not a gate failure (the help-language check
  compares structure, not wording).
- **No Genspark sign-in, anywhere in the product:**
  - Home: the Genspark Projects nav entry and view are removed; the
    bottom-left button now opens Settings, not a sign-in.
  - Settings: the account pane is now an AI setup card (free local
    Ollama at `http://localhost:11434/v1`, or your own key) with a
    button to AI model settings. The Genspark provider is filtered
    out of the Settings provider picker, and the cloud-tools switch
    is removed.
  - Main processes (shell, Docs, Sheets, Slides): account status
    always reports signed out, the login entry points are no-ops,
    cloud-project sync is disabled, and no Genspark key is ever
    injected into AI requests.
  - AI defaults (new installs, and stored Genspark selections on
    upgrade): provider Custom prefilled with the Ollama endpoint,
    cloud tools off, search default Parallel (keyless), media default
    OpenAI BYOK (unavailable until a key is entered). A half-filled
    BYOK setup falls back to Custom, never to Genspark.
  - Visible labels: "Genspark AI" → "TechPartner AI" in the Docs
    ribbon, Sheets, and Slides; onboarding credits/notes rewritten
    to local-model / own-key wording.
- **What is NOT removed:** the Genspark provider code itself still
  exists in `packages/ai-provider` (so old settings files, tests, and
  the catalog code keep working), but it is unreachable from the UI
  and never selected by the defaults or fallbacks. Removing the code
  entirely, and the deep internal rename, are Phase 3 — not started.
- **Honest limit:** cloud AI is never free. "Free AI" here means a
  local Ollama model on the user's machine, or a key the user already
  pays for elsewhere. There is no bundled free cloud AI.

## Verification so far

- `node tools/check-theme-colors.mjs` passes: "No new raw colors in renderer CSS."
- Pushed to https://github.com/alexwboles/techpartner-office (public,
  fork of genspark-ai/genoffice); first Windows release v0.12.0 built
  by GitHub Actions. The pristine upstream copy is at
  `~/workspace/genoffice-base/`.
