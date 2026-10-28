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
That rules out the full internal sweep, and settles the AI question:

- **NOT changed, on purpose:** package scopes (`@genoffice/*`), the
  `genoffice` CLI/MCP names, internal file/user-data handling, engines,
  file formats, and the AI provider code. Renaming those would break
  builds, skills, agent tools, and existing workflows for no user benefit.
- **Root package name stays `genoffice` — lesson learned 2026-10-28.**
  The pilot briefly renamed the root `package.json` to
  `techpartner-office`, and that broke the CLI: `packages/cli/src/resources.ts`
  finds the repo root by checking `package.json` name === `genoffice`,
  so with the new name it returned null and every Word/Markdown CLI
  operation failed with "jsdom not found", failing CI's CLI and MCP
  tests (while upstream CI at the same commit was green). The root
  name is invisible to users, so it was reverted — product name,
  appId, and window titles carry the TechPartner identity instead.
- **Genspark AI sign-in: kept, and kept honest.** The AI panel still
  offers Genspark (gsk) login plus bring-your-own-key, exactly as
  upstream. It is not relabelled as a TechPartner service. Removing it
  would reduce usability; keeping it under a false name would be worse.
- **Changed only where it is safe and visible:** product name, window
  titles, home-screen logo, shell packaging, and shared chrome colours
  (8 files, no engine or provider changes — verified by diff).
- Per-app display names (TechPartner Docs / Sheets / Slides / PDF) and
  new app icons are only to be done if they can be done without
  breaking file associations, builds, or tests. Upstream icons in
  `apps/*/build/` are left alone for now for the same reason.

## Verification so far
- `node tools/check-theme-colors.mjs` passes: "No new raw colors in renderer CSS."
- Pushed to https://github.com/alexwboles/techpartner-office (public,
  fork of genspark-ai/genoffice); first Windows release v0.12.0 built
  by GitHub Actions. The pristine upstream copy is at
  `~/workspace/genoffice-base/`.
