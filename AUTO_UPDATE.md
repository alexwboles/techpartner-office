# TechPartner Office — install once, updates come to you

## The promise
Install TechPartner Office **once**. You never delete it and reinstall it
to get our changes. New versions arrive on their own.

## How it works (already wired in `apps/shell/src/main/updater.ts`)
Alex's exact rule: **launch -> check -> download -> apply, on launch.**
- On launch, the app checks for updates after 5 seconds, then every
  4 hours for periodic checks.
- **Launch check:** if a newer version exists, the download starts
  automatically — no click, no delete/reinstall. When the download
  finishes, the app restarts into the new version automatically after
  a 3-second "Updating…" beat. Clicking "Later" cancels that automatic
  restart (the update then falls back to install-on-quit).
- **Periodic checks (every 4 hours):** these do NOT force a restart
  mid-session. A download from a periodic check installs on normal
  quit, or via "Restart & Install" — so work in progress is never
  interrupted by a forced restart hours into a session.
- **Manual check:** Help > Check for Updates still shows the window
  and waits for your click — an explicit check never downloads or
  restarts behind your back.
- **Close guard:** the app's normal unsaved-changes guard still runs
  on quit/restart, so dirty documents are prompted for rather than
  silently lost.
- **Opt-out:** set `TECHPARTNER_NO_AUTO_UPDATE=1` to stop background
  checks and background downloads. Manual checks still work.
- **Safety:** updates are full-package, never differential, and never
  allow a downgrade. After 2 failed download attempts for the same
  version, the window guides you to a manual download instead of
  retrying forever. Documents are not touched by the updater — but
  do not force-quit mid-install; let a normal quit/restart finish it.

## Release workflow (GitHub Actions, free)
No Windows build happens on this Linux VM. The first and every later
release is built by GitHub Actions on a Windows runner:

1. Make the change in this repo, bump the version in
   `apps/shell/package.json` (a release with the same version number
   is ignored by the updater — the version MUST increase every time).
   (The shell workspace is still named `@genoffice/shell` on purpose —
   renaming it would break builds, per Alex's usability rule.)
2. Push a version tag matching that version, e.g. `v0.12.0`.
3. `.github/workflows/release-windows.yml` builds the Windows
   installer with the feed URL baked in
   (`TECHPARTNER_UPDATE_URL=https://github.com/alexwboles/techpartner-office/releases/latest/download`)
   and uploads the installer + `latest.yml` to the GitHub Release for
   that tag. That `latest.yml` + installer is the feed.
4. **Install that first release once on Windows.** From then on,
   every later release works the automatic way above: launch the app,
   it finds the new release, downloads it, and restarts into it.

## What is NOT verified yet (honest caveat)
This updater code is upstream GenOffice code with a small,
test-friendly change (background checks auto-start the download).
It has **not** been verified end-to-end on Windows from this Linux
build VM — no Windows installer was built or run here, and no
TechPartner release exists yet. The first Windows install + the first
automatic update to a second release must be confirmed on Alex's PC
before this is called "live". Do not repeat the Vilano mistake of
declaring auto-update live before a real on-machine update succeeds.
