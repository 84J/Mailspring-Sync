# Mailspring-Sync — 84J patches

This fork of [Foundry376/Mailspring-Sync](https://github.com/Foundry376/Mailspring-Sync)
(the Mailspring sync engine) carries a few small patches on the `84j` branch and builds the
engine for Debian 13 (amd64) for each official Mailspring release.

## Patches

Commits on `84j` whose subject starts with `[84j] ` are patches:

| Patch | What it does |
|---|---|
| Unthreaded mode | With `MAILSYNC_UNTHREADED=1`, every message becomes its own thread, so Mailspring shows a plain chronological message list. Existing accounts need *Preferences → Accounts → Rebuild Cache*. |
| Microsoft 365 contacts *(optional)* | Syncs Outlook.com / Microsoft 365 contacts through Microsoft Graph, including edits and deletes from Mailspring. Needs an engine version with Mailspring's Graph code (after 1.26.2). |

Commits starting with `[84j-ci]` (workflow, this README) are not patches.

## Builds

`.github/workflows/84j-release.yml` runs daily. When Mailspring publishes a release, it checks
out the engine commit that release uses, cherry-picks the patches, builds on Debian 13 and
publishes release `mailspring-<version>` with the engine tarball, checksum, runtime
dependencies and build info.

- A patch with the trailer `84j-Optional: true` is skipped when it doesn't apply or doesn't
  build on that engine version.
- If a required patch fails, the run opens an issue and publishes nothing; the previous
  release stays available.
- Run it manually under *Actions → 84j engine release → Run workflow* (optionally for a given
  Mailspring version, or with *force* to rebuild).

[debiansetup](https://github.com/84J/debiansetup) installs the official Mailspring `.deb` and
replaces its engine with the matching release from here.

## Adding or updating a patch

Keep patches small and in new files where possible, so upstream changes rarely conflict.
Commit with a `[84j] ` subject on `84j`, rebase `84j` onto `upstream/master` from time to
time, and push — the push triggers a rebuild of the latest release.
