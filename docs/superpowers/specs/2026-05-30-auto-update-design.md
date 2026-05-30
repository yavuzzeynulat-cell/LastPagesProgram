# Auto-Update System Design

**Date:** 2026-05-30
**Status:** Approved by user
**Target version:** v0.1.5 (first release with updater)

## Purpose

LastPagesApp.exe currently requires manual download from GitHub releases for every new version. Add an in-app updater so the user is prompted on launch when a newer release exists, and a confirmed update downloads + replaces + restarts automatically.

## User Flow

1. App launches.
2. Background thread (started within ~2s) calls GitHub releases API.
3. If latest tag > `__version__`, show modal: **"Yeni versiyon vX.Y.Z mevcut. Güncellemek ister misin?"** with [Evet] / [Hayır].
4. **Evet** → progress window appears while zip downloads → PowerShell helper launches → app exits → helper replaces .exe → new app starts.
5. **Hayır** → modal closes, current version runs normally. Do not re-prompt during this session.
6. No internet / API error / rate-limited → silently skip, run normally. Errors are logged to stderr only.

## Components

### `__version__` constant in `app.py`
- Single source of truth: `__version__ = "0.1.5"`.
- Bumped before every PyInstaller build + GitHub release.

### `updater.py` (new module)
Three responsibilities, each its own function:
- `get_latest_version() -> str | None` — calls `https://api.github.com/repos/yavuzzeynulat-cell/LastPagesProgram/releases/latest`, returns `tag_name` stripped of leading "v", or `None` on any failure.
- `is_newer(latest: str, current: str) -> bool` — compares two version strings as tuples of ints (e.g. `(0,1,5) > (0,1,4)`). Returns False on parse error.
- `download_and_apply_update(asset_url: str, progress_callback)` — downloads zip to `%TEMP%\lastpages_update\`, extracts, writes `update.ps1`, launches it, returns so caller can exit.

### `update.ps1` (runtime-generated PowerShell script)
Generated at update time, written to `%TEMP%\lastpages_update\update.ps1`:
1. Wait until old `LastPagesApp.exe` process is gone (poll `Get-Process` for ~10s).
2. Copy new exe over the old path.
3. Start the new exe.
4. Delete the temp folder.

Why PowerShell (not .bat): user previously lost ~1.5h to .bat encoding/line-ending issues; PowerShell on Windows 11 is more reliable for file ops and process waits.

### Integration in `app.py`
- In the GUI bootstrap (after `tk.Tk()` is created), spawn a daemon thread that calls `updater.check_for_update(root)`.
- `check_for_update` does the API call off-thread, then on the Tk thread (via `root.after(0, ...)`) shows the modal if needed.
- The download progress window is a small `tk.Toplevel` with a `ttk.Progressbar`.

## Error Handling

| Failure | Behavior |
|---|---|
| No internet | Silent skip, app runs |
| API 403 (rate limit) | Silent skip, app runs |
| Asset download fails mid-way | Show error, app continues with current version |
| Zip extraction fails | Show error, app continues |
| PowerShell helper fails to replace exe | Old exe remains, user sees no app on next launch — mitigation: helper logs to `%TEMP%\lastpages_update\update.log` for diagnosis |
| Defender/SmartScreen blocks new exe | Out of scope for v0.1.5 — user will see SmartScreen warning the first time, accept it (existing pain point, unchanged) |

## Out of Scope

- Rollback on failed launch of new version (user can re-download manually if it breaks).
- Delta updates / partial downloads.
- Beta channel / multiple update tracks.
- Signing the exe (separate problem, not blocking auto-update UX).

## Files Changed / Added

- `app.py` — add `__version__`, add startup updater hook.
- `updater.py` — new file.
- No changes to `write_block()` or any block-writing logic (per longstanding user rule: format DOES NOT change).

## Test Plan

1. Local: set `__version__ = "0.1.4"`, mock `get_latest_version` to return `"0.1.5"` → modal appears.
2. Local: set `__version__ = "0.1.5"`, real API call returns `0.1.4` (current latest) → no modal.
3. Integration: build v0.1.5 .exe, upload to GitHub, then build v0.1.6 .exe with same updater, launch v0.1.5 → confirm full flow downloads + replaces + restarts.
4. Offline: disable network, launch app → no error popup, app works normally.
