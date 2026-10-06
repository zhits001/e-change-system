# e换系统 (e-Change-System)

**Windows OS Migration Assistant** — before you reinstall Windows or switch to macOS/Linux,
quickly find out **what needs to be carried over**: scan personal files across your disks,
enumerate installed software, and check each app's compatibility with the target system.

> 🤖 **This project was built entirely by an AI coding agent** — design, code, testing, and
> packaging — powered by **GLM-5.3-Flash** (trained by Z.ai). No human-written code.

- **Pure Python standard library** — zero third-party dependencies, fully offline, no telemetry
- Single-file portable app `e换系统.exe` (~11 MB), no installation required
- Targets Windows 10/11; handles OneDrive-redirected user folders automatically

## Quick Start

**Option 1: download the exe** from [GitHub Releases](../../releases) and double-click it.

**Option 2: run from source** (Python 3.8+):

```
python ehuan_system.py
```

**Headless CLI modes:**

```
python ehuan_system.py --report     # full scan, writes the migration report straight to Desktop
python ehuan_system.py --selftest   # verify scan logic and the compatibility knowledge base
```

## Features

### ① Personal Files
Pick scan categories (user folders / WeChat-QQ data / browser bookmarks & passwords / dev
environment / mail & office / deep scan of other disks), then hit **Quick Scan**. Every found
location is listed with its **size, file count, and migration advice**, grouped by category.

- Double-click a row (or press Space) to check/uncheck items you don't plan to migrate;
  the footer keeps a running total ("13 items checked, 3.46 GB")
- Right-click for *Open location* / *Copy path*; export the list as CSV
- While scanning it also collects a **Top large files list (≥200 MB)** so you can cherry-pick

Scan targets include: Desktop / Documents / Downloads / Pictures / Videos / Music,
WeChat & QQ chat-history folders, Chrome / Edge / Firefox / 360 browser profiles,
`.ssh` keys and `.gitconfig` and other dev configs, Outlook PST data, user-installed fonts,
plus large directories on other drives (deep scan).

### ② Software Inventory
Enumerates installed software from the registry **and** Microsoft Store apps, listing name,
version, publisher, size, and install location. A built-in knowledge base (60+ popular apps)
adds three compatibility columns:

| Mark | Meaning |
|------|---------|
| ✅ | Official build available on that system |
| ❌ | No official build; an alternative is needed |
| ➖ | It depends (version/channel specific) |

Each row also gets a **migration tip**, e.g. sign out of Adobe device activation before wiping,
use WeChat's built-in chat migration, sync VS Code / JetBrains via account, or move Steam's
`steamapps` folder to avoid re-downloading games. Unlisted apps get generic advice (record your
license keys and config paths). A filter box narrows the list instantly.

### ③ Migration Report
One click generates a styled **HTML report** on your Desktop (and opens it in the browser):
personal-file checklist + software compatibility table + dev-environment dependency lists
(automatically runs `pip list` / `npm ls -g` etc.) + a **14-point generic migration checklist**
(export browser passwords, save your BitLocker recovery key, confirm OneDrive sync, Windows
digital-license activation, …). Print it and work through it while migrating.

## Output Files

| File | Location |
|------|----------|
| `MigrationReport_YYYYMMDD_HHMM.html` | Desktop |
| Personal-files CSV / software CSV | chosen at export time |

## Privacy & Security

- Everything runs locally — **no network requests**; the report stays on your machine
- Sensitive paths (SSH/GPG keys) are only *listed* for backup; the app never reads or copies
  their contents

## Known Limitations

- "Deep scan other disks" walks whole drives and can be slow on large disks; press **Stop**
  anytime (all scans are cancellable)
- The compatibility knowledge base covers 60+ common apps; unlisted apps are marked as such
- UWP/Store apps show name and version only — Windows does not expose their install size
- Windows only; when restoring on the new system, use the paths recorded in the report

## Build from Source

```
pip install pyinstaller
python -m PyInstaller --onefile --windowed --name "e换系统" --clean --noconfirm ehuan_system.py
```

Output lands in `dist\e换系统.exe`; `e换系统.spec` is the reusable build config.
