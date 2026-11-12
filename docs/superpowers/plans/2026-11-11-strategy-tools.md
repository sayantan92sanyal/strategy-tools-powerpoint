# Strategy Tools Implementation Plan

**Goal:** Ship `dist/StrategyTools.ppam` plus shortcut setup and install docs for Mac PowerPoint.

**Architecture:** VBA source in `src/*.bas` + ribbon XML in `src/customUI14.xml`. A Python build script (`scripts/build.py`, stdlib only) writes an MS-OVBA `vbaProject.bin` (an OLE compound file holding compressed source, no p-code, so PowerPoint recompiles on load) and zips it with a minimal PresentationML add-in package. Keyboard shortcuts: macOS Quick Actions calling `run VB macro`, bound via `pbs` service shortcuts; alignment via `NSUserKeyEquivalents` App Shortcuts.

**Spec:** `docs/design.md`

## Global Constraints
- Mac PowerPoint 16.x; VBA only, no Win32 Declare, no VBProject access.
- Python 3.9 stdlib only for build tooling.
- Ribbon tab "Strategy Tools", groups Format / Visuals / Help.
- Shortcuts: Ctrl+Q text box; Ctrl+Cmd+arrows align; Cmd+Shift+O cheat sheet.

## Review Focus
- Nothing selected / wrong selection count → plain message, no change.
- Grouped shapes or placeholders in selection for Swap/Same Size → still work.
- Title Harmonizer on a deck where some slides have no title placeholder → skip silently.
- Harvey/Meter "update" when selection mixes visuals and other shapes → update only visuals.
- Add-in loaded with no presentation open → message, not a VBA error.

## Tasks
1. **Spike: build pipeline.** `scripts/build.py` produces a `.ppam` containing a test macro; verify by loading in PowerPoint via AppleScript and `run VB macro`. If it fails, fall back to the manual build path (BUILD.md).
2. **Format module** (`src/modFormat.bas`, `src/modCommon.bas`): Swap, SameSize/Width/Height, TitleHarmonizer. Verify on a generated test deck via AppleScript-run macros and inspected shape positions.
3. **Visuals module** (`src/modVisuals.bas`): Harvey balls (5 levels), Meter (3 levels), update-in-place.
4. **Help + shortcuts** (`src/modHelp.bas`, `scripts/setup-shortcuts.sh`): cheat sheet, AddTextBox, Quick Actions, App Shortcuts.
5. **Ribbon** (`src/customUI14.xml`) wired to all macros; verify tab appears.
6. **Docs**: `INSTALL.md`, `docs/adding-a-feature.md`.
