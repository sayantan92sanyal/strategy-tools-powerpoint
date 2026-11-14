# Strategy Tools: notes for Claude Code

PowerPoint **for Mac** add-in (`.ppam`) written in VBA. VBA source is plain text in `src/*.bas`; `scripts/build.py` writes the `.ppam` directly (including `vbaProject.bin` via `scripts/ovba.py`), so never ask the user to open the VBA editor.

## Commands

- Build: `python3 scripts/build.py` → `dist/StrategyTools.ppam` and `dist/Strategy Tools.zip`. Add `--pptm` for the test deck.
- Test: `scripts/run-tests.sh T_Name …` — quits PowerPoint **without saving** (ask the user first if they may have work open), rebuilds, opens `dist/StrategyTools-test.pptm`, and runs each `T_` function from `tests/modTests.bas` via AppleScript. Each prints `PASS` or `FAIL …`.
- Visual check: `T_Render*` tests export slides as PNG to `~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/`; read them to see the result.
- Full suite: the list in README → Development.

## Adding a feature

Follow `docs/adding-a-feature.md`: `Public Sub ST_Feature()` in a module → `Rb_Feature(control As Object)` in `src/modRibbon.bas` → `<button>` in `src/customUI14.xml` (custom icons: add to `scripts/icons.py`) → a `T_Feature` test → README + `src/installer/INSTALL.md` rows.

Conventions: shapes the add-in creates carry `Tags("ST_TYPE")` so they can be found/updated later; selecting an existing visual and re-running its button updates it in place; user-facing messages go through `Msg` (tests set `gQuiet`); use `Confirm` for yes/no.

## VBA rules that cause hangs or silent compile failures on Mac

- The Office type library is **not referenced**: no `msoTrue`, `msoShapeOval`… Use numbers or `ST_` constants in `modCommon.bas`. PowerPoint constants (`ppLayoutBlank`, `ppAlignCenter`…) are fine.
- Module-level `Dim/Public/Const` must come **before** the first procedure.
- **Never loop under `On Error Resume Next`** (an erroring loop header spins forever and PowerPoint hangs). Put the risky call in a small function with `On Error GoTo`.
- **Never call `Application.PathSeparator`** (hangs). Paths use `/`.
- No VBA keywords as names (`circle`, `line`…). A line starting `Foo:` is a label, not a call.
- Use named save-format constants (`ppSaveAsOpenXMLPresentation`); numeric values differ on Mac.
- `ZOrderPosition` counts shapes inside groups; group new shapes by unique names (`GroupShapes` helper), not indexes.

## Test-harness behaviour

- AppleScript error **-18** = PowerPoint isn't ready (still starting, still busy with a timed-out macro, or macros disabled). One timeout cascades into -18 for later calls; `run-tests.sh` waits and retries.
- A compile error shows a VBA dialog the agent can't see; every macro then fails. Bisect modules with `ST_ONLY=modA,modB scripts/run-tests.sh …`, or ask the user for a screenshot.
- Force-quitting PowerPoint makes it ask **Enable Macros** on the next open; the user must click it. Prefer a normal quit.
- Debug hangs with `Trace "step"` (writes `st_trace.log` in the temp folder above during tests).

## Release

Push a tag `vX.Y.Z`; `.github/workflows/release.yml` builds on a macOS runner and attaches `StrategyTools.zip` to the Release.
