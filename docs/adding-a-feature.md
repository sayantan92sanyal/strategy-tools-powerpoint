# Adding a feature to Strategy Tools

Everything is plain text in `src/`. The build script turns it into the add-in, so you never need PowerPoint's VBA editor.

## Project layout

| Path | What it is |
|---|---|
| `src/modCommon.bas` | Shared helpers: selection, messages, colours, Office constants |
| `src/modFormat.bas` | Swap, Same Size/Width/Height, Title Harmonizer |
| `src/modVisuals.bas` | Harvey balls, meters |
| `src/modHelp.bas` | Ctrl+Q text box, cheat sheet |
| `src/modRibbon.bas` | One small callback per ribbon button |
| `src/customUI14.xml` | The ribbon tab layout |
| `src/shortcuts.html` | Cheat sheet (`GROUPS` list at the bottom) |
| `src/installer/` | Install/uninstall scripts and INSTALL.md for colleagues |
| `tests/modTests.bas` | In-PowerPoint tests (not shipped) |
| `scripts/build.py` | Builds `dist/StrategyTools.ppam` and `dist/Strategy Tools.zip` |
| `scripts/run-tests.sh` | Restarts PowerPoint and runs tests (quits PowerPoint without saving!) |

## Steps

1. **Write the macro** as `Public Sub ST_YourFeature()` in the module that fits (or a new `src/modSomething.bas` whose first line is `Attribute VB_Name = "modSomething"`). Use `SelectedShapes()` and `Msg` from `modCommon`.
   - The Office type library isn't referenced, so use numbers (or add a `Public Const ST_...` to `modCommon`) instead of `msoTrue`, `msoShapeOval` and similar. PowerPoint constants such as `ppLayoutBlank` are fine.
   - Module-level `Dim`/`Public`/`Const` lines must come before the first Sub/Function.
   - Never put a loop inside `On Error Resume Next`: if the loop header fails, VBA can spin forever and PowerPoint hangs. Wrap the risky call in a small function with `On Error GoTo` instead (see `NotesBody` in `modReview.bas`).
   - Avoid `Application.PathSeparator`: on Mac PowerPoint it hangs. Paths always use `/`.
   - Debugging a hang: call `Trace "step"` (in `modCommon`). During tests it appends to `st_trace.log` in PowerPoint's temp folder (`~/Library/Containers/com.microsoft.Powerpoint/Data/tmp/`).
2. **Add a ribbon callback** in `src/modRibbon.bas`:
   `Public Sub Rb_YourFeature(control As Object): ST_YourFeature: End Sub`
3. **Add a button** in `src/customUI14.xml` inside the right `<group>`:
   `<button id="stYourFeature" label="Your Feature" imageMso="..." onAction="Rb_YourFeature"/>`
   - Use `imageMso="..."` for a built-in Office icon, or `image="stYourIcon"` for a custom one: add a function that draws it in `scripts/icons.py` and register the name in `all_icons()`. The build stops with an error if a ribbon `image=` name has no icon.
4. **(Optional) a keyboard shortcut:** add a `make_workflow` line and a `defaults write pbs` line in `src/installer/Install Strategy Tools.command` (copy the Ctrl+Q ones), and add it to the "Strategy Tools" group in `src/shortcuts.html`.
5. **(Optional) a test:** add a `T_YourFeature()` function in `tests/modTests.bas` that returns "PASS" or "FAIL ...", then run `scripts/run-tests.sh T_YourFeature`.
6. **Build:** `python3 scripts/build.py` and share `dist/Strategy Tools.zip`. Colleagues re-run the installer to update.
