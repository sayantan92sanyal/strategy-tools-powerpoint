# Strategy Tools: PowerPoint Add-in for Mac (Design)

Date: 2026-11-11

## Intent

A PowerPoint add-in with the one-click helpers that consulting firms build into their in-house PowerPoint toolkits: quick formatting, standard consulting visuals and keyboard shortcuts. It's for strategy work on Mac. The user installs it first and then shares it with colleagues, who are also on Mac. New features will keep being added, so it must be easy to extend.

## Goals

- A **Strategy Tools** ribbon tab with three groups:
  - **Format**: Swap, Same Size, Same Width, Same Height, Title Harmonizer
  - **Visuals**: Harvey Ball (dropdown), Meter (dropdown)
  - **Help**: Shortcuts cheat sheet
- Keyboard shortcuts for adding a text box, aligning objects and opening the cheat sheet.
- Colleagues install it by following a short guide written for non-developers.
- Adding a feature later means adding one macro and one ribbon entry, nothing more.

## Non-goals

- Windows support (may come later; the code shouldn't block it, but it isn't tested).
- PowerPoint on the web or iPad.
- A web add-in, hosting, accounts or any network access.

## Anti-goals (would count as failure even if it technically works)

- Breaking normal PowerPoint behaviour: text editing, undo, or built-in shortcuts the user relies on.
- An install process colleagues can't follow without help.
- Errors that show raw VBA dialogs. Every misuse (wrong selection, and so on) gets a plain one-line message.

## Constraints

- **Platform:** a macro-enabled PowerPoint add-in (`.ppam`) with custom ribbon XML (`customUI/customUI14.xml`). The user's Mac allows macros. Use VBA only, with no Windows-only APIs (no Win32 `Declare` calls, no `VBProject` access).
- **Keyboard shortcuts:** PowerPoint for Mac doesn't let add-ins register keyboard shortcuts. Use macOS App Shortcuts (System Settings → Keyboard → Keyboard Shortcuts → App Shortcuts), which bind keys to PowerPoint **menu bar** item titles. A one-time setup script writes these with `defaults write com.microsoft.Powerpoint NSUserKeyEquivalents` so nobody types them by hand.
- **Building the `.ppam`:** VBA can't be compiled outside PowerPoint, and Mac PowerPoint can't import modules from a script. The source lives as plain text files (`.bas` modules and ribbon XML) in the repo. Building takes a one-time manual step (import the modules in the VBA editor and save as `.ppam`), then a script adds the ribbon XML into the saved file. Colleagues only ever receive the finished `.ppam`.

## Feature behaviour (as agreed)

### Format group

| Button | Selection | Behaviour |
|---|---|---|
| Swap | exactly 2 objects | The two objects trade places by their **center points**. Each keeps its own size. |
| Same Size | 2+ objects | Every object gets the width and height of the **first one clicked** (`Selection.ShapeRange(1)`). The top-left corner stays fixed. Aspect-ratio lock is temporarily switched off so pictures can stretch. |
| Same Width | 2+ objects | As above, width only. |
| Same Height | 2+ objects | As above, height only ("same length" in the request means height). |
| Title Harmonizer | cursor in, or selection of, a slide title | Copies that title's **font name and font size** onto the title placeholder of every other slide. Color, bold and position aren't touched. Slides whose layout is a *Title Slide* layout are skipped. A message reports how many slides changed. |

A wrong selection shows a short message (for example "Select exactly 2 objects to swap") and changes nothing.

### Visuals group

**Harvey Ball** dropdown: 0%, 25%, 50%, 75%, 100%.
- Style: no outline, a light grey base circle (≈ #E3E6EA), with the filled part in the deck's **Accent 1** theme colour. 100% is a solid accent circle and 0% is grey only.
- Size: about 0.4" × 0.4", placed at the slide center. It's a grouped shape named like `HarveyBall_50` so it can be found later.
- If one or more existing Harvey balls are selected, picking a level **updates those balls** instead of adding a new one.

**Meter** dropdown: Low, Medium, High.
- Style: three **rising bars** (short, medium, tall, bottom-aligned, like a phone signal icon). Filled bars use Accent 1 and empty ones are light grey. Low fills 1 bar, Medium 2, High 3.
- Size: about 0.4" wide overall. Grouped and named like `Meter_Medium`.
- Selecting existing meters and picking a level updates them.

### Keyboard shortcuts

| Keys | Action | Mechanism |
|---|---|---|
| Ctrl + Q | New text box: a ready-made box dropped in the **middle of the current slide**, already selected and in typing mode | macOS Quick Action that runs the `AddTextBox` macro (same mechanism as the cheat sheet, below). The user chose this over the instant menu route even though it takes about a second. |
| Ctrl + Cmd + ← | Align left | App Shortcut on Arrange → Align → Align Left |
| Ctrl + Cmd + → | Align right | App Shortcut on Align Right |
| Ctrl + Cmd + ↑ | Align top | App Shortcut on Align Top |
| Ctrl + Cmd + ↓ | Align bottom | App Shortcut on Align Bottom |
| Cmd + Shift + O | Shortcuts cheat sheet | See below |

Alignment uses PowerPoint's own behaviour: with 2+ objects they align to each other, and with 1 object it aligns to the slide.

**Cheat sheet:** a window inside PowerPoint listing common Mac PowerPoint shortcuts by category (File, Editing, Formatting, Slides, View, Objects), plus a **Strategy Tools** section listing the custom shortcuts. It opens from the Help ribbon button. For **Cmd + Shift + O**, the add-in's macros don't appear in the menu bar, so the shortcut is a macOS Quick Action (Automator) that runs `tell application "Microsoft PowerPoint" to run VB macro macro name "ShowShortcuts"`, bound to Cmd + Shift + O under Keyboard Shortcuts → Services. If PowerPoint already uses Cmd + Shift + O, the builder picks the nearest free combo and documents it.

## Extensibility

- One VBA module per area (`modFormat`, `modVisuals`, `modHelp`) and a shared `modCommon` (selection checks, user messages, theme colour lookup).
- Each ribbon button calls one public macro with the standard `(control As IRibbonControl)` signature.
- The cheat sheet's custom-shortcut list lives in one place in the code, so adding a shortcut means adding one line.
- `docs/adding-a-feature.md` covers the steps: write the macro, add the ribbon button XML, rebuild, re-share.

## Deliverables

1. `src/`: the `.bas` modules, any form files, and `customUI14.xml` with icons (built-in Office icons via `imageMso` where possible).
2. `scripts/`: the script that adds the ribbon into the `.ppam`, and the keyboard-shortcut setup script (App Shortcuts, plus installing the Quick Action).
3. `dist/StrategyTools.ppam`: the shareable add-in, produced after the user's one-time build step.
4. `INSTALL.md`: install steps for colleagues (copy the file, add it via Tools → PowerPoint Add-ins, enable macros, run the shortcut setup, grant Automation permission).
5. `BUILD.md`: the one-time build steps for the user, written so a non-developer can follow them.

## Builder's calls (not discussed with the user)

- The exact menu item titles in Mac PowerPoint get checked on the installed version. App Shortcuts need an exact title match.
- The default text box: about 3" wide, using the theme's body font at 14 pt, with auto-fit switched on.
- The cheat sheet is a VBA UserForm with a scrollable read-only text area. If UserForms prove unreliable on Mac, fall back to opening a bundled HTML page.
- Each button is a single undo step where PowerPoint allows it.

## Additions (2026-11-12): status icons, chevrons, trackers

All three go in the **Visuals** group.

- **Status ▾**: Complete (green #2E9E5B circle, white tick), In progress (amber #F0A202 circle, tick drawn as white dots), Delayed (red #D64545 circle, white X), Not started (grey #A3AAB5 circle, white dash). The user chose four states over three. Fixed colours, not theme colours. About 0.4". Selecting existing icons and picking a status redraws them in place (same position and size).
- **Chevrons**: an input box asks for 2–8. One click gives a row of chevrons plus a bulleted, borderless detail text box under each (the user picked this over chevrons only). *(Builder's call)* The first shape is a flat-left pentagon. The row lines up with the title placeholder's left edge and width when there is one, otherwise it uses 6% margins, and it sits just under the title. Labels read "Phase 1…n" in white bold 14 pt on Accent 1. Everything is one group tagged CHEVRONS.
- **Tracker ▾**: Next number, Next letter, Custom… (up to 3 characters). Each tracker is a 0.3" Accent 1 circle with bold white text. "Next" uses the highest existing tracker of that kind on the slide. *(Builder's call)* Each new tracker is placed below the previous one of its kind and steps to a free spot, so trackers never overlap.
