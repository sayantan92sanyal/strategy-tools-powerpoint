# Build it with Claude Code

This guide is for people who want to use [Claude Code](https://claude.com/claude-code) to either:

- **Path A: add your own features to Strategy Tools** (your firm's slide formats, a new framework, a different shortcut), or
- **Path B: build your own PowerPoint add-in from scratch**, using the approach that made this one work on Mac.

You don't need to be a developer. You describe what you want; Claude Code writes the code, builds the add-in, runs it inside your PowerPoint and checks the result. Your job is to answer its questions, click **Enable Macros** when PowerPoint asks, and send a screenshot when it can't see what's on your screen.

## What you need

- A Mac with **PowerPoint for Mac** (Microsoft 365 or Office 2019+), with macros allowed (not a locked-down work laptop).
- **Claude Code** installed (terminal, desktop app or VS Code extension).
- **Python 3**: type `python3 --version` in Terminal. If macOS offers to install developer tools, accept.
- Optional: a free **GitHub** account to keep your own copy online.

Close any PowerPoint decks you care about before Claude Code runs tests: the test runner quits PowerPoint without saving.

---

## Path A: add features to Strategy Tools

1. **Copy the project.** On GitHub, click **Fork** to get your own copy (or **Code → Download ZIP** if you don't use GitHub). Then in Terminal:
   ```bash
   git clone https://github.com/<you>/strategy-tools-powerpoint.git
   cd strategy-tools-powerpoint
   claude
   ```
   Claude Code automatically reads `CLAUDE.md`, which holds the build/test commands and the Mac pitfalls listed below.

2. **Describe the feature.** Paste something like this and edit the middle part:

   > I want to add a feature to this PowerPoint add-in. Read CLAUDE.md and docs/adding-a-feature.md first.
   >
   > **Feature:** a "Disclaimer" button in the Review group that inserts our standard disclaimer slide at the end of the deck: title "Important notice", the text below in 10 pt grey, and our logo bottom-right from `assets/logo.png`.
   > *(text…)*
   >
   > Before building, show me how it will look (a mockup) and ask me anything that's unclear. Then build it, add a test, run the test suite, export a PNG of the result so you can check it, and update README and INSTALL.md. Don't push anything without asking.

3. **Install your version.** When Claude says it's done: quit PowerPoint, run `dist/Strategy Tools/Install Strategy Tools.command`, reopen PowerPoint.

4. **Share it.** Send colleagues `dist/Strategy Tools.zip`, or push a version tag to your fork (`git tag v1.0.0 && git push origin v1.0.0`) and GitHub builds a downloadable release for you.

**Good feature requests** say *where* the button goes (which ribbon group), *what is selected* when you click it, *exactly what appears* (position, size, colours, text), and *what happens on a second click* (update in place? add another?). Ask for a mockup first; it's much cheaper to change a picture than code.

---

## Path B: build your own add-in from scratch

Start an empty folder, run `claude` in it, and paste the plan below (edit **Your features** first). The plan bakes in the decisions and dead-ends this project hit, so Claude Code doesn't have to rediscover them.

### The implementation plan (paste this)

```markdown
# Goal
Build a PowerPoint for Mac add-in (.ppam) with a custom ribbon tab for
<audience, e.g. strategy consultants>. Work in phases; after each phase run the
checks listed and show me the result before moving on. Ask me before anything
that changes my Mac's settings, pushes to GitHub or quits PowerPoint while I
might have work open.

# Your features
<List each feature: ribbon group, button name, what the user selects first,
exactly what appears (position, size, colours, text), behaviour on re-click.
Example: "Visuals > Harvey Ball menu (0/25/50/75/100%): grouped pie + circle,
0.4 inch, Accent 1 colour on light grey, centred on the slide; if Harvey balls
are selected, change their level instead of adding one.">

# Architecture (decided - don't change without asking)
- VBA in a macro-enabled add-in (.ppam). Office.js web add-ins can't register
  keyboard shortcuts in PowerPoint for Mac and can't reliably tell selection
  order ("first object clicked"), so VBA it is.
- Source of truth = plain text: src/*.bas (one module per area + modCommon for
  helpers), src/customUI14.xml (ribbon), src/base.pptx (empty deck created by
  PowerPoint via AppleScript, author metadata scrubbed).
- scripts/build.py (Python stdlib only) assembles the .ppam directly:
  * writes vbaProject.bin itself: an OLE Compound File [MS-CFB] containing
    PROJECT, PROJECTwm, VBA/_VBA_PROJECT (7-byte header, version 0xFFFF so
    Office recompiles), VBA/dir and one stream per module, all using
    [MS-OVBA] compression. No p-code: PowerPoint compiles on first load.
  * zips base.pptx + vbaProject.bin + customUI/customUI14.xml (+ icon PNGs and
    customUI/_rels) with the add-in content type
    application/vnd.ms-powerpoint.addin.macroEnabled.main+xml.
  * optional --pptm: same project plus tests/*.bas in a .pptm test deck with a
    different VBA project name (two open projects with the same name confuse
    PowerPoint).
- Ribbon callbacks take (control As Object), not IRibbonControl (no Office
  type library reference). Menus pass a value through the button's tag.
- Icons: draw PNGs in Python at build time (stdlib rasteriser with
  supersampling) and reference them with image="name"; the build fails if a
  referenced icon is missing.
- Keyboard shortcuts (add-ins can't register them on Mac):
  * shortcuts that run a macro = macOS Quick Action (~/Library/Services/X.workflow,
    "Run AppleScript": tell PowerPoint to run VB macro "Name"), scoped to
    PowerPoint, key bound via `defaults write pbs NSServicesStatus`.
  * shortcuts for built-in commands = App Shortcuts:
    `defaults write com.microsoft.Powerpoint NSUserKeyEquivalents -dict-add
    "<ESC>Arrange<ESC>Align<ESC>Align Left" "@^<U+F702>"` (menu path with ESC
    separators; arrow keys are U+F700-F703; ^ ctrl, @ cmd, $ shift, ~ option).
  * anything VBA needs to open outside PowerPoint's sandbox goes through
    AppleScriptTask with a script in
    ~/Library/Application Scripts/com.microsoft.Powerpoint/.
- Installer: a double-clickable .command (bash, no Python) that copies files to
  ~/Library/Application Support/<Name>, writes the Quick Actions and shortcuts,
  and loads the add-in with AppleScript `register add in` (switching off any
  other copy with the same name). Plus an uninstaller and INSTALL.md.
- Release: GitHub Actions on version tags, macos-latest runner, runs the build
  and attaches the zip to a Release.

# Phases and checks
1. Pipeline spike: build a .pptm with ONE trivial function, open it in
   PowerPoint, call it with `osascript -e 'tell application "Microsoft
   PowerPoint" to run VB macro macro name "Hello"'`. Check: it returns the
   value. Then TWO modules (multi-sector streams expose CFB bugs). Validate the
   file with oletools' olevba in a throwaway venv as well.
2. Ribbon: customUI14.xml with one button calling a callback; load the .ppam
   with AppleScript `register add in`; ask me to confirm the tab appears.
3. Shared helpers (modCommon): selection in click order
   (ActiveWindow.Selection.ShapeRange), TargetSlides (thumbnail selection or
   current slide), Msg/Confirm that tests can silence, tagging helpers,
   GroupShapes by unique names, content area under the title.
4. Test harness: tests/modTests.bas with T_ functions returning "PASS"/"FAIL
   reason"; scripts/run-tests.sh rebuilds, quits PowerPoint normally, opens the
   test deck, waits until a trivial macro answers, then runs each test with
   retries on error -18. Render tests export slides to PNG (in PowerPoint's
   container tmp folder) so you can look at the result.
5. Features, one group at a time: for each, mockup -> my OK -> build -> test
   -> PNG check -> docs.
6. Shortcuts + installer + uninstaller; ask me to run the installer and check
   each shortcut.
7. README, INSTALL.md, LICENSE, release workflow.

# Pitfalls to avoid (each one cost hours)
- CFB writer: sibling directory entries form a red-black tree; colour only the
  deepest level red in a midpoint-built tree. FAT chains: sector i points to
  i+1 (easy off-by-one that only shows with multi-sector streams).
- No Office type library: msoTrue etc. are undefined -> use numeric constants.
  Module-level declarations must precede procedures. A compile error shows an
  invisible-to-you dialog and every macro then fails with -18.
- Never loop under On Error Resume Next (infinite loop, PowerPoint hangs).
- Application.PathSeparator hangs on Mac.
- Don't use VBA keywords as names (circle, line); "Foo:" at line start is a
  label.
- Save-format numbers differ on Mac (24 exported PNGs); use named constants.
- ZOrderPosition counts shapes inside groups; never index Shapes.Range by it.
- New text boxes shrink-to-fit on creation; set AutoSize none, then set
  Top/Height again.
- Force-quitting PowerPoint triggers an Enable Macros prompt on next open; the
  user must click it. Use a normal AppleScript quit.
- You cannot see the screen. When PowerPoint stops answering, ask the user for
  a screenshot instead of guessing.
```

### Tips while it builds

- **Expect a few rounds of "what does PowerPoint show right now?"** Claude Code can't see your screen; a screenshot of the dialog is usually all it needs.
- **Say yes to mockups.** Claude Code can open a browser preview of how a slide will look before writing any code.
- **Keep features small.** One ribbon group at a time, tested, then the next.
- **Ask for a test per feature.** The test suite is what lets you add feature 20 without breaking feature 3.

---

## How Strategy Tools itself was built

The whole project, from the first idea to a public release, was built in a few sessions with Claude Code using exactly this process: brainstorm, then a mockup, then build, test inside PowerPoint and check the rendered PNG, then release. `docs/design.md` and `docs/superpowers/plans/` show the design decisions as they were made.
