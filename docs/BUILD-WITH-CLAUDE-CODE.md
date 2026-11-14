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

Start an empty folder, run `claude` in it, and paste the plan below. The plan bakes in the decisions and dead-ends this project hit, so Claude Code doesn't have to rediscover them.

**The Features section already lists everything Strategy Tools does**, so pasting the plan as-is rebuilds the full toolbar. Make it yours before pasting: delete the features you don't need, change any detail (colours, sizes, wording, which ribbon group a button sits in), and add your own using the same pattern: *where the button goes, what the user selects first, exactly what appears, and what happens on a second click.*

### The implementation plan (paste this)

```markdown
# Goal
Build a PowerPoint for Mac add-in (.ppam) with a custom ribbon tab for
<audience, e.g. strategy consultants>. Work in phases; after each phase run the
checks listed and show me the result before moving on. Ask me before anything
that changes my Mac's settings, pushes to GitHub or quits PowerPoint while I
might have work open.

# Features
Everything below is the current Strategy Tools feature set. Build all of it
unless I've deleted or changed items. Conventions for every feature:
- One ribbon tab "Strategy Tools" with groups in this order: Format, Visuals,
  Frameworks, Storyline, Review, Help. Menus (▾) show a small picture of what
  each item inserts.
- "Accent colour" = the deck theme's Accent 1, so visuals match any template.
- Everything the add-in creates is grouped where it has several parts and
  tagged (e.g. Tags("ST_TYPE") = "HARVEY") so it can be found and updated.
- "Updates in place": if existing items of that kind are selected when the
  button is used, change them instead of adding a new one.
- Wrong selection or missing input -> a short plain-English message, no change.
- "Body area" = under the slide title (or from 22% down) to 86% of slide
  height, between the title's left edge (or 4%) and the mirrored right edge.

## Format
- Swap: exactly 2 objects selected; they trade places by their centre points,
  each keeping its own size.
- Same Size / Same Width / Same Height: 2+ objects; all match the FIRST object
  clicked (selection order). Top-left corners stay put; aspect-ratio lock is
  switched off while resizing and restored after.
- Title Harmonizer: cursor in (or selection of) a slide title; copy its font
  name and size to the title of every other slide, skipping slides with a
  Title Slide layout and slides without a title. Report how many changed.
- Same Corners: 2+ rounded shapes; all get the first one's corner radius in
  points (so small and large boxes look identical).
- Match Format: 2+ objects; copy fill, outline, shadow, font name, text colour,
  bold and italic from the first clicked. Sizes and positions unchanged.

## Visuals
- Harvey Ball ▾ (0, 25, 50, 75, 100%): light-grey circle with an accent pie
  filled clockwise from 12 o'clock; 0% = grey only, 100% = solid accent;
  0.4 inch, no outline, centred on the slide. Updates in place.
- Meter ▾ (Low, Medium, High): three rising bars (50/75/100% height,
  bottom-aligned), 1/2/3 filled in accent, the rest light grey; 0.4 inch wide.
  Updates in place.
- Status ▾: Complete = green #2E9E5B circle with white tick; In progress =
  amber #F0A202 circle with the tick drawn as white dots; Delayed = red #D64545
  circle with white X; Not started = grey #A3AAB5 circle with white dash.
  Fixed colours (not theme), 0.4 inch. Updating redraws at the same position
  and size.
- Chevrons: ask how many (2-8). A row of chevrons (first one flat-left) in
  accent with white bold "Phase n" labels, spanning the title's width just
  under it, plus a bulleted borderless "Detail" text box under each chevron
  lined up to it. One group.
- Tracker ▾ (Next number, Next letter, Custom...): 0.3 inch accent circle with
  bold white text. "Next" continues from the highest existing tracker of that
  kind on the slide (1,2,3 / A,B,C). Custom = up to 3 characters. Each new one
  goes below the previous of its kind and steps to a free spot so trackers
  never overlap.

## Frameworks
- 2x2 Matrix: fills the body area: four quadrants with a small gap (top-right
  tinted accent, others light greys), editable bold "Quadrant n" titles, axes
  with arrowheads, editable axis labels (vertical one rotated).
- Process Flow: ask how many steps (2-7). Numbered accent circles evenly
  spaced on a grey connecting line, with a bold step name and a grey
  description under each.
- Gantt: ask rows (2-10) and span ("6 months" or "12 weeks"). Header band with
  month names from the current month (or W1, W2...), faint vertical gridlines,
  a label and a staggered accent bar per row to stretch into place, and a
  Milestones row with two amber diamonds. Not linked to dates or Excel.
- Issue Tree: user types an outline in a text box (one item per line, Tab to
  indent, max 4 levels), selects it and clicks. Build a left-to-right tree:
  level 1 accent box with white text, level 2 tinted, deeper levels light grey;
  parents centred on their children; elbow connectors. Replaces the outline.
  With nothing selected, insert an example outline to edit and explain how.
- Key Takeaways: panel on the right of the body area, 27% of slide width:
  accent header "Key takeaways" with white bold text, tinted body with three
  placeholder bullets. One per slide (select the existing one instead).
- Rating ▾ (1-5): five small dots, score filled in accent, rest light grey.
  Updates in place.
- CAGR Arrow: ask start value, end value, years. CAGR = (end/start)^(1/years)-1,
  shown as "+26% p.a." (one decimal below 10%). With two shapes selected, draw
  an elbow arrow from the top of the left one up, across and down to the top
  of the right one, with the label in a rounded white box on the top line;
  otherwise place it mid-slide to drag into place.

## Storyline
- Agenda ▾
  - New agenda: ask how many sections (2-8); add a Title Only slide "Agenda"
    after the current slide with numbered accent boxes and a text box beside
    each for the user to type the section name.
  - Divider here: copy the agenda in front of the current slide with the next
    section highlighted (accent number, bold accent text, tinted band) and the
    others greyed out. Sections are numbered in deck order.
  - Add section tags (optional): small tinted "2 · SECTION NAME" tag in the
    top-right corner of every slide after a divider. A stamp on the same slide
    moves to the left of the tag.
  - Update agenda: after renaming sections on the agenda slide or moving
    slides, rebuild every divider (and tags if used) in place.
- Storyline: add report slide(s) at the end listing every slide's number and
  title, grouped under bold section headers; 16 lines per slide.
- Deck Check: add report slide(s) at the end listing, by slide number:
  non-template fonts, text overflowing its box (autofit off), objects partly
  off the slide, empty placeholders, slides with no title, leftover stamps and
  sticky notes. Show the total in a message; "no issues" if clean.

## Review
- Stamp ▾ (Draft, Preliminary, For discussion, Confidential, Custom...): red
  outlined box with bold red caps text in the top-right corner of the selected
  slides (or current slide); picking another stamp re-labels it. Also: copy
  this slide's stamp to all slides; remove all stamps.
- Sticky Note ▾: yellow note, slightly tilted, centred, signed with the
  author's initials (from the deck's author name), cursor ready to type.
  Remove all sticky notes.
- Source ▾ (Source / Note + Source): small grey "Source:" (and "Note:") text
  at the bottom left, bottom edge fixed at 94% of slide height so it sits in
  exactly the same place on every slide. One per slide.
- Back-up ▾: grey "BACK-UP" tag top-left on the selected slides; offer to
  insert an "Appendix" divider before the first one (not if one is already
  there). Remove back-up label.
- Clean for Sending: deck must be saved. Save "<name> (client)" next to it in
  the same format, then in the copy remove stamps, sticky notes, storyline and
  deck-check slides, speaker notes and comments, and open it. The original is
  never changed.
- Fix Fonts: list fonts that aren't the theme's heading/body fonts (ignore
  symbol fonts), ask, then replace them with the body font everywhere
  including masters and notes.

## Help and keyboard shortcuts
- Text Box button and Ctrl+Q: 3-inch text box centred on the slide, 14 pt,
  word wrap on, grows to fit, cursor inside ready to type.
- Shortcuts button and Shift+Cmd+O: open a searchable cheat sheet page (local
  HTML) of PowerPoint for Mac shortcuts with the add-in's own shortcuts first.
- Ctrl+Cmd+Left/Right/Up/Down: align selected objects left/right/top/bottom
  (to each other; a single object aligns to the slide) via PowerPoint's
  Arrange > Align menu items.

## Add your own
<Delete anything above you don't need, change details, and add features in
the same style: ribbon group, button name, what the user selects first,
exactly what appears (position, size, colours, text), behaviour on re-click.>

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
5. Features, one ribbon group at a time in the order listed above: for each
   feature, mockup -> my OK -> build -> test -> PNG check -> docs.
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
