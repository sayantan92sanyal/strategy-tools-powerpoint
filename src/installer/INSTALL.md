# Strategy Tools for PowerPoint (Mac): Install Guide

Strategy Tools adds a **Strategy Tools** tab to PowerPoint with one-click formatting, Harvey balls, meters and keyboard shortcuts.

It takes about 2 minutes.

## You need

- A Mac with Microsoft PowerPoint (Microsoft 365 or Office 2019 or later)
- This **Strategy Tools** folder, unzipped

## Install

1. **Quit PowerPoint** if it's open.
2. In the Strategy Tools folder, **right-click "Install Strategy Tools.command" → Open**.
   - The first time, macOS warns that it's from an unidentified developer. Click **Open**.
     (If there's no Open button, go to System Settings → Privacy & Security, scroll down and click **Open Anyway**.)
3. A Terminal window runs the installer. Answer the prompts as they appear:
   - *"Terminal wants to control Microsoft PowerPoint"* → **OK**
   - PowerPoint asks about access to the file → **Grant Access** / **Select**
   - PowerPoint asks about macros → **Enable Macros**
4. When the window says **Done**, press Return, then **quit PowerPoint (Cmd+Q) and open it again**.
5. You should now see the **Strategy Tools** tab on the ribbon.

The first time you press **Ctrl+Q**, macOS asks whether it can control PowerPoint. Click **OK**. This is a one-off.

> **If the tab doesn't appear:** in PowerPoint choose **Tools → PowerPoint Add-ins…**, click **+**, press **Shift+Cmd+G**, paste
> `~/Library/Application Support/StrategyTools` and choose **StrategyTools.ppam**. Make sure its box is ticked.

## What's included

| Ribbon button / shortcut | What it does |
|---|---|
| **Swap** | Select exactly 2 objects. They trade places. |
| **Same Size / Same Width / Same Height** | Select 2+ objects. All match the **first object you clicked**. |
| **Title Harmonizer** | Click into a slide title. All other slide titles get its font and size (title slides are skipped). |
| **Harvey Ball ▾** | Adds a 0/25/50/75/100% Harvey ball. Select existing balls first to change their level. |
| **Meter ▾** | Adds a Low/Medium/High 3-bar meter. Select existing meters first to change them. |
| **Ctrl+Q** | New text box in the middle of the slide, ready to type. |
| **Ctrl+Cmd+← → ↑ ↓** | Align selected objects left / right / top / bottom (one object aligns to the slide). |
| **Shift+Cmd+O** or **Shortcuts** button | Opens the shortcuts cheat sheet. |

Harvey balls and meters use your template's main accent colour, so they match each client deck.

## Updating

When you get a new version, run **Install Strategy Tools.command** again and restart PowerPoint.

## Uninstall

Run **Uninstall Strategy Tools.command**, then restart PowerPoint.

## Troubleshooting

- **A shortcut does nothing.** Restart PowerPoint. Then check System Settings → Keyboard → Keyboard Shortcuts → **Services** (look for "ST New Text Box" and "ST Shortcuts Cheat Sheet" under General) and **App Shortcuts** (Microsoft PowerPoint).
- **"Macros are disabled".** PowerPoint → Settings → Security & Privacy → Macro Security: choose *Disable all macros with notification*, restart, and click **Enable Macros** when asked.
- **The add-in tab disappeared.** Tools → PowerPoint Add-ins… and tick **StrategyTools**.
