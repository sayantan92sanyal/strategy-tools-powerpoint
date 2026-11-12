#!/bin/bash
# Strategy Tools installer for PowerPoint for Mac.
# Double-click to run (first time: right-click > Open). Safe to run again to update.
set -u
HERE="$(cd "$(dirname "$0")" && pwd)"
SUPPORT="$HOME/Library/Application Support/StrategyTools"
SCRIPTS="$HOME/Library/Application Scripts/com.microsoft.Powerpoint"
SERVICES="$HOME/Library/Services"
PPT_ID="com.microsoft.Powerpoint"

say() { printf '\n\033[1m%s\033[0m\n' "$1"; }
fail() { printf '\n\033[31mProblem: %s\033[0m\n' "$1"; exit 1; }

[ -f "$HERE/StrategyTools.ppam" ] || fail "StrategyTools.ppam must be in the same folder as this installer."
[ -d "/Applications/Microsoft PowerPoint.app" ] || fail "Microsoft PowerPoint is not installed in /Applications."

say "1/5  Copying Strategy Tools files"
mkdir -p "$SUPPORT"
cp "$HERE/StrategyTools.ppam" "$SUPPORT/StrategyTools.ppam"
cp "$HERE/shortcuts.html" "$SUPPORT/shortcuts.html"
echo "     -> $SUPPORT"

say "2/5  Installing the cheat-sheet helper"
mkdir -p "$SCRIPTS"
osacompile -o "$SCRIPTS/StrategyTools.scpt" \
  -e 'on openCheatSheet(arg)' \
  -e '  do shell script "open " & quoted form of (POSIX path of (path to home folder) & "Library/Application Support/StrategyTools/shortcuts.html")' \
  -e 'end openCheatSheet' || fail "could not create the cheat-sheet helper."

say "3/5  Creating keyboard-shortcut actions"
make_workflow() {   # $1 = name, $2 = AppleScript body (one line)
  local wf="$SERVICES/$1.workflow"
  rm -rf "$wf"
  mkdir -p "$wf/Contents"
  cat > "$wf/Contents/Info.plist" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>NSServices</key>
	<array>
		<dict>
			<key>NSMenuItem</key>
			<dict><key>default</key><string>$1</string></dict>
			<key>NSMessage</key>
			<string>runWorkflowAsService</string>
			<key>NSRequiredContext</key>
			<dict><key>NSApplicationIdentifier</key><string>$PPT_ID</string></dict>
		</dict>
	</array>
</dict>
</plist>
PLIST
  cat > "$wf/Contents/document.wflow" <<PLIST
<?xml version="1.0" encoding="UTF-8"?>
<!DOCTYPE plist PUBLIC "-//Apple//DTD PLIST 1.0//EN" "http://www.apple.com/DTDs/PropertyList-1.0.dtd">
<plist version="1.0">
<dict>
	<key>AMApplicationBuild</key><string>523</string>
	<key>AMApplicationVersion</key><string>2.10</string>
	<key>AMDocumentVersion</key><string>2</string>
	<key>actions</key>
	<array>
		<dict>
			<key>action</key>
			<dict>
				<key>AMAccepts</key>
				<dict>
					<key>Container</key><string>List</string>
					<key>Optional</key><true/>
					<key>Types</key><array><string>com.apple.applescript.object</string></array>
				</dict>
				<key>AMActionVersion</key><string>1.0.2</string>
				<key>AMApplication</key><array><string>Automator</string></array>
				<key>AMParameterProperties</key><dict><key>source</key><dict/></dict>
				<key>AMProvides</key>
				<dict>
					<key>Container</key><string>List</string>
					<key>Types</key><array><string>com.apple.applescript.object</string></array>
				</dict>
				<key>ActionBundlePath</key><string>/System/Library/Automator/Run AppleScript.action</string>
				<key>ActionName</key><string>Run AppleScript</string>
				<key>ActionParameters</key>
				<dict>
					<key>source</key>
					<string>on run {input, parameters}
	$2
	return input
end run</string>
				</dict>
				<key>BundleIdentifier</key><string>com.apple.Automator.RunScript</string>
				<key>CFBundleVersion</key><string>1.0.2</string>
				<key>CanShowSelectedItemsWhenRun</key><false/>
				<key>CanShowWhenRun</key><true/>
				<key>Category</key><array><string>AMCategoryUtilities</string></array>
				<key>Class Name</key><string>RunScriptAction</string>
				<key>InputUUID</key><string>6C1D4C1A-1B0E-4E0B-9C61-0A1F1E7D0001</string>
				<key>Keywords</key><array><string>Run</string></array>
				<key>OutputUUID</key><string>6C1D4C1A-1B0E-4E0B-9C61-0A1F1E7D0002</string>
				<key>UUID</key><string>6C1D4C1A-1B0E-4E0B-9C61-0A1F1E7D0003</string>
				<key>UnlocalizedApplications</key><array><string>Automator</string></array>
				<key>arguments</key>
				<dict>
					<key>0</key>
					<dict>
						<key>default value</key><string>on run {input, parameters}
	return input
end run</string>
						<key>name</key><string>source</string>
						<key>required</key><string>0</string>
						<key>type</key><string>0</string>
						<key>uuid</key><string>0</string>
					</dict>
				</dict>
				<key>isViewVisible</key><true/>
				<key>location</key><string>309.5:253.0</string>
			</dict>
			<key>isViewVisible</key><true/>
		</dict>
	</array>
	<key>connectors</key><dict/>
	<key>workflowMetaData</key>
	<dict>
		<key>applicationBundleID</key><string>$PPT_ID</string>
		<key>applicationBundleIDsByPath</key>
		<dict><key>/Applications/Microsoft PowerPoint.app</key><string>$PPT_ID</string></dict>
		<key>applicationPath</key><string>/Applications/Microsoft PowerPoint.app</string>
		<key>applicationPaths</key><array><string>/Applications/Microsoft PowerPoint.app</string></array>
		<key>inputTypeIdentifier</key><string>com.apple.Automator.nothing</string>
		<key>outputTypeIdentifier</key><string>com.apple.Automator.nothing</string>
		<key>presentationMode</key><integer>11</integer>
		<key>processesInput</key><integer>0</integer>
		<key>serviceApplicationBundleID</key><string>$PPT_ID</string>
		<key>serviceApplicationPath</key><string>/Applications/Microsoft PowerPoint.app</string>
		<key>serviceInputTypeIdentifier</key><string>com.apple.Automator.nothing</string>
		<key>serviceOutputTypeIdentifier</key><string>com.apple.Automator.nothing</string>
		<key>serviceProcessesInput</key><integer>0</integer>
		<key>systemImageName</key><string>NSActionTemplate</string>
		<key>useAutomaticInputType</key><integer>0</integer>
		<key>workflowTypeIdentifier</key><string>com.apple.Automator.servicesMenu</string>
	</dict>
</dict>
</plist>
PLIST
}
mkdir -p "$SERVICES"
make_workflow "ST New Text Box" 'tell application "Microsoft PowerPoint" to run VB macro macro name "ST_AddTextBox"'
make_workflow "ST Shortcuts Cheat Sheet" 'do shell script "open " &amp; quoted form of (POSIX path of (path to home folder) &amp; "Library/Application Support/StrategyTools/shortcuts.html")'

say "4/5  Assigning keyboard shortcuts"
# Quick Action shortcuts: Ctrl+Q and Shift+Cmd+O
defaults write pbs NSServicesStatus -dict-add \
  '"(null) - ST New Text Box - runWorkflowAsService"' \
  '{ "enabled_context_menu" = 1; "enabled_services_menu" = 1; "key_equivalent" = "^q"; }'
defaults write pbs NSServicesStatus -dict-add \
  '"(null) - ST Shortcuts Cheat Sheet - runWorkflowAsService"' \
  '{ "enabled_context_menu" = 1; "enabled_services_menu" = 1; "key_equivalent" = "@$o"; }'
/System/Library/CoreServices/pbs -update >/dev/null 2>&1 || true

# Alignment: Ctrl+Cmd+arrows bound to PowerPoint's Arrange > Align menu items.
ESC=$(printf '\033')
LEFT=$(printf '\357\234\202'); RIGHT=$(printf '\357\234\203')
UP=$(printf '\357\234\200');   DOWN=$(printf '\357\234\201')
for pair in "Align Left:$LEFT" "Align Right:$RIGHT" "Align Top:$UP" "Align Bottom:$DOWN"; do
  item="${pair%%:*}"; key="${pair#*:}"
  defaults write "$PPT_ID" NSUserKeyEquivalents -dict-add "${ESC}Arrange${ESC}Align${ESC}${item}" "@^${key}"
done
echo "     Ctrl+Q            new text box"
echo "     Ctrl+Cmd+arrows   align left / right / top / bottom"
echo "     Shift+Cmd+O       shortcuts cheat sheet"

say "5/5  Loading the add-in into PowerPoint"
echo "     macOS may ask whether Terminal can control PowerPoint: click OK."
echo "     PowerPoint may ask about macros: click Enable Macros."
if osascript <<OSA
with timeout of 180 seconds
tell application "Microsoft PowerPoint"
  activate
  set a to register add in "$SUPPORT/StrategyTools.ppam"
  set loaded of a to true
  set auto load of a to true
end tell
end timeout
OSA
then
  echo "     Add-in loaded."
else
  echo "     Couldn't load it automatically. In PowerPoint choose Tools > PowerPoint Add-ins…,"
  echo "     click +, press Shift+Cmd+G, paste the folder below and pick StrategyTools.ppam:"
  echo "     $SUPPORT"
fi

say "Done!  Quit PowerPoint (Cmd+Q) and open it again so the keyboard shortcuts take effect."
echo "     The first time you press Ctrl+Q, macOS asks to let it control PowerPoint: click OK."
echo
read -r -p "Press Return to close this window." _
