#!/bin/bash
# Removes Strategy Tools from PowerPoint for Mac.
set -u
SUPPORT="$HOME/Library/Application Support/StrategyTools"
PPT_ID="com.microsoft.Powerpoint"

osascript <<OSA >/dev/null 2>&1
with timeout of 60 seconds
tell application "Microsoft PowerPoint"
  repeat with a in add ins
    if name of a is "StrategyTools" then
      set loaded of a to false
      set auto load of a to false
    end if
  end repeat
end tell
end timeout
OSA

rm -rf "$HOME/Library/Services/ST New Text Box.workflow" \
       "$HOME/Library/Services/ST Shortcuts Cheat Sheet.workflow" \
       "$HOME/Library/Application Scripts/$PPT_ID/StrategyTools.scpt" \
       "$SUPPORT"
/System/Library/CoreServices/pbs -update >/dev/null 2>&1 || true

ESC=$(printf '\033')
for item in "Align Left" "Align Right" "Align Top" "Align Bottom"; do
  /usr/libexec/PlistBuddy -c "Delete :NSUserKeyEquivalents:${ESC}Arrange${ESC}Align${ESC}${item}" \
    "$HOME/Library/Containers/$PPT_ID/Data/Library/Preferences/$PPT_ID.plist" >/dev/null 2>&1 || true
done

echo "Strategy Tools removed. If 'StrategyTools' still shows under Tools > PowerPoint Add-ins,"
echo "select it there and click the minus (-) button. Restart PowerPoint to finish."
read -r -p "Press Return to close this window." _
