#!/bin/bash
# Developer test runner: rebuilds, restarts PowerPoint with the test .pptm,
# and runs each macro name given as an argument, printing its return value.
# WARNING: quits PowerPoint without saving. Only use when no real work is open.
set -u
cd "$(dirname "$0")/.."
python3 scripts/build.py --pptm >/dev/null || exit 1
osascript -e 'with timeout of 15 seconds' -e 'tell application "Microsoft PowerPoint" to quit saving no' -e 'end timeout' >/dev/null 2>&1
for _ in $(seq 1 20); do pgrep -q "Microsoft PowerPoint" || break; sleep 1; done
open -a "Microsoft PowerPoint" "dist/StrategyTools-test.pptm"
for _ in $(seq 1 30); do
  sleep 1
  r=$(osascript -e 'with timeout of 5 seconds' -e 'tell application "Microsoft PowerPoint" to get name of every presentation' -e 'end timeout' 2>/dev/null)
  [[ "$r" == *StrategyTools-test* ]] && break
done
# Warm up: wait until PowerPoint will actually run a macro (it reports -18
# or times out for a while after opening a macro file).
for _ in $(seq 1 20); do
  w=$(osascript -e 'with timeout of 10 seconds' -e 'tell application "Microsoft PowerPoint" to run VB macro macro name "T_GetStep"' -e 'end timeout' 2>&1)
  [[ "$w" == *"execution error"* ]] || break
  sleep 3
done
for m in "$@"; do
  # PowerPoint can report error -18 for a few seconds after start-up; retry.
  for attempt in 1 2 3 4; do
    out=$(osascript -e 'with timeout of 120 seconds' -e "tell application \"Microsoft PowerPoint\" to run VB macro macro name \"$m\"" -e 'end timeout' 2>&1)
    [[ "$out" == *"(-18)"* ]] || break
    sleep 4
  done
  echo "$m => $out"
done
