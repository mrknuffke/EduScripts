#!/bin/bash
# Runs the roster-scanning tests against ReportScript.gs.
#
# Apps Script can't be executed locally, so this extracts the pure roster
# functions from ReportScript.gs and runs them against fixture sheets that
# mirror real gradebook layouts. Uses node if present, else macOS's built-in
# JavaScriptCore (osascript) so no install is required.
#
# Usage:  ./test/run-tests.sh
set -e
cd "$(dirname "$0")/.."

WORK="$(mktemp -d)"
trap 'rm -rf "$WORK"' EXIT

python3 - "$WORK/roster.js" <<'PY'
import io, sys
src = io.open("ReportScript.gs", encoding="utf-8").read()
start = src.index("// --- ROSTER SCANNING")
end = src.index("/**\n * Scans the sheet and opens the Student Selector Dialog")
io.open(sys.argv[1], "w", encoding="utf-8").write(src[start:end])
PY

cat "$WORK/roster.js" test/fixtures.js test/assertions.js > "$WORK/suite.js"

if command -v node >/dev/null 2>&1; then
  node "$WORK/suite.js"
else
  # JavaScriptCore treats a top-level run() as the script entry point, so the
  # suite deliberately avoids that name.
  osascript -l JavaScript "$WORK/suite.js"
fi
