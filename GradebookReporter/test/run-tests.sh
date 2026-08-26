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

python3 - "$WORK/extracted.js" <<'PY'
import io, sys
src = io.open("ReportScript.gs", encoding="utf-8").read()

def slice_between(start_marker, end_marker):
    start = src.index(start_marker)
    return src[start:src.index(end_marker, start)]

roster = slice_between("// --- ROSTER SCANNING",
                       "/**\n * Scans the sheet and opens the Student Selector Dialog")
selector = slice_between("function buildStudentSelectorHtml",
                         "\n/**\n * Shows the tutorial sidebar.")
encouragement = slice_between("// --- COMPLETION ENCOURAGEMENT ---",
                              "function generateHtmlSummaryStats(rows) {")
io.open(sys.argv[1], "w", encoding="utf-8").write(
    roster + "\n" + selector + "\n" + encouragement + "\n")
PY

cat "$WORK/extracted.js" test/fixtures.js test/assertions.js > "$WORK/suite.js"

if command -v node >/dev/null 2>&1; then
  node "$WORK/suite.js"
else
  # JavaScriptCore treats a top-level run() as the script entry point, so the
  # suite deliberately avoids that name.
  osascript -l JavaScript "$WORK/suite.js"
fi
