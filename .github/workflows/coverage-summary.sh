#!/usr/bin/env bash
#
# Turns the Cobertura XML that "dotnet test --collect:XPlat Code Coverage" writes into a
# human-readable job summary, and prints the same figures to the log.
#
# WHY THIS EXISTS. .claude/CLAUDE.md deliberately refuses to record a coverage percentage in any
# document, because a figure written down goes stale silently and is then quoted as fact. That is
# exactly what happened: the working assumption was "about 25%" when the real overall figure was
# 40.2% and 25% was the UI-only number. A figure the pipeline recomputes on every push cannot go
# stale, which is the point of doing it here rather than in a README.
#
# It deliberately does NOT fail the build on a threshold. Coverage is a trend to watch, not a gate,
# and a floor set while the suite is still growing either blocks unrelated work or is set so low it
# means nothing. If a gate is wanted later, that is a separate and deliberate decision.
#
# Usage: coverage-summary.sh <path-to-coverage.cobertura.xml>
#
# Hand it the MERGED report (build-and-unittest.yml merges the four test projects' reports with
# ReportGenerator first). One project's own report is that project's figure, not the total.

set -euo pipefail

REPORT="${1:-}"

if [ -z "$REPORT" ] || [ ! -f "$REPORT" ]; then
    echo "No coverage report found at '${REPORT}' - skipping the coverage summary."
    exit 0
fi

# The totals live in attributes on the root <coverage> element, so a grep is enough and this
# needs no XML tooling on the runner.
#
# The attributes are read from THAT ELEMENT ONLY, not from the file at large. Cobertura writers
# put line-rate/branch-rate on <package> and <class> too, and coverlet has emitted
# lines-covered/lines-valid on <package> as well - so a whole-file "grep | head -1" is right only
# as long as the root element happens to come first in document order. If it ever does not (a
# merged multi-assembly report, a reordered writer), that grep would silently report ONE PACKAGE's
# numbers as the project total: no error, just a wrong figure quoted as fact, which is the exact
# failure this script exists to prevent. So isolate the root element once, then read from it.
COVERAGE_ELEMENT=$(grep -m1 -o '<coverage [^>]*>' "$REPORT" || true)

if [ -z "$COVERAGE_ELEMENT" ]; then
    echo "Coverage report at '${REPORT}' has no root <coverage> element - skipping the summary."
    exit 0
fi

read_attr() {
    printf '%s' "$COVERAGE_ELEMENT" | grep -o "$1=\"[0-9.]*\"" | head -1 | grep -o '[0-9.]*'
}

LINES_COVERED=$(read_attr 'lines-covered')
LINES_VALID=$(read_attr 'lines-valid')
BRANCHES_COVERED=$(read_attr 'branches-covered')
BRANCHES_VALID=$(read_attr 'branches-valid')

if [ -z "$LINES_VALID" ] || [ "$LINES_VALID" = "0" ]; then
    echo "Coverage report at '${REPORT}' carried no line totals - skipping the summary."
    exit 0
fi

percent() {
    awk -v covered="$1" -v valid="$2" 'BEGIN { if (valid == 0) print "n/a"; else printf "%.1f", (covered / valid) * 100 }'
}

LINE_PERCENT=$(percent "$LINES_COVERED" "$LINES_VALID")
BRANCH_PERCENT=$(percent "$BRANCHES_COVERED" "$BRANCHES_VALID")

echo "Line coverage:   ${LINE_PERCENT}% (${LINES_COVERED}/${LINES_VALID})"
echo "Branch coverage: ${BRANCH_PERCENT}% (${BRANCHES_COVERED}/${BRANCHES_VALID})"

# GITHUB_STEP_SUMMARY is absent when this is run by hand, which is why it is guarded rather than
# assumed - the script stays useful locally.
if [ -n "${GITHUB_STEP_SUMMARY:-}" ]; then
    {
        echo "### Code coverage"
        echo
        echo "| Metric | Covered | Total | Coverage |"
        echo "| --- | ---: | ---: | ---: |"
        echo "| Lines | ${LINES_COVERED} | ${LINES_VALID} | **${LINE_PERCENT}%** |"
        echo "| Branches | ${BRANCHES_COVERED} | ${BRANCHES_VALID} | ${BRANCH_PERCENT}% |"
        echo
        echo "Release configuration. Debug instruments a different line count, so figures from the"
        echo "two builds are not comparable - always quote the denominator and the configuration."
    } >> "$GITHUB_STEP_SUMMARY"
fi
