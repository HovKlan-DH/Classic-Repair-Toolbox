#!/usr/bin/env bash
#
# Stop hook - refuses to hand back to the user while the test suite is red.
#
# This enforces rule 3 of .claude/CLAUDE.md ("ALWAYS run dotnet test before
# reporting any code change as done") mechanically, so that it no longer
# depends on the agent choosing to remember it.
#
# Behaviour:
#   * Code unchanged since the last green run -> exits silently and instantly.
#   * Code changed -> runs the suite in Release (same as CI, and clear of the
#     bin/Debug lock held while the app is running). Green: records state, exits 0.
#     Red: exits 2, which blocks the handover and feeds the failing tests back
#     to Claude so it has to deal with them.
#   * Already blocked once for this same stop -> warns loudly but does not block
#     again, so a suite the agent cannot fix never traps the session in a loop.
#
# Turn it off temporarily with /hooks, or permanently by deleting the "Stop"
# block from .claude/settings.json.

set -uo pipefail

# The repo root is two levels up from this script (.claude/hooks/<script>).
ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)" || exit 0
cd "$ROOT" || exit 0

SOLUTION="Classic-Repair-Toolbox.slnx"
STAMP=".claude/.tests-green"

INPUT="$(cat)"

# Claude Code sets stop_hook_active when the stop is itself the result of a
# previous stop hook blocking. Blocking again from there risks an endless
# fix/stop/block cycle, so the second pass only warns.
ALREADY_BLOCKED=0
case "$INPUT" in
    *'"stop_hook_active":true'* | *'"stop_hook_active": true'*) ALREADY_BLOCKED=1 ;;
esac

# A content fingerprint of the current code: HEAD, plus every uncommitted and
# untracked C#/project file. Identical to the last green run means there is
# genuinely nothing new to verify - which is what keeps question-and-answer
# turns free instead of paying 30 seconds each.
fingerprint()
{
    {
        git rev-parse HEAD 2>/dev/null
        git diff HEAD -- '*.cs' '*.csproj' '*.axaml' '*.slnx' 2>/dev/null
        git ls-files --others --exclude-standard -- '*.cs' '*.csproj' '*.axaml' '*.slnx' 2>/dev/null |
            while IFS= read -r f; do printf '%s\n' "$f"; cat "$f" 2>/dev/null; done
    } | sha256sum | cut -d' ' -f1
}

CURRENT="$(fingerprint)"

if [ -f "$STAMP" ] && [ "$(cat "$STAMP" 2>/dev/null)" = "$CURRENT" ]; then
    exit 0
fi

OUTPUT="$(dotnet test "$SOLUTION" -c Release --nologo 2>&1)"
STATUS=$?

if [ $STATUS -eq 0 ]; then
    printf '%s' "$CURRENT" > "$STAMP"
    exit 0
fi

# Pull out the parts worth reading: the failing test names, their assertions,
# any compiler error, and the final tally.
SUMMARY="$(printf '%s\n' "$OUTPUT" |
    grep -E 'error [A-Z]+[0-9]+|^[[:space:]]*(Failed|Error Message|Assert|Expected|Actual|Stack Trace)|Failed!|Passed!|Build FAILED|FATAL ERROR' |
    head -40)"

# A test PROCESS that died - xunit v3 runs each test assembly as a child process, and its adapter
# reports the death as "[FATAL ERROR] Xunit.Sdk.TestPipelineException" - still ends in a "Passed!"
# line counting only the tests that ran before it. Say so, or the block reads as all green.
CRASHED=""
case "$OUTPUT" in
    *'[FATAL ERROR]'*)
        CRASHED="A test process DIED (the [FATAL ERROR] line): its \"Passed!\" line counts only the tests that ran before. Its cause is in Windows' Application log, \".NET Runtime\" event 1023 (an \"internal error\", 0x80131506, has hit this machine's .NET processes since 2026-11-02 - the compiler server too - so it may be no test's fault: run the suite again)."
        ;;
esac

# A locked output file means the app is running, so the build could not even be
# attempted. That is not a red suite and must not block a handover - say so and move on.
case "$OUTPUT" in
    *MSB3027* | *MSB3021* | *'being used by another process'*)
        echo '{"systemMessage":"Tests were skipped: the build output is locked because Classic Repair Toolbox is still running. Close the app to let the Stop hook verify the suite."}'
        exit 0
        ;;
esac

if [ "$ALREADY_BLOCKED" -eq 1 ]; then
    echo '{"systemMessage":"WARNING: handing over with a RED test suite - dotnet test still fails. The Stop hook already blocked once and will not block again this turn."}'
    exit 0
fi

{
    echo "BLOCKED: 'dotnet test' does not pass, so this change is not done."
    echo
    echo "$SUMMARY"
    echo
    [ -n "$CRASHED" ] && { echo "$CRASHED"; echo; }
    echo "Fix this before handing back. Per rule 4 of .claude/CLAUDE.md, a failing test is a"
    echo "question, not an obstacle: decide whether the behaviour change was intended. If it was,"
    echo "update the expectation and say so explicitly in your summary. Never edit or weaken an"
    echo "assertion just to reach green, and never delete a test to make a change pass."
} >&2

exit 2
