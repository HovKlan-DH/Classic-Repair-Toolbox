#!/usr/bin/env bash
#
# Stop hook - says so when the API VERSION moved:
#
#   API changed: Server changed from [1] to [2], Application changed from [1] to [2]
#
# The project owner (2026-10-09): "whenever you change anything for the API, then do comment
# what is the API version on SERVER and APPLICATION, so I always can track this." Narrowed
# 2026-10-13: "It should only show when this changes" - an API change that leaves the version
# alone (an added optional field or route) says nothing.
#
# The API version is CRT.Data's ClientVersionContract.ApiRevision - the number Account >
# "Server version" shows as "API server version" and "API application version". Both are
# built from that ONE constant, so in a working tree the server and the application always
# move together; they differ only between a deployed server and an installed CRT built from
# different sources, which no file here can tell.
#
# "Moved" is against the last commit. Said once per move; moved back, it is quiet again, and
# a later move is said afresh. It REPORTS, never blocks.
#
# Turn it off with /hooks, or by deleting its block from .claude/settings.json.

set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)" || exit 0
cd "$ROOT" || exit 0

CONTRACT="src/CRT.Data/ClientVersionContract.cs"
STAMP=".claude/.api-version-reported"

[ -f "$CONTRACT" ] || exit 0

revision_in() {
    sed -n 's/.*public const int ApiRevision = \([0-9][0-9]*\);.*/\1/p' | tail -1
}

NOW="$(revision_in < "$CONTRACT")"
BEFORE="$(git show "HEAD:$CONTRACT" 2>/dev/null | revision_in)"

# Nothing to compare, or nothing moved: quiet, and ready to say the next move.
if [ -z "$NOW" ] || [ -z "$BEFORE" ] || [ "$NOW" = "$BEFORE" ]; then
    rm -f "$STAMP"
    exit 0
fi

CURRENT_STATE="$BEFORE->$NOW"
if [ -f "$STAMP" ] && [ "$(cat "$STAMP" 2>/dev/null)" = "$CURRENT_STATE" ]; then
    exit 0
fi
printf '%s' "$CURRENT_STATE" > "$STAMP"

printf '{"systemMessage":"API changed: Server changed from [%s] to [%s], Application changed from [%s] to [%s]"}\n' \
    "$BEFORE" "$NOW" "$BEFORE" "$NOW"
exit 0
