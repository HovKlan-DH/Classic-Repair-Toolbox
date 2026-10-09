#!/usr/bin/env bash
#
# Stop hook - says which API version the server and CRT are at whenever the API changed.
#
# The project owner (2026-10-09): "whenever you change anything for the API, then do comment
# what is the API version on SERVER and APPLICATION, so I always can track this." The API
# version is CRT.Data's ClientVersionContract.ApiRevision - the number Account > "Server
# version" shows as "API server version" and "API application version". Both are built
# from that one constant, so a working tree states ONE number for both; they differ only
# between a deployed server and an installed CRT built from different sources.
#
# It REPORTS, never blocks: the line goes to the owner, and CLAUDE.md asks the agent to
# end the turn's summary with the same line ("API version: server N, application N").
#
# Fires only when a file that shapes the API changed against the last commit: the wire
# records (CRT.Data's *Contract.cs), the server's endpoints (*Endpoints.cs), CRT's routes
# and parser, or the recorded API surface (api-revision-*.txt). Says it once per distinct
# working state, as server-version-bump.sh does.
#
# Turn it off with /hooks, or by deleting its block from .claude/settings.json.

set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)" || exit 0
cd "$ROOT" || exit 0

CONTRACT="src/CRT.Data/ClientVersionContract.cs"
CSPROJ="src/CRT.Server/CRT.Server.csproj"
STAMP=".claude/.api-version-reported"

[ -f "$CONTRACT" ] || exit 0

CHANGED="$(
    {
        git diff --name-only HEAD 2>/dev/null
        git ls-files --others --exclude-standard 2>/dev/null
    } | sort -u
)"

WATCHED='^src/CRT\.Data/[A-Za-z]*Contract\.cs$|^src/CRT\.Server/Handlers/.*Endpoints\.cs$|^src/CRT\.App/Handlers/Maintainer/ReviewApi(Routes|Parser[A-Za-z.]*)\.cs$|^tests/CRT\.Server\.Tests/ApiCompatibility/api-revision-[0-9]+\.txt$'

RELEVANT="$(printf '%s\n' "$CHANGED" | grep -E "$WATCHED")"
[ -n "$RELEVANT" ] || exit 0

revision_in() {
    sed -n 's/.*public const int ApiRevision = \([0-9][0-9]*\);.*/\1/p' | tail -1
}

NOW="$(revision_in < "$CONTRACT")"
BEFORE="$(git show "HEAD:$CONTRACT" 2>/dev/null | revision_in)"
SERVER="$(sed -n 's/.*<InformationalVersion>\([^<$]*\)<\/InformationalVersion>.*/\1/p' "$CSPROJ" 2>/dev/null | tail -1)"

[ -n "$NOW" ] || exit 0

CURRENT_STATE="$(printf '%s\n%s\n%s' "$RELEVANT" "$NOW" "$SERVER" | sha256sum | cut -d' ' -f1)"
if [ -f "$STAMP" ] && [ "$(cat "$STAMP" 2>/dev/null)" = "$CURRENT_STATE" ]; then
    exit 0
fi
printf '%s' "$CURRENT_STATE" > "$STAMP"

if [ -n "$BEFORE" ] && [ "$BEFORE" != "$NOW" ]; then
    MOVED="raised from $BEFORE at the last commit - CRTs built for $BEFORE are told to update"
else
    MOVED="unchanged since the last commit - the change only added to the API"
fi

COUNT="$(printf '%s\n' "$RELEVANT" | grep -c . )"

printf '{"systemMessage":"API version: server %s, application %s (ClientVersionContract.ApiRevision, %s). CRT.Server is %s. %s file(s) that shape the API changed since the last commit."}\n' \
    "$NOW" "$NOW" "$MOVED" "${SERVER:-unknown}" "$COUNT"
exit 0
