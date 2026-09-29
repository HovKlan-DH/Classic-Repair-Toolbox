#!/usr/bin/env bash
#
# Stop hook - points out when CRT.Server changed but its version did not.
#
# The project owner bumps CRT themselves (CRT Maintainer was merged into it on
# 2026-09-29); they asked
# (2026-09-26) for the SERVER's version to be handled for them, judged as real
# SemVer by whoever changes it. That makes the bump an instruction the agent has
# to remember on every server change - exactly the kind of thing this project
# has learned to enforce mechanically rather than trust (see CLAUDE.md on the
# green-tests and Wiki-mirror hooks).
#
# It WARNS, never blocks, for the same reason remind-wiki-mirror.sh does: which
# of MAJOR/MINOR/PATCH a change deserves is a judgement call with no objective
# pass/fail, and a wrong warning must not trap a session. The agent gets the
# nudge and decides - including deciding that a comment-only change needs no
# bump at all.
#
# Deliberately quiet: it fires only when a file the DEPLOYED SERVICE is built
# from changed, it ignores docs, tests and a csproj-only edit (a comment or a
# package bump changes nothing a caller sees), it goes silent as soon as the
# version literal moves, and it repeats at most once per distinct working state.
#
# Turn it off with /hooks, or by deleting its block from .claude/settings.json.

set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)" || exit 0
cd "$ROOT" || exit 0

CSPROJ="src/CRT.Server/CRT.Server.csproj"
HISTORY="src/CRT.Server/VERSION.md"
STAMP=".claude/.server-version-reminded"

[ -f "$CSPROJ" ] || exit 0

# Same base resolution as remind-wiki-mirror.sh: the version must move in the
# same CHANGE as the code, so committed-but-unpushed work counts as changed.
UPSTREAM="$(git rev-parse --abbrev-ref --symbolic-full-name '@{u}' 2>/dev/null || echo '')"
BASE="HEAD"
if [ -n "$UPSTREAM" ] && git merge-base --is-ancestor "$UPSTREAM" HEAD 2>/dev/null; then
    BASE="$UPSTREAM"
fi

CHANGED="$(
    {
        git diff --name-only "$BASE" 2>/dev/null
        git diff --name-only 2>/dev/null
        git diff --name-only --cached 2>/dev/null
        git ls-files --others --exclude-standard 2>/dev/null
    } | sort -u | grep -v '^\.claude/\.'
)"

[ -n "$CHANGED" ] || exit 0

# What the deployed service is actually built from. CRT.Data is included on
# purpose: it ships inside the same publish output and carries no version of its
# own, so a rule changed there changes what this service does. Tests are NOT -
# a test-only change alters nothing a caller can observe, which is PATCH at most
# and usually nothing.
WATCHED='^src/CRT\.Server/|^src/CRT\.Data/'

# Anything under CRT.Server that cannot change the service's behaviour. Docs and
# the version history itself must not trigger the very reminder they answer.
IGNORED='^src/CRT\.Server/(VERSION|DEPLOYMENT|README)\.md$|^src/CRT\.Server/Properties/'

RELEVANT="$(printf '%s\n' "$CHANGED" | grep -E "$WATCHED" | grep -vE "$IGNORED")"
[ -n "$RELEVANT" ] || exit 0

# The csproj on its own is NOT a behaviour change. It is edited for plenty of
# reasons that change nothing a caller sees - a comment, a package bump, a new
# ItemGroup - and this hook's own documentation lives in a comment in it, so
# left in the list the hook fires at the very change that installed it. If the
# csproj is the ONLY thing that moved, stay quiet: a package bump worth a
# version really does arrive with a code change somewhere.
if [ "$RELEVANT" = "$CSPROJ" ]; then
    exit 0
fi

# Has the version literal actually moved in this change? Compare the committed
# base against the working tree rather than trusting that the csproj appears in
# the changed list - the csproj is edited for plenty of reasons that are not a
# version bump (a package bump, a new ItemGroup).
version_of() {
    sed -n 's/.*<InformationalVersion>\([^<]*\)<\/InformationalVersion>.*/\1/p' "$1" 2>/dev/null \
        | grep -v '\$(' | tail -1
}

CURRENT_VERSION="$(version_of "$CSPROJ")"
BASE_VERSION="$(git show "$BASE:$CSPROJ" 2>/dev/null \
    | sed -n 's/.*<InformationalVersion>\([^<]*\)<\/InformationalVersion>.*/\1/p' \
    | grep -v '\$(' | tail -1)"

# Version moved: nothing to say. Also covers the case where the owner bumped it
# by hand.
if [ -n "$CURRENT_VERSION" ] && [ -n "$BASE_VERSION" ] && [ "$CURRENT_VERSION" != "$BASE_VERSION" ]; then
    exit 0
fi

# A version that cannot be read at all is a different problem, and not this
# hook's to shout about - the build and HealthReportTests cover it.
[ -n "$CURRENT_VERSION" ] || exit 0

# One nudge per distinct working state, not one per turn.
CURRENT_STATE="$(printf '%s\n%s' "$RELEVANT" "$CURRENT_VERSION" | sha256sum | cut -d' ' -f1)"
if [ -f "$STAMP" ] && [ "$(cat "$STAMP" 2>/dev/null)" = "$CURRENT_STATE" ]; then
    exit 0
fi
printf '%s' "$CURRENT_STATE" > "$STAMP"

COUNT="$(printf '%s\n' "$RELEVANT" | grep -c . )"
SAMPLE="$(printf '%s\n' "$RELEVANT" | head -3 | tr '\n' ',' | sed 's/,$//; s/,/, /g')"

printf '{"systemMessage":"CRT.Server version: %s file(s) the deployed service is built from changed (%s) but InformationalVersion is still %s. Decide the SemVer bump per src/CRT.Server/VERSION.md - MAJOR if a client that worked can now fail, MINOR for new behaviour that breaks nobody, PATCH for a defect fix or an unobservable change - then edit it in CRT.Server.csproj and add the line to VERSION.md. If nothing a caller can observe changed (comments, tests, a pure refactor), ignore this."}\n' \
    "$COUNT" "$SAMPLE" "$CURRENT_VERSION"
exit 0
