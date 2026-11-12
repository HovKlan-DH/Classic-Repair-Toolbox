#!/usr/bin/env bash
#
# Stop hook - points out when a tab with an introduction tour changed while its tour did not.
#
# The introduction tours (src/CRT.App/Handlers/Tours/TourCatalogue.cs) describe the tabs in words and frame
# their controls. Two things guard them by themselves: MainToursTests fails on a control a step names that is
# gone, and TourLabelsTests on a bold name ("Mark all") no control of the window says any more. What no test
# can see is a NEW feature worth a step, or a button that now does something else than its step says - so
# this names the tours whose tab changed (owner question, 2026-11-12: "Isn't Wiki a hook that never can be
# forgotten/missed, and should this also then be?").
#
# It WARNS, never blocks - as remind-wiki-mirror.sh: whether a change matters to a tour is a judgement call,
# and a hook that blocked on one would trap a session. A tour counts as looked at once ITS part of
# TourCatalogue.cs changed in the same working state - per tab, as the Wiki hook goes per page (code review,
# 2026-11-12: any change there silenced it for every tab) - and it says a working state only once.
#
# Turn it off with /hooks, or by deleting its block from .claude/settings.json.

set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)" || exit 0
cd "$ROOT" || exit 0

CATALOGUE="src/CRT.App/Handlers/Tours/TourCatalogue.cs"
STAMP=".claude/.tours-reminded"

[ -f "$CATALOGUE" ] || exit 0

# Every changed file since HEAD (or the upstream, when ahead of it), committed-but-unpushed included - as the
# Wiki hook reads it. The hooks' own dot-stamps are left out, or writing one would change the very state.
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

# The tours whose part of the catalogue changed in this same working state: they were looked at. A catalogue git
# does not track yet is new as a whole - every tour looked at. Otherwise each changed line (the working tree's
# side of the diff, a deletion counted at the line it follows) belongs to the tour whose "new TourOfTab(" block
# holds it, named by the TourTab.<Tab> right after it; the lines above the first tour are no tour's.
LOOKED_AT=""
if printf '%s\n' "$CHANGED" | grep -qx "$CATALOGUE"; then
    git ls-files --error-unmatch "$CATALOGUE" >/dev/null 2>&1 || exit 0

    # Each tour from its "new TourOfTab(" line, named by the TourTab.<Tab> after it; the last ends at the "];" closing
    # the list (code review, 2026-11-12: from the TourTab line, and the last running to the end of the file, an edit of
    # TourCatalogue.Of below them counted as the "Feedback" tour looked at).
    STARTS="$(awk '
        /new TourOfTab\(/ { start = NR; want = 1; next }
        want && match($0, /TourTab\.[A-Za-z]+/) { print start, substr($0, RSTART + 8, RLENGTH - 8); want = 0; seen = 1; next }
        seen && !ended && /^[[:space:]]*\];/ { print "END", NR; ended = 1 }' "$CATALOGUE")"

    LOOKED_AT="$(
        git diff -U0 "$BASE" -- "$CATALOGUE" 2>/dev/null \
            | sed -n 's/^@@ -[0-9,]* +\([0-9]*\)\(,\([0-9]*\)\)\{0,1\} @@.*/\1 \3/p' \
            | while read -r first count; do
                [ -n "${first:-}" ] || continue
                [ -n "${count:-}" ] || count=1
                [ "$count" -eq 0 ] && count=1
                printf '%s\n' "$STARTS" | awk -v a="$first" -v b="$((first + count - 1))" '
                    $1 == "END" { last = $2; next }
                    { start[++n] = $1; tab[n] = $2 }
                    END { if (!last) last = 1000000000; for (i = 1; i <= n; i++) { stop = (i < n) ? start[i + 1] - 1 : last; if (start[i] <= b && stop >= a) print tab[i] } }'
            done | sort -u
    )"
fi

# code-path regex -> the tour (by its tab) that shows it. The tabs without a tour ("Resources", "About",
# "Maintainer" - Main.TabItemsWithoutTour) are left out on purpose. Main's left panel and worklog bar are
# shown by the "Schematics", "Contribute" and "Workbooks" tours.
MAP="
src/CRT.App/Tabs/Schematics/|src/CRT.App/Main/Main\.axaml$|src/CRT.App/Main/Main\.ComponentPopup|src/CRT.App/Main/Main\.DataSyncStatus=Schematics
src/CRT.App/Tabs/Drafts/|src/CRT.App/Main/Main\.ContributeLists=Contribute
src/CRT.App/Tabs/Workbooks/|src/CRT.App/Tabs/Worklog/|src/CRT.App/Main/Main\.Worklog|src/CRT.App/Main/Main\.DeviceList=Workbooks
src/CRT.App/Tabs/Overview/=Overview
src/CRT.App/Tabs/Oscilloscope/=Oscilloscope
src/CRT.App/Tabs/Configuration/=Configuration
src/CRT.App/Tabs/Feedback/=Feedback
"

HITS=""
while IFS='=' read -r pattern tour; do
    [ -n "${pattern:-}" ] || continue
    printf '%s\n' "$CHANGED" | grep -qE "^($pattern)" || continue
    case " $HITS " in *" $tour "*) ;; *) HITS="$HITS $tour" ;; esac
done <<< "$MAP"

HITS="$(printf '%s' "$HITS" | tr ' ' '\n' | grep -v '^$' | sort -u)"
if [ -n "$LOOKED_AT" ]; then
    HITS="$(printf '%s\n' "$HITS" | grep -vxF -f <(printf '%s\n' "$LOOKED_AT"))"
fi
[ -n "$HITS" ] || exit 0

# One nudge per distinct working state, not one per turn.
CURRENT="$(printf '%s\n%s' "$CHANGED" "$HITS" | sha256sum | cut -d' ' -f1)"
if [ -f "$STAMP" ] && [ "$(cat "$STAMP" 2>/dev/null)" = "$CURRENT" ]; then
    exit 0
fi
printf '%s' "$CURRENT" > "$STAMP"

LIST="$(printf '%s' "$HITS" | tr '\n' ',' | sed 's/,$//; s/,/, /g')"

printf '{"systemMessage":"Introduction tours: the UI of these tabs changed while their tours in TourCatalogue.cs did not - %s. Check each tour against the change (CLAUDE.md, INTRODUCTION TOUR): a feature worth showing gets a step with the version it ships in, a changed one its step reworded and given the new version, a removed one its step deleted. If nothing a tour shows changed, ignore this."}\n' "$LIST"
exit 0
