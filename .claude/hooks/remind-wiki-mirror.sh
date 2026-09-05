#!/usr/bin/env bash
#
# Stop hook - points out when code changed that a mirrored Wiki page documents.
#
# Assets/Wiki/ is the source of truth for the GitHub Wiki, which the maintainer
# updates by hand. CLAUDE.md says documentation ships in the same commit as the
# code it describes, but that is an instruction the agent has to remember, and
# a page that silently goes stale is not discovered until someone reads it and
# is misled. This makes the reminder mechanical instead.
#
# It WARNS, never blocks. A wrong or unnecessary warning must not trap a
# session, and unlike a red test suite there is no objective pass/fail here -
# whether a change is user-visible is a judgement call. The agent gets the
# list and decides.
#
# It is also deliberately CONSERVATIVE about staying quiet: it fires only on
# paths that map to a page, and it goes silent once the mapped page has been
# touched in the same working state. Better to miss an edge case than to cry
# wolf every turn until the noise gets the hook deleted.
#
# Turn it off with /hooks, or by deleting its block from .claude/settings.json.

set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)" || exit 0
cd "$ROOT" || exit 0

WIKI="Assets/Wiki"
STAMP=".claude/.wiki-reminded"

[ -d "$WIKI" ] || exit 0

# Every changed file since HEAD, committed-but-unpushed included: a page must be
# updated in the same CHANGE as the code, not merely before the next push.
#
# The hooks' own dot-stamps under .claude/ are filtered out: they are untracked,
# so writing one would otherwise change the very fingerprint used to decide
# whether this state has already been reported, and the reminder would repeat
# on every single turn.
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

# code-path regex -> the Wiki pages that document it.
# Kept deliberately short: only pages whose content is decided by that code.
MAP="
Tabs/Workbooks/|Tabs/Worklog/|Handlers/Data/Worklog|Handlers/Data/Workbook=Workbooks Workbooks-Daily-use Workbooks-The-Workbooks-tab Workbooks-Export-and-data Workbooks-Getting-started
Handlers/Data/SimulationOptions|Handlers/Data/DataManager=Commandline-parameters
Handlers/Data/BoardDataReader|Handlers/Data/BoardData\.cs=Board-Excel Main-Excel
Handlers/Data/BoardComponentHighlightStorage=Board-JSON
Handlers/MiniPro/=MiniPro-programmer
Handlers/Oscilloscope/|Tabs/Oscilloscope/=Synchronize-oscilloscope Controlling-oscilloscope-with-keyboard
Handlers/Data/KiCadRawProjectLoader|Handlers/Data/KiCadProjectData=KiCad-folder Add-new-board-with-KiCad-data
Tabs/Contribute/=Contribute-data-via-CRT
Classic-Repair-Toolbox\.csproj=Compiling-yourself-from-source Development-tools-used
"

HITS=""
while IFS='=' read -r pattern pages; do
    [ -n "${pattern:-}" ] || continue
    printf '%s\n' "$CHANGED" | grep -qE "^($pattern)" || continue
    for page in $pages; do
        # Already updated in this same change? Then it is not stale - stay quiet.
        printf '%s\n' "$CHANGED" | grep -qx "$WIKI/$page.md" && continue
        [ -f "$WIKI/$page.md" ] || continue
        case " $HITS " in *" $page "*) ;; *) HITS="$HITS $page" ;; esac
    done
done <<< "$MAP"

HITS="$(printf '%s' "$HITS" | tr ' ' '\n' | grep -v '^$' | sort -u)"
[ -n "$HITS" ] || exit 0

# Do not repeat the same reminder for the same working state - one nudge per
# distinct change, not one per turn.
CURRENT="$(printf '%s\n%s' "$CHANGED" "$HITS" | sha256sum | cut -d' ' -f1)"
if [ -f "$STAMP" ] && [ "$(cat "$STAMP" 2>/dev/null)" = "$CURRENT" ]; then
    exit 0
fi
printf '%s' "$CURRENT" > "$STAMP"

LIST="$(printf '%s' "$HITS" | tr '\n' ',' | sed 's/,$//; s/,/, /g')"

printf '{"systemMessage":"Wiki mirror: code changed that these pages document - %s. If the change is user-visible, update the page(s) in Assets/Wiki/ now (same change as the code, per CLAUDE.md) and tell the maintainer which ones need re-pasting into the Wiki. If nothing user-facing changed, ignore this."}\n' "$LIST"
exit 0
