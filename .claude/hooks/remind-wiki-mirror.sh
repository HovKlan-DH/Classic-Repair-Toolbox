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
src/CRT.App/Tabs/Workbooks/|src/CRT.App/Tabs/Worklog/|src/CRT.App/Handlers/Data/Worklog|src/CRT.App/Handlers/Data/Workbook=Workbooks-tab Workbooks-Daily-use Workbooks-Browsing-and-search Workbooks-Export-and-data Workbooks-Getting-started
src/CRT.App/Handlers/Data/SimulationOptions|src/CRT.App/Handlers/Data/DataManager|src/CRT.App/Handlers/Data/DraftManager=Commandline-parameters
src/CRT.Data/BoardDataReader|src/CRT.Data/BoardData\.cs=Board-Excel Main-Excel
src/CRT.Data/BoardComponentHighlightStorage=Board-JSON
src/CRT.App/Handlers/MiniPro/=MiniPro-programmer
src/CRT.App/Handlers/Oscilloscope/|src/CRT.App/Tabs/Oscilloscope/=Synchronize-oscilloscope Controlling-oscilloscope-with-keyboard
src/CRT.App/Handlers/Data/KiCadRawProjectLoader|src/CRT.App/Handlers/Data/KiCadProjectData=KiCad-folder Add-new-board-with-KiCad-data
src/CRT.App/Tabs/Contribute/=Contribute-data-via-CRT Contribute-tab
src/CRT.App/Tabs/Drafts/|src/CRT.App/Main/Main\.NewSystem|src/CRT.Data/NewSystem=Add-new-board-with-KiCad-data Contribute-tab Contribute-data-via-CRT
src/CRT.Data/BoardTable|src/CRT.Data/DraftTableSession=Contribute-data-via-CRT
src/CRT.Data/DraftDrift|src/CRT.Data/DraftRevisionComparer|src/CRT.Data/DraftBaseRevision|src/CRT.App/Main/Main\.DraftDrift=Contribute-data-via-CRT Contribute-tab Board-Excel
src/CRT.Data/DraftRetirement|src/CRT.App/Handlers/Data/PublishedDraftRetirer|src/CRT.Data/DraftWorkbookStore=Contribute-data-via-CRT
src/CRT.Data/DraftFolderImport=Contribute-data-via-CRT Add-new-board-with-KiCad-data
src/CRT.App/Main/Main\.SourceSwitchNotice|src/CRT.Data/SubmissionReceipt=Contribute-data-via-CRT Configuration-tab
src/CRT.Data/DraftFileResolver|src/CRT.Data/DraftBoardSource=Configuration-tab
src/CRT.App/Handlers/Online/UpdateService|src/CRT.App/Handlers/Online/UpdateChannelFilter|src/CRT.App/Handlers/Online/StageFilteredUpdateSource=Configuration-tab
src/CRT.App/CRT\.App\.csproj|Classic-Repair-Toolbox\.slnx=Compiling-yourself-from-source Development-tools-used
src/CRT.App/Handlers/Online/BoardView|src/CRT.App/Main/Main\.BoardViews|src/CRT.Data/BoardViewContract=Information-collected
src/CRT.App/Tabs/Maintainer/|src/CRT.App/Main/Main\.Maintainer|src/CRT.App/Handlers/Maintainer/=Maintainer-tab
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
