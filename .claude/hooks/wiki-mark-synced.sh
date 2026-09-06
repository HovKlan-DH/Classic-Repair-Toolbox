#!/usr/bin/env bash
#
# Stamps Wiki pages as pasted into the live Wiki.
#
# Run when the maintainer confirms a paste. It records each named page against
# the CURRENT content in .claude/wiki-synced.tsv, which is what the Stop hook
# wiki-sync-status.sh compares against to build the waiting list. Once stamped,
# a page drops off that list until it changes again.
#
#   .claude/hooks/wiki-mark-synced.sh Configuration-tab Workbooks-tab
#   .claude/hooks/wiki-mark-synced.sh --all      # everything currently waiting
#
# WHAT IS ACTUALLY RECORDED, AND WHY IT IS A CONTENT HASH
#
# Each row stores the page's git blob hash - the hash of the exact bytes that
# were pasted - alongside the commit and date. The hash is what the Stop hook
# compares; the commit and date are for the reader.
#
# A commit SHA alone is NOT enough, and this was caught in testing. Wiki pages
# are routinely pasted while their edits are still uncommitted (that is the
# normal state when the maintainer reads the file). Comparing commits then, a
# page stamped at HEAD still looks dirty against HEAD forever, so every stamped
# page bounced straight back onto the waiting list and the list could never be
# cleared. A content hash answers the real question - "are these the bytes that
# were pasted?" - and is correct whether or not the work has been committed.

set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)" || exit 1
cd "$ROOT" || exit 1

WIKI="Assets/Wiki"
STATE=".claude/wiki-synced.tsv"
TODAY="$(date +%Y-%m-%d)"
HEAD_SHA="$(git rev-parse HEAD 2>/dev/null || echo unknown)"

[ -f "$STATE" ] || { echo "No $STATE - nothing to stamp." >&2; exit 1; }
[ $# -gt 0 ] || { echo "Usage: $(basename "$0") <page>... | --all" >&2; exit 1; }

PAGES=()
if [ "$1" = "--all" ]; then
    while IFS= read -r f; do
        p="$(basename "$f" .md)"
        case "$p" in README | '!sync-status') continue ;; esac
        PAGES+=("$p")
    done < <(find "$WIKI" -maxdepth 1 -name '*.md' | LC_ALL=C sort)
else
    PAGES=("$@")
fi

STAMPED=0
for page in "${PAGES[@]}"; do
    page="${page%.md}"
    if [ ! -f "$WIKI/$page.md" ]; then
        echo "  ! no such page: $page" >&2
        continue
    fi
    # Drop any existing row, then append the new one.
    awk -F'\t' -v p="$page" '$1 != p' "$STATE" > "$STATE.tmp" 2>/dev/null || continue
    hash="$(git hash-object "$WIKI/$page.md" 2>/dev/null || echo unknown)"
    printf '%s\t%s\t%s\t%s\n' "$page" "$HEAD_SHA" "$TODAY" "$hash" >> "$STATE.tmp"
    # Keep comments on top, rows sorted, so the file diffs cleanly.
    { grep '^#' "$STATE.tmp"; grep -v '^#' "$STATE.tmp" | grep -v '^$' | LC_ALL=C sort; } > "$STATE"
    rm -f "$STATE.tmp"
    echo "  stamped $page"
    STAMPED=$((STAMPED + 1))
done

echo "Stamped $STAMPED page(s) at ${HEAD_SHA:0:8} on $TODAY."
