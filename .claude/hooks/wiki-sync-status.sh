#!/usr/bin/env bash
#
# Stop hook - keeps Assets/Wiki/!sync-status.md correct by itself.
#
# WHY THIS EXISTS
#
# !sync-status.md records which Wiki pages still need pasting into the live
# Wiki. It used to record that only in PROSE ("Verified against the live Wiki
# on 2026-09-06"), which nothing could verify - so when six pages had changed
# and the prose still named three, nothing noticed. The maintainer had to ask
# for a manual double-check every time, which is the job this script now does.
#
# The machine-readable half is .claude/wiki-synced.tsv: one row per page,
# recording the blob hash of the bytes last pasted. This script compares every
# page against its row and REWRITES the marked block in !sync-status.md with
# the true list.
#
# THE FILE IS THE TABLE, AND NOTHING ELSE
#
# The maintainer reads this file for one purpose: to go and paste the pages it
# names. Everything else that used to sit in it - how the mechanism works, what
# changed in each page, which files are not Wiki pages - was noise they never
# read, and it pushed the table itself below the fold. The whole file is now
# generated between the markers and holds only the table. Explanations of the
# mechanism live in .claude/CLAUDE.md and Assets/Wiki/README.md, where they
# belong; the per-page "what changed" prose is not written at all any more.
#
# TWO COLUMNS: the FILE, and WHERE TO FIND IT LIVE
#
# The old table's other columns ("Last pasted", "Changed since") answered a
# question nobody asked - a page is in the table because it needs pasting, and
# how stale it is changes nothing about what to do about it. What was MISSING
# is the thing that costs real time: where in the live Wiki that page actually
# is. So the second column is a navigation trail: "Home > Workbooks > Daily
# use".
#
# That trail is DERIVED FROM _Sidebar.md at run time, not hardcoded here. The
# sidebar is what GitHub renders beside every Wiki page, so it is what the
# maintainer navigates by; deriving it means a sidebar edit cannot leave this
# table describing a structure that no longer exists. A page the sidebar does
# not list falls back to "Home".
#
# WHAT IT DOES NOT DO
#
# It never blocks (exit 0 always). A wrong warning must not trap a session.
#
# Turn it off with /hooks, or by deleting its block from .claude/settings.json.

set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)" || exit 0
cd "$ROOT" || exit 0

WIKI="Assets/Wiki"
STATE=".claude/wiki-synced.tsv"
STATUS="$WIKI/!sync-status.md"
SIDEBAR="$WIKI/_Sidebar.md"
BEGIN="<!-- crt:waiting-start -->"
END="<!-- crt:waiting-end -->"

[ -d "$WIKI" ] || exit 0
[ -f "$STATE" ] || exit 0
[ -f "$STATUS" ] || exit 0

# ---------------------------------------------------------------------------
# Build "page<TAB>Home > Section > Page" for every page named in _Sidebar.md.
#
# The sidebar's shape is: "### Section" or "**Section**" headings, top-level
# "[Title](Page)" links, and "- [Title](Page)" bullets belonging to whichever
# link or heading precedes them. Extra indentation on a bullet ("  - ") makes
# it a child of the previous bullet, which is how the Data files group is
# written.
#
# A page listed TWICE (Workbooks-tab is under both "At the bench" and "The
# tabs") keeps its FIRST trail: that is the deeper, more specific one, and one
# answer is more useful than two.
# ---------------------------------------------------------------------------
TRAILS=""
if [ -f "$SIDEBAR" ]; then
    TRAILS="$(awk '
        function emit(page, trail) {
            if (page != "" && !(page in seen)) { seen[page] = 1; print page "\t" trail }
        }
        # "### At the bench" - a section heading.
        /^###[[:space:]]+/ {
            section = $0; sub(/^###[[:space:]]+/, "", section)
            parent = ""; child = ""; next
        }
        # "**Data files**" - also a section heading, written bold.
        /^\*\*[^*]+\*\*[[:space:]]*$/ {
            section = $0; gsub(/^\*\*|\*\*[[:space:]]*$/, "", section)
            parent = ""; child = ""; next
        }
        # Any line carrying a [Title](Page) link.
        match($0, /\[[^]]*\]\([^)]*\)/) {
            link = substr($0, RSTART, RLENGTH)
            title = link; sub(/^\[/, "", title); sub(/\].*$/, "", title)
            page  = link; sub(/^.*\(/, "", page);  sub(/\)$/, "", page)

            indent = 0
            if (match($0, /^[[:space:]]*/)) indent = RLENGTH
            bullet = ($0 ~ /^[[:space:]]*-[[:space:]]/)

            trail = "Home"
            if (section != "") trail = trail " > " section

            if (!bullet) {
                # A top-level link: it becomes the parent for bullets below it.
                parent = title; child = ""
                emit(page, trail " > " title)
            } else if (indent >= 2) {
                # Indented bullet: a child of the previous bullet.
                if (parent != "") trail = trail " > " parent
                if (child  != "") trail = trail " > " child
                emit(page, trail " > " title)
            } else {
                # Plain bullet: under the current parent link, else the section.
                child = title
                if (parent != "") trail = trail " > " parent
                emit(page, trail " > " title)
            }
        }
    ' "$SIDEBAR" 2>/dev/null)"
fi

lookup_trail() {
    # $1 = page name. Falls back to "Home" for anything the sidebar omits -
    # which is Home itself, and any page added here before the sidebar lists it.
    local found
    found="$(printf '%s\n' "$TRAILS" | awk -F'\t' -v p="$1" '$1 == p { print $2; exit }')"
    [ -n "$found" ] || found="Home"
    printf '%s' "$found"
}

# ---------------------------------------------------------------------------
# Work out which pages are waiting.
#
# THE COMPARISON IS ON CONTENT, NOT ON COMMITS. Pages are normally pasted while
# their edits are still uncommitted, so a commit-based check reports every
# stamped page as dirty against HEAD forever, and the list could never be
# cleared - which is exactly what testing showed. The blob hash answers the
# question that actually matters: are these the bytes that were pasted?
# ---------------------------------------------------------------------------
WAITING=""

while IFS= read -r file; do
    page="$(basename "$file" .md)"
    case "$page" in README | '!sync-status') continue ;; esac

    # Exact field-1 match via awk: grep -P is unavailable in some Git Bash
    # locales ("-P supports only unibyte and UTF-8 locales"), and a plain
    # grep pattern would treat a page name's own characters as regex.
    row="$(awk -F'\t' -v p="$page" '$1 == p { print; exit }' "$STATE" 2>/dev/null)"

    if [ -z "$row" ]; then
        # No row at all: a brand-new page that has never been pasted.
        WAITING="$WAITING$page	$(lookup_trail "$page")
"
        continue
    fi

    pasted_hash="$(printf '%s' "$row" | cut -f4)"
    current_hash="$(git hash-object "$file" 2>/dev/null || echo unknown)"

    if [ -n "$pasted_hash" ] && [ "$pasted_hash" = "$current_hash" ]; then
        continue
    fi

    WAITING="$WAITING$page	$(lookup_trail "$page")
"
done < <(find "$WIKI" -maxdepth 1 -name '*.md' | LC_ALL=C sort)

COUNT="$(printf '%s' "$WAITING" | grep -c . || true)"

# ---------------------------------------------------------------------------
# Build the replacement block. The whole file lives between the markers.
# ---------------------------------------------------------------------------
NL="
"
BLOCK="$BEGIN$NL"
BLOCK="$BLOCK<!-- Maintained automatically by .claude/hooks/wiki-sync-status.sh.$NL"
BLOCK="$BLOCK     Do not hand-edit; your edits are overwritten on the next turn.$NL"
BLOCK="$BLOCK     Tell Claude which pages you have pasted and this clears itself. -->$NL$NL"
BLOCK="$BLOCK# Wiki pages waiting to be pasted$NL$NL"

if [ "$COUNT" -eq 0 ]; then
    BLOCK="$BLOCK**Nothing to paste.** Every page matches what is in the live Wiki.$NL"
else
    label="pages"
    [ "$COUNT" -eq 1 ] && label="page"
    BLOCK="$BLOCK**$COUNT $label waiting.**$NL$NL"
    BLOCK="$BLOCK| File in \`Assets/Wiki\` | Where it is in the Wiki |$NL"
    BLOCK="$BLOCK| --- | --- |$NL"
    while IFS=$'\t' read -r page trail; do
        [ -n "$page" ] || continue
        BLOCK="$BLOCK| \`$page.md\` | $trail |$NL"
    done <<< "$WAITING"
fi

BLOCK="$BLOCK$NL$END"

# Splice it in, replacing whatever is between the markers.
if ! grep -qF "$BEGIN" "$STATUS" || ! grep -qF "$END" "$STATUS"; then
    exit 0
fi

TMP="$(mktemp)" || exit 0
awk -v begin="$BEGIN" -v end="$END" -v block="$BLOCK" '
    index($0, begin) { print block; skip = 1; next }
    index($0, end)   { skip = 0; next }
    !skip            { print }
' "$STATUS" > "$TMP" 2>/dev/null || { rm -f "$TMP"; exit 0; }

if ! cmp -s "$TMP" "$STATUS"; then
    cp "$TMP" "$STATUS"
    rm -f "$TMP"
    if [ "$COUNT" -eq 0 ]; then
        echo '{"systemMessage":"Wiki sync: !sync-status.md updated - nothing is waiting to be pasted."}'
    else
        LIST="$(printf '%s' "$WAITING" | cut -f1 | tr '\n' ',' | sed 's/,$//; s/,/, /g')"
        printf '{"systemMessage":"Wiki sync: refreshed the waiting list in !sync-status.md (%s waiting: %s)."}\n' "$COUNT" "$LIST"
    fi
fi
rm -f "$TMP" 2>/dev/null
exit 0
