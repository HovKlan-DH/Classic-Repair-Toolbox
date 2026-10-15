#!/usr/bin/env bash
#
# Stop hook - when CRT's MAJOR or MINOR version has moved past the last release,
# stops the turn ONCE so the agent checks the NuGet packages for updates and
# advises the project owner.
#
# Owner request (2026-10-12): "if I bump the application MAJOR or MINOR version,
# then you will advise me to check if any NuGet packages should be updated". The
# owner bumps InformationalVersion in src/CRT.App/CRT.App.csproj by hand, often in
# Visual Studio, so the hook cannot watch the edit itself: at the end of every
# turn it compares the version on disk with the last RELEASED version.
#
# "Released" is read from tests/CRT.Server.Tests/ApiCompatibility/crt-<version>.txt:
# every release freezes one (CLAUDE.md, "Installed CRTs keep working"). The
# release this version follows is the highest one BELOW it - not the highest of
# all, because the test run writes crt-<version>.txt for a release version as soon
# as it is set, and that file must not count as its own predecessor.
#
# It BLOCKS, unlike the other reminders, because the owner asked for advice, not
# a line in the UI: the agent goes on to list the outdated packages (read-only)
# and recommend, and updates nothing until the owner says which. Once per
# MAJOR.MINOR, stamped in .claude/.nuget-check-reminded, so 3.1.0-alpha.1 ->
# alpha.2 -> 3.1.0 asks once; a PATCH bump (3.0.0 -> 3.0.1) asks nothing.
#
# Turn it off with /hooks, or by deleting its block from .claude/settings.json.

set -uo pipefail

ROOT="$(cd "$(dirname "${BASH_SOURCE[0]}")/../.." && pwd)" || exit 0
cd "$ROOT" || exit 0

CSPROJ="src/CRT.App/CRT.App.csproj"
RELEASES="tests/CRT.Server.Tests/ApiCompatibility"
STAMP=".claude/.nuget-check-reminded"

[ -f "$CSPROJ" ] || exit 0

# The csproj sets InformationalVersion twice - "$(AssemblyVersion)" first, then
# the literal that wins - so take the last line that is a literal.
CURRENT="$(sed -n 's/.*<InformationalVersion>\([^<]*\)<\/InformationalVersion>.*/\1/p' "$CSPROJ" 2>/dev/null \
    | grep -v '\$(' | tail -1 | tr -d '[:space:]')"
CORE="${CURRENT%%[-+]*}"
case "$CORE" in
    [0-9]*.[0-9]*.[0-9]*) ;;
    *) exit 0 ;;   # unreadable version: the build says so, not this hook
esac

# a < b, as versions
version_lt() {
    [ "$1" != "$2" ] && [ "$(printf '%s\n%s\n' "$1" "$2" | sort -V | head -1)" = "$1" ]
}

# The newest release below the current version (release names carry no "-")
PREVIOUS=""
for file in "$RELEASES"/crt-*.txt; do
    [ -f "$file" ] || continue
    name="$(basename "$file" .txt)"
    released="${name#crt-}"
    case "$released" in
        *-*) continue ;;
        [0-9]*.[0-9]*.[0-9]*) ;;
        *) continue ;;
    esac
    if version_lt "$released" "$CORE" && { [ -z "$PREVIOUS" ] || version_lt "$PREVIOUS" "$released"; }; then
        PREVIOUS="$released"
    fi
done
[ -n "$PREVIOUS" ] || exit 0

CURRENT_LINE="$(printf '%s' "$CORE" | cut -d. -f1,2)"
PREVIOUS_LINE="$(printf '%s' "$PREVIOUS" | cut -d. -f1,2)"

# Only a MAJOR or MINOR step counts
version_lt "$PREVIOUS_LINE" "$CURRENT_LINE" || exit 0

# Once per MAJOR.MINOR
if [ -f "$STAMP" ] && [ "$(cat "$STAMP" 2>/dev/null)" = "$CURRENT_LINE" ]; then
    exit 0
fi
printf '%s' "$CURRENT_LINE" > "$STAMP"

printf '{"decision":"block","reason":"CRT is now %s, a MAJOR or MINOR step from the last release %s. The project owner asked to be advised at every such bump to check whether NuGet packages should be updated. Run `dotnet list Classic-Repair-Toolbox.slnx package --outdated` (read-only, about 10 s) and tell the owner, per package: current -> latest, which projects use it, and whether you would take it - read the release notes of a MAJOR jump before recommending it. Mind the project rules: QuestPDF - read its release notes on every bump (WorkbookPdfExporter.ConfigureQuestPdf); xunit stays on xunit.v3.mtp-off (Microsoft Testing Platform is not a package bump); the Avalonia packages move together, with ProDataGrid; a Velopack update also moves the vpk every release job installs (read from CRT.App.csproj), so the macOS Info.plist check may need our plist to follow; EPPlus stays EPPlus. If nothing is outdated, say so in one line. Update nothing until the owner says which."}\n' \
    "$CURRENT" "$PREVIOUS"
exit 0
