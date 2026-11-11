# ###########################################################################################
# The release notes of a server version's build of main, for its release in HovKlan-DH/CRT-Server (owner
# request, 2026-11-11: "do manage by yourself a release notes per version"). They are that version's row of
# src/CRT.Server/VERSION.md - written at every server bump anyway - so the notes never need writing twice, and
# a row reworded later reaches the release with the next push of that version.
#
#   python crt-server-release-notes.py <server version> <commit> <VERSION.md> > notes.md
#
# Reads only the LIVE table (the first "| Version |" table): the archived numberings below it carry versions
# without backticks, so a restarted number can never pick an old row. A version with no row yet still gets
# notes, saying so - a missing row never stops the build reaching the test server.
# ###########################################################################################
import re
import sys

REPOSITORY = "HovKlan-DH/Classic-Repair-Toolbox"


def row_of(lines, version):
    in_table = False

    for line in lines:
        if line.startswith("| Version |"):
            if in_table:
                break
            in_table = True
            continue

        if in_table and not line.startswith("|"):
            break

        if in_table and line.startswith("| `" + version + "` |"):
            cells = line.strip().strip("|").split("|")
            if len(cells) < 5:
                return None
            # Version | Date | Bump | API | What changed - the last cell taken whole, in case it holds a "|".
            return [cell.strip() for cell in cells[:4]] + ["|".join(cells[4:]).strip()]

    return None


def absolute_links(text, commit):
    # A link VERSION.md makes relative to its own folder points into CRT's repository at the commit built.
    base = "https://github.com/" + REPOSITORY + "/blob/" + commit + "/src/CRT.Server/"
    return re.sub(r"\]\((?!https?:|mailto:|#)([^)]+)\)", lambda m: "](" + base + m.group(1) + ")", text)


def main():
    version, commit, version_file = sys.argv[1], sys.argv[2], sys.argv[3]

    with open(version_file, encoding="utf-8") as handle:
        lines = handle.read().splitlines()

    row = row_of(lines, version)
    commit_link = "[" + commit[:7] + "](https://github.com/" + REPOSITORY + "/commit/" + commit + ")"
    history = "https://github.com/" + REPOSITORY + "/blob/main/src/CRT.Server/VERSION.md"

    if row is None:
        print("**CRT.Server " + version + "**")
        print()
        print("No entry for this version in the [version history](" + history + ") yet.")
    else:
        _, date, bump, api, what = row
        print("**CRT.Server " + version + "** - " + date + " - " + bump + " - API revision " + api)
        print()
        print(absolute_links(what, commit))

    print()
    print("---")
    print()
    print("Built from commit " + commit_link + " of Classic Repair Toolbox's `main` branch, for the project's own test "
          "server, which installs it by itself. Nobody else needs it - see the README. Every version: "
          "[version history](" + history + ").")


main()
