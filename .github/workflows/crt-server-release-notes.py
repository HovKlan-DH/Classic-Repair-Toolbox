# ###########################################################################################
# The release notes of a server version's build of main, for its release in HovKlan-DH/CRT-Server (owner
# requests, 2026-11-11: "do manage by yourself a release notes per version", and then "similar format as the
# application CHANGELOG, which is a bulletlist containing one-liners for what has changed").
#
#   python crt-server-release-notes.py <server version> <commit> <RELEASE-NOTES.md> <VERSION.md> > notes.md
#   python crt-server-release-notes.py --has-section <server version> <RELEASE-NOTES.md>
#
# The bullets are that version's section of src/CRT.Server/RELEASE-NOTES.md ("# 10.4.2, 2026-November-11"),
# rewritten at every push of the version, so a section reworded later reaches the release. Under them one line
# says the bump and the API revision - from the version's row of VERSION.md, read only in its LIVE table (the
# archived numberings below it carry versions without backticks) - and the commit built. A version with no
# section yet, or a file missing, still gets notes, saying so: missing notes never stop the build reaching the
# test server (crt-server-publish-main.sh runs under set -e).
#
# --has-section is the Stop hook server-version-bump.sh's question, so the hook and the release find a section by
# ONE rule: exit 0 when the version has a section, 3 when it has none.
# ###########################################################################################
import os
import re
import sys

REPOSITORY = "HovKlan-DH/Classic-Repair-Toolbox"
NO_SECTION = 3


def read_lines(path):
    # utf-8-sig: a byte-order mark an editor put in front would hide the first heading.
    if not os.path.isfile(path):
        return []

    with open(path, encoding="utf-8-sig") as handle:
        return handle.read().splitlines()


def section_of(lines, version):
    # The lines under "# <version>," up to the next "# " heading; the line of dashes right under the heading is
    # the heading's underline and left out, any other line kept as written.
    found = None

    for line in lines:
        if line.startswith("# "):
            if found is not None:
                break
            if line.startswith("# " + version + ","):
                found = []
            continue

        if found is not None:
            found.append(line)

    if found is None:
        return None

    while found and not found[0].strip():
        found.pop(0)
    if found and re.fullmatch(r"-{3,}\s*", found[0]):
        found.pop(0)
    while found and not found[0].strip():
        found.pop(0)
    while found and not found[-1].strip():
        found.pop()

    return found or None


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
            cells = [cell.strip() for cell in line.strip().strip("|").split("|")]
            # Version | Date | Bump | API | What changed
            return cells if len(cells) >= 4 else None

    return None


def absolute_links(text, commit):
    # A link the notes make relative to their own folder points into CRT's repository at the commit built - on
    # CRT-Server's release page it would resolve against the wrong repository.
    base = "https://github.com/" + REPOSITORY + "/blob/" + commit + "/src/CRT.Server/"
    return re.sub(r"\]\((?!https?:|mailto:|#)([^)]+)\)", lambda m: "](" + base + m.group(1) + ")", text)


def main():
    if len(sys.argv) == 4 and sys.argv[1] == "--has-section":
        sys.exit(0 if section_of(read_lines(sys.argv[3]), sys.argv[2]) is not None else NO_SECTION)

    version, commit, notes_file, version_file = sys.argv[1], sys.argv[2], sys.argv[3], sys.argv[4]

    section = section_of(read_lines(notes_file), version)
    row = row_of(read_lines(version_file), version)

    if section is None:
        print("- No release notes for this version yet")
    else:
        print(absolute_links("\n".join(section), commit))

    commit_link = "[" + commit[:7] + "](https://github.com/" + REPOSITORY + "/commit/" + commit + ")"
    history = "https://github.com/" + REPOSITORY + "/blob/main/src/CRT.Server/VERSION.md"
    kind = (row[2] + ", API revision " + row[3] + ". ") if row is not None else ""

    print()
    print("---")
    print()
    print(kind + "Built from commit " + commit_link + " of Classic Repair Toolbox, for the project's own test "
          "server only - nobody else needs it. The reasons behind each version: [version history](" + history + ").")


main()
