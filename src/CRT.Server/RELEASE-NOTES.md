<!--
The release notes of each server version, as shown on its release in https://github.com/HovKlan-DH/CRT-Server
(owner request, 2026-11-11: "similar format as the application CHANGELOG, which is a bulletlist containing
one-liners for what has changed"). Maintained by Claude, like VERSION.md: every server bump adds its section
here, newest first. .github/workflows/crt-server-release-notes.py puts a version's bullets on its release at
every push of that version, so a section reworded later reaches the release with the next push.

- A heading is "# <version>, <yyyy>-<Month>-<d>" with a line of dashes under it, as CHANGELOG.md's.
- One line per change, said as the person reading the release list meets it - no file or class names,
  no reasons. The why stays in VERSION.md.
- Plain ASCII punctuation.
-->

# 11.0.0, 2026-November-15
--------------------------
- A board's credits carry what each contribution was and the date it was published
- A new or changed credit must say what the contribution was, or the submission is refused - also when it is the same as a credit already published
- A credit's date that cannot be read as a date is refused
- A new or changed credit's "what" and "which part of it" are at most 60 characters, its name 50 and its contact 100
- A new or changed credit's main contribution is at most 500 characters
- The server dates each new credit as it is published; credits from before have no date
- A new credit is always dated the day it is published, never with a date sent with it
- The credits are kept with the undated first, then oldest first

# 10.4.3, 2026-November-13
--------------------------
- The database driver and the Excel library are updated to their newest versions
- Nothing changed in what the server answers

# 10.4.2, 2026-November-11
--------------------------
- The test server's builds are published in this repository, one release per server version
- The test server takes a build only when its release is named after the version it carries
- Nothing changed for the live server, nor in what a build must prove before it is installed
