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

# 12.0.0, 2026-November-18
--------------------------
- Converting a data tree for CRT 4.0 gives every chip whose datasheet CRT holds all its pins, each with the datasheet's name and direction, and the datasheet's case where the board names none - so the C128 250477's and the CPC 664's chips, which had neither, are drawn
- A pin the board left unnamed gets the datasheet's name; a name the board gives is kept
- The datasheets now include the 8568, the 5710, the CPC's Gate Array, the CX20185 and the 40-pin 64H156 (part number 251828-01 or -03), and the 4464's pins are named /G, D0-D3 and /W
- Surface-mounted chips are understood: a SOIC, SSOP, TSSOP and the like is numbered as a DIP of as many pins, and a PLCC or QFP has its pins on four sides
- A pin past the last pin of a SOIC, PLCC or QFP is refused, as one past a DIP's always was
- The datasheets now include the Amiga 500's and 1000's chips: the 68000, Agnus (also the A500's Fat Agnus in its PLCC-84), Denise, Paula, Gary, the 8520, the 27C400, the 41256 and 44256 RAMs, the MSM6242 clock, the MC1377, MC1488 and MC1489, the LF347, the LM2901 and the A1000's 74374, 74399, 74393 and 7451
- A 74LS51 is told apart from a 7451 or 74S51, whose pins differ
- A crystal oscillator in a metal can is understood, on a DIP-14's or a DIP-8's corners: converting the Open128's KiCad data gives its 16 MHz oscillator that case
- A pin of such an oscillator that is not on one of its four corners is refused
- The conversion's notes say how many chips were given pins or a case
- A submission with such a pin, which the server took before, is now refused - a CRT 4.0 pre-release may have sent one

# 11.1.0, 2026-November-17
--------------------------
- Converting a data tree for CRT 4.0 gives every pin of a chip whose datasheet CRT holds the direction that datasheet says: the 74, 40 and 55 series logic chips, and the CPUs, CIAs, VIAs, video and sound chips, ROMs, RAMs and PLAs of the published boards
- A pin the board says is not connected stays so
- Other chips' pins get the direction their board's KiCad data says, also where the pin was made from an oscilloscope baseline - unless KiCad calls every signal pin of the chip an input
- The expected oscilloscope readings are written in plain words: "High" and "Low" for "HIGH" and "LOW", "Floating" for "FLOAT" and "Video" for "VIDEO" - also inside a longer reading such as "LOW/FLOAT"
- An active-low pin's name is written with "/" before it, as the published boards write theirs: "/IRQ" for KiCad's "~{IRQ}"
- A resistor, capacitor, diode or connector is never taken for a chip, nor a Z80 PIO or CTC for the Z80 itself
- The conversion's notes say how many chips took their datasheet's pin directions, apart from the pinouts made
- Nothing changed in what the server answers

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
