<!--
The server's changelog: the release notes of each server version, as shown on its release in
https://github.com/HovKlan-DH/CRT-Server (owner request, 2026-11-11: "similar format as the application
CHANGELOG, which is a bulletlist containing one-liners for what has changed") AND in CRT, under the "Maintainer"
tab's Account > "Server version" (owner request, 2026-11-24: 'have a "Changelog" like there is in "About" ... so I
can check what has changed in server also') - the server is built with this file inside it and sends it from its
first heading on, so this comment is never shown. Maintained by Claude, like VERSION.md: every server bump adds its
section here, newest first, in the same change. .github/workflows/crt-server-release-notes.py puts a version's
bullets on its release at every push of that version, so a section reworded later reaches the release with the
next push. The sections from 10.4.1 down to 1.0.0 were written from VERSION.md's rows on 2026-11-24.

- A heading is "# <version>, <yyyy>-<Month>-<d>" with a line of dashes under it, as CHANGELOG.md's.
- One line per change, said as the person reading the release list meets it - no file or class names,
  no reasons. The why stays in VERSION.md.
- Plain ASCII punctuation.
-->

# 14.0.0, 2026-November-25
--------------------------
- An account can sign in only once its address is confirmed - until then it is answered as a wrong password, so nothing tells whether an address has an account
- A display name is held only once its account's address is confirmed; of two new accounts asking for one name, the second to confirm gets it with a number after it, and a mail saying so
- An expired link to confirm an address points to "I forgot my password", whose code confirms the address too
- "Send invitation" with an address that already has an account makes that account the board's maintainer, instead of refusing
- A form of the account web pages sent from another website is refused, so no other site can make its visitors' browsers send mail through them
- The mails about an account also name the "Maintainer" tab's sign-in, for CRT 4.0 and older

# 13.0.0, 2026-November-25
--------------------------
- Anybody can make an account - in CRT, or on the website's new account pages - and the same account works in both
- A display name is now 2 to 30 characters of A-Z, 0-9 and + - / @ . [ ] ( ), and no two accounts can share one, whatever the capitals; a name set before is kept
- The website has pages to make an account, confirm its address and set a new password with a code from a mail
- The link that confirms an address opens a page instead of a line of text
- A signed-in account can ask for a new link to confirm its address
- The mails point to the account icon in CRT's top right corner, and the password mails also to the website's page for the code
- Asking for a new password says a code is on its way, not a link - the mail has carried a code since September

# 12.2.0, 2026-November-24
--------------------------
- The server sends its own changelog to signed-in maintainers, so the "Maintainer" tab can show what changed in the server deployed
- The changelog goes back to 1.0.0

# 12.1.1, 2026-November-24
--------------------------
- The library that reads Excel workbooks when a data tree is converted is updated (EPPlus 8.7.3); nothing it converts comes out differently

# 12.1.0, 2026-November-22
--------------------------
- Converting a data tree for CRT 4.0 dates every board and the list of boards and oscilloscopes on the day of the conversion, as a publish dates what it writes
- The conversion's report says the date it gave

# 12.0.1, 2026-November-20
--------------------------
- The mails ask you to "check" a checkbox in CRT, as CRT's own words do

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

# 10.4.1, 2026-November-11
--------------------------
- A region's colors sent without the board's list of regions are held to the rules for a region's name, and refused when a name breaks them
- A color for a region the published board does not have is left out at the publish, never written
- The automatic update looks again, an hour later, at a build whose proof of origin it could not check, instead of waiting for a retry by hand
- A build held for its systemd or polkit files keeps the mark that they are installed when the next build of the same version ships the same files
- A held build is never installed over a build as new or newer that runs since
- Without the GitHub command-line tool the update refuses the build and looks again an hour later, instead of downloading it every minute
- An update starting as another ends no longer unpacks the build just installed again over the one the service runs from

# 10.4.0, 2026-November-11
--------------------------
- A board's regions can have colors of their own - a background and a text color each - published with the board and kept when a submission does not send them
- Region colors not written as #RRGGBB, for a region the board does not have, or given twice for one region are refused
- A change of the region colors alone is a change, said in the review's summary and in the board's history

# 10.3.3, 2026-November-9
-------------------------
- A build the automatic update unpacks can be read by the service, so it starts - 10.3.1 and 10.3.2 went back to 10.3.0 for that
- A box whose updater is 10.3.0, 10.3.1 or 10.3.2 needs one command run by hand, once, before this build can start

# 10.3.2, 2026-November-9
-------------------------
- A build held back for its systemd or polkit files says the commands that install them
- A unit installed with a plain copy could not be read by the updater, which kept the build held; installed as the commands say, it can

# 10.3.1, 2026-November-9
-------------------------
- The test server gets data trees of its own, apart from the live service's: its systemd unit may write only those, and CRT's alphas and betas download from them
- The test server's updater holds this build until its new unit is installed
- Nothing changed for the live server

# 10.3.0, 2026-November-9
-------------------------
- The server can update itself from GitHub, looking every minute: the live service takes the build a real CRT release carries, the test server the build of every push to main whose tests passed
- A build is installed only once GitHub's attestation proves CRT's own workflows built it, on main, on a GitHub-hosted runner
- A look reads one small file on GitHub and downloads only a newer build
- After installing, the update waits for the new version to answer and goes back to the build before when it does not
- A build that ships a changed systemd or polkit file is held until the project owner has installed those
- The server's status tells CRT how its updates go: whether they are set up, the channel, the last look and what came of it, and the last build installed
- It needs a one-time setup on the box; without it nothing else changes

# 10.2.1, 2026-November-9
-------------------------
- A board said to come in one region alone is refused, with words saying a board told apart by region comes in two at least, or none
- Nothing else changed

# 10.2.0, 2026-November-9
-------------------------
- A board may name regions of its own besides PAL and NTSC: at most 6, each 1 to 20 letters, digits, spaces and "- + / ( ) .", never "All regions"
- A region name breaking those rules, or given twice, is refused as before, with words saying why
- A file or link of a region the board does not have is a warning in the table, naming the board's regions
- A component can carry a "Purpose or note", a note with formatting, published as sent
- A submission's change summary and a board's history say the region names before and now
- A warning names "Short description" as "Other technical details", and a refusal lists the pin directions in a new order

# 10.1.2, 2026-November-8
-------------------------
- The mail telling the administrators an oscilloscope model was sent in names "Management" > "Oscilloscope models", where that entry is, instead of "Account"
- The server's log names a board file it cannot read once until the file changes, not at every read of it
- A picture's or baseline's note changed only in how deep its first line's bullet is nested counts as a change in a submission's summary and the maintainer's table

# 10.1.1, 2026-November-7
-------------------------
- The messages sending the administrator to "Rebuild checksum manifests" or "Maintainers" name the "Maintainer" tab's new "Management" screen instead of "Account"

# 10.1.0, 2026-November-6
-------------------------
- A board's files and links can carry a note of several lines on what each is about, published as sent and shown in the review
- Converting a data tree for CRT 4.0 gives every board its pinouts: the pins its oscilloscope baselines and KiCad data say, and the package where its footprint or baselines make it certain - only where the board has none

# 10.0.1, 2026-November-4
-------------------------
- A refusal for a needed field left empty puts the field's name in quotes: This credit has no "Name or handle".

# 10.0.0, 2026-November-2
-------------------------
- A file of no bytes is refused, whatever its type
- A highlight coordinate a board file cannot hold is refused, and one already waiting in the queue stops its publish with words instead of a server error
- Pushing BETA back to the queue is refused when BETA has changed since the confirmation was shown
- An oscilloscope model sent in says whether a verified account sent it, and the administrators' mail says the test suite result is what the sender's CRT reported
- A contributor counts as signed in only with a verified account
- A submission's change summary compares a component's several rectangles on one schematic image by all their values, so rectangles only in another order are no change, and the counts may rise
- The files an approval, a board edit or a publish to stable removes, and the board's history, no longer name a file replaced under its own name - only what nothing uses any more
- A maintainer's amendment with findings too long is cut to fit, as at creation, instead of failing after it was saved

# 9.0.0, 2026-November-2
------------------------
- A component's repair notes are gone from submissions and the review
- An oscilloscope baseline with no signal name or no image is refused
- Any component image naming a pin is now an oscilloscope baseline, its T/DIV, V/DIV and trigger level all optional - no published row changes kind
- A file or link, of a component or of the board, can belong to one region; left blank it is of every region

# 8.0.0, 2026-November-1
------------------------
- A replacement part has one note of several lines in place of its "changes" and "source", in submissions, the review and the published board files
- An oscilloscope baseline no longer says where it comes from, and the warning about that is gone

# 7.0.0, 2026-November-1
------------------------
- A submission naming a region other than PAL or NTSC is refused
- A pin whose digits run straight into a letter, such as "12V" or "1A", is a name and no longer refused as a pin past its package; "2 [64K]" is still pin 2

# 6.0.0, 2026-November-1
------------------------
- A submission with a component that has no category is refused
- An oscilloscope baseline or a pin on a pin number its component's package does not have, such as pin 48 of a DIP-24, is refused
- A board can say whether it comes in regions, published only when it says any; a submission that does not say keeps what is published
- The review, the board's history and an amendment carry the board's regions

# 5.0.2, 2026-November-1
------------------------
- Messages and mails put the name of a tab or a button in quotes: the "Maintainer" tab, press "Delete" again
- The mail inviting a maintainer says how to turn on the "Maintainer" tab

# 5.0.1, 2026-November-1
------------------------
- A refused link's message no longer says CRT opens nothing else, and says what to write instead
- A submission no longer reports the component images as reordered for an oscilloscope baseline standing among the pictures
- The line a maintainer reads says "important KiCad signals", as CRT names them

# 5.0.0, 2026-November-1
------------------------
- Numbered 4.3.0 until it was renumbered; never deployed under that number
- A row of the board's own lists with a field it needs left empty is refused: a file with no category or name, a link with no category, name or address, a credit with no category or name or handle, an important signal with no display name or KiCad net name
- A submission's credits and important signals are checked now, as its files and links were
- A change of order is a change: the review says which lists a submission put in another order, and the board's history keeps it
- A schematic image's KiCad traces can have a color and opacities of their own, published as sent

# 4.2.0, 2026-October-29
------------------------
- Oscilloscope models can be sent in from CRT, with a contact address or the signed-in account
- At most 10 models a day per internet connection, and at most 100 waiting at once from everybody together
- A newer model of the same name from the same contributor replaces one still waiting
- The administrators are mailed of each model sent in, and the administrator decides on it: "Publish" writes it into the BETA and the stable lists of oscilloscopes at once, in place of the model of its name; "Reject" needs a reason; the contributor is mailed either way
- A model the stable source could not take stays listed as in BETA, and publishing it again finishes it; it cannot be rejected
- The models waiting are counted for the administrator's badge on the "Maintainer" tab
- Resetting the contribution data counts, shows and deletes the models too
- A sign-in no longer fails when the name of the program signing in is cut to fit in the middle of a character such as an emoji

# 4.1.2, 2026-October-26
------------------------
- A highlight drawn between whole pixels is published rounded, instead of stopping the publish
- An approval reads the published board once instead of some five times, and its KiCad calibrations once instead of once per schematic image

# 4.1.1, 2026-October-25
------------------------
- A publish no longer stops on a file something holds open for a moment, such as a reader or a virus scanner - it tries again a few times

# 4.1.0, 2026-October-25
------------------------
- A submission may carry each component's pins and replacement parts; one that does not keeps the published ones
- A component may carry a package and repair notes, an oscilloscope baseline a machine state and a source
- They are published into the board file, compared in the review's summary and the board's history, and shown and amended in the "Maintainer" tab's table
- Nothing is refused that was taken before: the new checks are warnings only

# 4.0.2, 2026-October-25
------------------------
- A board whose "KiCad data" folder holds a file CRT does not read can be changed again; such a file is taken only as the published board's own bytes
- A queue row of a large board answers in a moment instead of a pause

# 4.0.1, 2026-October-25
------------------------
- A field the server does not know is kept on every row of a board file, not only at its top, so what a later CRT adds survives a publish
- A BETA board file that cannot be read in full is no longer written over: the publish stops with words, and the board is as it was
- A second systemd unit ships, for the test server CRT 4.0's pre-releases use

# 4.0.0, 2026-October-25
------------------------
- The server publishes into the CRT 4.0 data tree - board data as JSON, every board naming its own files, each file's bytes stored once - and no longer writes the Excel generations, which stay as they are for CRT 2.x and 3.x
- API version 2: a file is named by the board's name for it, not by its place in the tree, so every CRT 3.x is told "CRT has to be updated" when it submits or reviews; the check-in, feedback and board views go on as before
- A submission takes from the tree any bytes it already holds, whatever board named them
- Nothing waits for the administrator's second approval any more, as no file is shared between boards
- A publish to BETA or to stable, a push-back and a board's deletion remove the stored files no board lists any more, and deleting a board is no longer refused for a file another board uses
- Not to be deployed on the live service before CRT 4.0 is released

# 3.4.4, 2026-October-24
------------------------
- The server's program can convert an Excel-era data tree into the CRT 4.0 tree, as a command of its own: it only reads the source, and writes nothing when it finds anything it cannot carry
- Nothing changed in what the server answers

# 3.4.3, 2026-October-16
------------------------
- A feedback no longer fails on picture widths that are not valid text; they are ignored
- The feedback mail keeps to its 20 MB, the sender's words counted as the mail draws them
- A feedback mail's notes and its list of saved files end in "... and N more" instead of going on for megabytes
- A picture near the edge of the room left may be saved with the files rather than shown

# 3.4.2, 2026-October-16
------------------------
- A picture named on several lines of a feedback is shown once, at its first line, and carried in the mail once however often it is shown
- A line naming an attached file that is no picture leaves it an ordinary attachment, instead of a note that a picture could not be shown

# 3.4.1, 2026-October-16
------------------------
- Every mail's quotation - a feedback's words and pictures, a contributor's description, a maintainer's comment - is a light blue-grey panel, the same in every mail program
- A feedback of pictures alone shows them inside that panel, and no mail program folds a feedback's words away as quoted text
- The feedback mail is a little wider, so the widest picture CRT shows still fits

# 3.4.0, 2026-October-15
------------------------
- A feedback's pictures are shown in its mail at the width their sender gave them in CRT, never wider than the quotation
- A feedback mail may be up to 20 MB, its pictures sharing what the text leaves; a picture that does not fit is saved with the files
- The mail system on the box must accept a mail that large - postfix's own limit is 10 MB

# 3.3.1, 2026-October-15
------------------------
- A feedback's words and pictures are one quotation in its mail, each picture where its line sat
- A picture the mail shows is no longer saved in the feedback folder as well; one it cannot show is still saved, and the mail says why
- A feedback of words and pictures alone leaves no folder, so its mail has no internal reference or list of files

# 3.3.0, 2026-October-15
------------------------
- The pictures pasted or dropped into a feedback are shown in its mail, between the words above and below where each sat
- Up to 2 MB of pictures a mail; one past that, or one that is not the image its name says, stays a line of text, and the mail says why
- Every picture is saved with the files as before, and a feedback without pictures is mailed exactly as before

# 3.2.0, 2026-October-14
------------------------
- Asking about a submission has an allowance of its own, 20,000 requests per address an hour, apart from creating and uploading, which keep 5,000
- "Rebuild checksum manifests" reads every file again, so a file copied in by hand that kept its old length and time gets its new checksum
- An approval makes the contributor the board's maintainer only when neither BETA nor the stable source holds the board
- A queue row says a submission was made from an older revision exactly when its detail does
- The server needs the ASP.NET Core runtime 10.0.13 or later, for the security fixes in two packages it uses

# 3.1.0, 2026-October-13
------------------------
- Numbered 2.2.0 until it was renumbered the same day; never deployed under that number
- A new board's creator becomes its maintainer when its first publish to BETA is approved: sent signed in, the account joins the board's pool and is mailed; sent with a typed address that has no account, that address is invited
- A typed address that already has an account makes nobody maintainer, and the approver is told; an unverified or locked account is left alone
- The board's pool and the audit name the account that approved it
- An older submission kept because a maintainer amended it or gave the first of two approvals names the same contributor's newer one, still waiting
- A file a submission carries unchanged from the stable source no longer writes the older file back over one BETA has replaced since

# 3.0.0, 2026-October-13
------------------------
- Numbered 2.1.0 until it was renumbered the same day; never deployed under that number
- Finalising a submission after its upload window is refused, as uploading already was
- The contributor routes are limited to 5,000 requests per address an hour, and a board id with a control character is refused
- A queue row and a submission's detail say whether it was amended and whether it was made from an older revision than BETA has now
- A decision is refused when another maintainer amended the submission since it was looked at, and an approval's comment goes into the contributor's mail
- A replaced shared file names the other boards using it
- "Reject" and "Request changes" no longer overwrite a publish that got there first, and a finalise sent again no longer mails the maintainers again
- A publish, a push to stable and a push-back are no longer stopped half way when CRT goes away, and nothing after a written publish can fail it
- An approval leaves a file the contributor sent unchanged as it is published now, so an old submission no longer reverts a newer shared file
- Five wrong passwords lock out one address for an account, not the account for everybody; the account itself locks after 25 in 15 minutes
- A mail the mail system refuses is tried again for three days, and the server's status tells signed-in maintainers how many wait to be sent
- Refusals name folders from the data's root, never the server's own paths, and say "the stable source" where they said "production"
- A submission's change summary matches rows ignoring capitals and compares values trimmed, as the contributor's own count does
- A component relabelled in a maintainer's table keeps its highlights under the new label, and a highlight opacity that cannot be read is a warning

# 2.0.0, 2026-October-11
------------------------
- A KiCad schematic naming a sheet outside its board's "KiCad data" folder - by a full path, a network share, a drive letter or climbing out with ".." - is refused, at finalise and again before a publish
- So is a schematic whose sheet names a file that is not a KiCad schematic, or names it by a path longer than 200 characters
- Every shipped board passes, and a real KiCad project is not affected
- Login and registration take the same time whether or not the address has an account
- No more password checks run at once than the machine has cores for; a login, registration or password change waits its turn
- A finding's message longer than 16,000 characters is cut to fit, instead of failing the submission

# 1.1.2, 2026-October-11
------------------------
- A maintainer's table change moves a renamed schematic image's highlights and KiCad calibration with it, and drops a deleted one's, by the same rule a draft uses - every field kept
- Nothing changed in what the server answers

# 1.1.1, 2026-October-11
------------------------
- A board's KiCad calibrations are read in one go instead of once per schematic image - the same calibrations, faster on a board with many images
- Nothing changed in what the server answers

# 1.1.0, 2026-October-10
------------------------
- A publish to BETA records every schematic image it renames - the image file, its old and new name and the date - beside the board's data, and the push to stable copies it
- CRT follows such a rename in its user's own worklogs and settings
- A board nobody renamed keeps exactly the file it had

# 1.0.3, 2026-October-10
------------------------
- A board's KiCad calibration is found when its schematic image's name differs only in capitals, as every other schematic name already was
- Saving one under new capitals replaces the old one

# 1.0.2, 2026-October-10
------------------------
- The contributor's mail saying a submission is ready to test in BETA words its encouragement anew, in the project owner's words

# 1.0.1, 2026-October-10
------------------------
- The reason a board's table cannot be changed while the account's earlier change of it still waits is written once, so the "Maintainer" tab says the same words when it cannot read the table again
- Nothing changed in what the server answers

# 1.0.0, 2026-October-9
-----------------------
- The numbering starts again at 1.0.0: the contribution data was reset again, and nothing had been released to anybody but the project owner
- Nothing changed: this is the server as the earlier numbering left it at 2.1.0, with the API version back at 1, which now stands for the API with "board" everywhere
- A CRT built before this day, for the old API version 1 with "system" in its routes, is no longer told to update and fails instead; none was ever given to anybody else
