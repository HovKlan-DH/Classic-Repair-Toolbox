[Wiki Home](Home)

Review and publish the changes other people send in - for maintainers of the hardware data only.

---

The **Maintainer** tab is where a maintainer looks through the contributions sent in for the systems they look after, and publishes the good ones for everybody who uses CRT. It is hidden unless you turn it on, because it needs a maintainer account, and almost nobody using CRT has one - you can use every other part of CRT, and contribute data yourself, without it.

## Turning it on

In the [Configuration](Configuration-tab) tab, under **Contributing and maintaining data**, tick **Enable Maintainer tab**. The tab then appears in the row of tabs, just before Configuration. Untick it again to hide it; that does not sign you out.

Tick **Hide the Maintainer tab while no work is waiting for me** as well, and the tab only appears while its red number (below) shows something to do. A tab you are on, or one with unsaved changes in a table, stays until you leave it.

While the Maintainer tab is selected, the hardware and board list on the left and the worklog bar are put away, so its screens get the whole width of the window. They come back as soon as you choose another tab.

Once you are signed in, a red number on the tab tells you how much is waiting for you: the systems with a contribution for you to review, plus the systems in BETA waiting for you to publish them. It is the two numbers on the **Queue: Contributor submissions** and **Queue: Awaiting push from BETA to stable** tabs added up, and it goes away when nothing is left for you. CRT checks for new work every minute while its window is open, even when you are on another tab.

## Getting a maintainer account

You cannot create a maintainer account yourself. The administrator invites you by email, for the system or systems you will look after. The email holds an invitation code.

To accept it, open the Maintainer tab and choose **I have an invitation**. The sign-in boxes make way for the invitation's own: paste the code, pick the name other maintainers and contributors will see and a password, and choose **Create my account**. Your account is made for the address the invitation was sent to, and you are taken back to signing in with that address and your password already filled in. **Cancel** takes you back to signing in without creating anything.

## Signing in

Sign in with your email address and password. If you have forgotten your password, choose **I forgot my password**: you are sent a code by email, which you paste into the tab together with the new password you want.

On Windows, CRT remembers that you are signed in, and signs you back in by itself when it starts, so the number on the tab is there without opening the tab first. What it keeps is encrypted to your Windows user account and cannot be used from another computer or another user. On Linux and macOS there is no such safe place to keep it, so you sign in each time you open CRT.

With the tab turned on, you stay signed in for as long as you open CRT at least once every 30 days.

Under the list on the left, the tab always says who is logged in - your name in bold and your email address - so you can see which account you are acting as. **Sign out** is under **Account** > **My account** (below); it signs you out and forgets the remembered sign-in on this computer.

**While you are signed in, CRT uses your account's email address everywhere it asks for one.** The [Feedback tab](Feedback-tab) shows it, and so does the dialog that submits a draft from the Drafts tab - which also sends the submission with your account, so the maintainer reviewing it sees that the address is checked. Neither box can be changed while you are signed in - whatever the Configuration tab says, even with the Maintainer tab turned off or hidden; sign out to use another address, or change your account's address (below). CRT remembers the sign-in between launches until you sign out.

## My account

**Account** > **My account** is where you change your own details, and sign out:

- **Name or handle** - what other maintainers see, and what the emails call you. Type the new one and choose **Save name**.
- **Email address** - type the new email address and choose **Send code**. A code is sent to the new email address; when it arrives, choose **I have received a code**, type the code in and choose **Change address**. Your address does not change until the code is used, so a mistyped address cannot lock you out. You can go to another screen, or close CRT, in between - the code works for 24 hours. Your old address then gets an email saying the address was changed.
- **Password** - type the new password twice and choose **Change password**. You get an email saying it was changed.
- **Sign out** - signs you out and forgets the remembered sign-in on this computer. If a table holds changes you have not saved, you are asked about them first.

You are already signed in, so none of these asks for your current password.

When you change your email address or your password, every other computer where you are signed in is signed out - its Maintainer tab asks you to sign in again. The computer you made the change on stays signed in, and the Feedback tab and the submit dialog use the new address straight away. A new name shows on your other computers the next time CRT starts there with the Maintainer tab turned on.

## The four screens

The row of tabs along the top chooses what the Maintainer tab shows. Each is a list on the left and the item you choose in it on the right. Moving between them never closes anything, so a submission you were looking at is still there when you come back.

When you open the Maintainer tab and nothing is waiting for you in either queue, it goes to **Systems** and shows the system you looked at last - even after closing CRT - or otherwise the first one in the list. A submission that waits only for the other approver does not count as waiting for you. If something is open on a queue already - a submission you were looking at - the tab stays where you left it.

**Queue: Contributor submissions** and **Queue: Awaiting push from BETA to stable** open straight on something, so you do not have to pick from the list first: the entry you looked at last, if it is still there - even after closing CRT - or otherwise the first one in the list.

**Systems** - every system in the data. Under each name the list says how many maintain it, and otherwise only what needs attention: not published yet, not in the stable source yet, BETA ahead of the stable source, in the stable source but not in BETA, nobody assigned, closed to contributions - and anything off with CRT's hardware and board lists: a system listed only in the stable source's list and not in BETA's, or a system the stable source has that is missing from its list. A system that is in BETA and the stable source, as it should be, says nothing more. Choose a system, and under its name a line shows where it is, in three steps: **Submitted** (its newest submission, how it went, and when), **BETA** and **Stable** (its revision in each, or "Not there"). A system that was contributed but never published shows here why - for example its submission was not accepted, and neither BETA nor stable has it. The switch below chooses how you look at the system:

- **Board data** and **Files** have a **Data: BETA | Stable** switch above them. **BETA** is what the views describe below; **Stable** shows the system as the stable data holds it - what everybody using CRT gets - and can only be looked at, never changed there: a change is made on BETA's table and reaches stable when BETA is published to it. A choice the system is not in cannot be picked. Moving between BETA and Stable never loses a change you are making in BETA's table.
- **Board data** - the system's board as the BETA data holds it now, as a table - the same table a submission opens on, with the same checks, colour key and search. Nothing is coloured until you change something. If you maintain the system, you can correct it here: make your change and choose **Save changes**. CRT asks for a reason for the change - a sentence is enough, and it goes along with the change just like a contribution's description. If the change means some files are no longer used by anything, they are listed in the same window, because publishing removes them. Choose **Publish to BETA**, and the change goes straight into the BETA data, where you can try it in CRT with "Download data from the BETA source instead of the stable source" ticked in the Configuration tab. It then waits under **Queue: Awaiting push from BETA to stable** until it is published to the stable source.
  - While a system waits under **Queue: Awaiting push from BETA to stable**, the table here only lets you look - nobody can change it, you included. It has to be published to stable, pushed back or rejected first.
  - If publishing is refused - because BETA changed since you opened the table, say - your change stays in the table with the reason, and the next **Save changes** asks with the reason you typed already filled in.
  - The table follows BETA: when something is published to BETA for the system, pushed back or published to stable, the table (and the Files view) is read again by itself, so it never shows something BETA no longer has. The table is also read again when who maintains the system changes, since that decides who may change it. If you are in the middle of a change then, the table is left as it is and says the change can no longer be published.
  - Should the change not be published after all - because something else changed in BETA at the same moment - nothing is lost: it waits as a submission under **Queue: Contributor submissions**, CRT takes you there, and you can approve it like any other. While it waits, the table here only lets you look - decide that submission first, or make the new change in its own table.
  - If you do not maintain the system, you can look and search, but not change anything.
- **Files** - every file the system uses in BETA, as a folder tree: its own folder, and the shared files its board uses, each with its size. Point at a file to see it, and double-click it to open it.
- **Contributor** - everybody who has contributed to the system, with how many contributions were accepted, are waiting, were sent back for changes or were rejected, and when the last one came.
- **Maintainer** - who maintains the system. A new system also gets its place in CRT's hardware and board lists here, under "Place it in the drop-down lists" - its hardware name, board name and hardware notes, and where it goes in the list: drag it there and choose **Save its place**. The hardware notes start with whatever the contributor wrote in "Create system"; what you save is what CRT shows on the Overview tab. It must have a place before it can be approved - the line under its name says so until it has one. Who maintains a system is changed by the administrator, under **Account** > **Maintainers**.
- **History** - everything that has happened to the system, newest first, grouped by month. Each submission is one card: what the contributor wrote, where it stands now, when it was sent, changed by a maintainer and decided, what the contributor was told - and what it changed when it went into BETA: for each sheet, which rows were added, changed (and in which columns), removed or renamed, and which files were added, replaced or removed. Everything else - a publish to the stable source, a push back, a maintainer added or removed - is a line of its own. A submission published to BETA before CRT kept this record says that what it changed was not recorded.
- **Statistics** - how often the board is looked at in CRT.

Email addresses are shown only for the systems you maintain yourself (the administrator sees them on every system). On every other system you see the same views, with people named by the name on their account - a contributor who sent something without an account is shown as "A contributor without an account" - and the Contributor and Maintainer views say that the addresses are left out.

The view you pick stays picked when you choose another system. Choosing another system while the table holds a change you have not published asks first whether to publish it (with a reason, as above) or throw it away.

**Queue: Contributor submissions** - the queue of changes waiting for review, grouped by board. The number on its tab counts the systems waiting for you, and the queue checks for new submissions by itself every minute. Choose a submission, and the switch above it chooses how you look at it:

- **Board data** - the submission's changes as a table, one tab per sheet, with what changed coloured in. This is what a submission opens on. You can correct things in the table yourself before deciding. Above the table are the few changes a table cannot show - the component highlights on the schematics and the KiCad calibration points, each kind under a heading that counts them ("Component highlights have 1 change:"), one change a line (*Removed component [hest] from schematic "Board layout"*) - and any warnings from the automatic checks. The table is the same one contributors edit on the Drafts tab, with the same checks: a cell with a warning has an amber corner (hover it to read it), and clicking a count in the colour key - "Modified", "Warnings" and so on - shows only those rows, until you click it again. CRT remembers which counts you clicked for the next submission. The search box above the table finds text in it the same way, Shift+click or Ctrl+click selects several rows to delete in one go, and a long text is always shown in full, over several lines - double-click the edge of a column's heading to fit the column to its text. A row whose name alone changed - a credit's name, a component's board label - is shown as that row changed, not as one row deleted and another added.
- **Files** - every file of the system as a folder tree, as the BETA data will be after approving the submission, with new, changed and removed files marked and each file's size. The number on the button counts the files the submission adds, replaces or removes, so you can see there is something to look at before you open it. This matters most for a file replaced under the same name, which does not show as a change in the table at all. Point at a file to see it, and double-click it to open it.
- **Contributor** - who sent it, whether it came from an account (so the email address is checked) or with an address simply typed in, how many other submissions the contributor has sent and how many of those were published to the stable source, published to BETA, are waiting or were rejected, and then each of those submissions: what it was, which system, how it went and what the contributor was told.

Then decide, with the buttons under the views: **Approve and publish to BETA** publishes it to the BETA data, for checking; **Request changes** sends it back to the contributor; **Reject** turns it down. **Request changes** and **Reject** need a comment of at least 10 characters in the box above the buttons - it is the only message the contributor receives. Approving sends no comment. A decision waits until changes you made in the table are saved. The decision buttons are there whichever of the three you are looking at, and moving between them never loses anything you changed in the table. If **Approve** is greyed out for a reason of its own - a new system with no place in the lists yet, or an earlier submission of the system still waiting in BETA - the reason is said above the switch, as is a warning when the contributor has since discarded the draft, so you see them whichever view is open.

A submission that replaces a shared file other boards use needs two approvals: a maintainer of the board AND the administrator. The line above the buttons says so and who has approved so far. The first approval only records it - the button then reads, for example, "Approve - the administrator must approve too", and the contributor sees "Approved, waiting to be published" - and the second one publishes it to BETA.

**Queue: Awaiting push from BETA to stable** - systems whose approved changes are in the BETA data and waiting to go out to everybody. Choose a system to see the contributions it carries - who sent each one, what it was and when it was accepted - then anything that needs a second look, and the files that would be copied. Try the board in CRT with "Download data from the BETA source instead of the stable source" ticked in the Configuration tab, then:

- **Publish to stable** - copies the system to the stable source. Tick "I have checked this board in CRT with the BETA data (Configuration: "Download data from the BETA source instead of the stable source"), and it is right." first.
- **Push back to queue** - puts the system's BETA data back to what the stable source has (a system never published to stable is taken out of BETA), and every submission published to BETA since then goes back to **Queue: Contributor submissions**. You are asked why, and every contributor named is told exactly that.
- **Reject** - the same, but those submissions are rejected instead of going back to the queue.

For now, only the administrator publishes to the stable source: for everybody else the publish button is greyed out and says so, each system's line in the list ends "with the administrator", and these systems do not count in your red number - but you can still look through them, push them back or reject them.

**Account** - for every maintainer:

- **My account** - your name, email address and password, and signing out - see [My account](#my-account) above. The screen opens on it.
- **Server version** - what is deployed on the server, as three lines: the server version, the API
  server version and the API application version. The two API versions are the API revision - a
  number that changes only when the server changes in a way an older CRT cannot follow - of the
  server and of this CRT. When they are the same, the two work together. When the server's is
  higher, this CRT is behind and has to be updated: the server turns it away from submitting and
  from this tab, saying so. When this CRT's is higher, the newer server has not been deployed yet.
  The server is asked every time you open this, so after an update the new version shows straight
  away.

Below those, the administrator also gets seven tools, each marked with a padlock. They are only for the administrator - nobody else sees them in the list, and the server turns everybody else away from them anyway:

- **Maintainers** - who maintains each system. Choose a system to see its maintainers and the invitations nobody has used yet. Choose somebody who already has an account from the list and choose **Add as maintainer**, or type an email address and choose **Send invitation** to invite somebody new - an invitation code is mailed to that address. **Remove** takes a maintainer off the system, and **Withdraw** cancels an invitation. A system can have as many maintainers as you like. You can add yourself too, so other maintainers can see that you look after a system: you then get that system's emails about new submissions like its other maintainers - one email, not one as administrator and another as maintainer - and nothing else changes, as you can already review and publish every system.
- **Order of systems** - the order CRT shows the hardware and boards in its drop-down lists. Drag a system up or down, then choose **Save the order**. The order is written into the hardware and board list of both the BETA and the stable data, and CRT picks it up the next time it starts. **Cancel** drops the moves you have not saved.
- **Unused files** - files in the data that nothing uses any more, in the BETA or the stable data, shown as the same folder tree as a system's files, each with its size: point at a file to see it, and double-click it to open it. A file is listed only when no hardware and board list names it, no board's data file of any version - the older ones kept for older CRT versions included - cites it, and it is not in a "KiCad data" folder, the MiniPro IC tests or a file whose name starts with "!". If any data file cannot be read, nothing is listed at all. The BETA and the stable data are each checked on their own - pick one under **Data:** - so a file listed for BETA may still be used by the stable data, and the other way round. To delete the listed files from the server, tick "I have looked through this list, and every file on it can be removed from the server." and choose the remove button, which says how many files and how much space it removes. **Refresh** reads the list again.
- **Rebuild checksum manifests** - **Rebuild both manifests** rebuilds `dataChecksums.json` for both
  the BETA and the stable data. The server does this by itself whenever it publishes, promotes or rolls something back, so
  this is only needed when data files have been changed on the server by hand, which the service
  never sees - without a rebuild, those changes reach nobody. It only reads the data and replaces
  one file, so it is safe to press again if something fails.
- **Delete a system** - removes a system from everywhere: its folder in the BETA and the stable
  data, its place in CRT's hardware and board lists, and everything the server knows about it - its
  submissions, maintainers, invitations and view statistics. Every system is listed with a
  **Delete** button. Pressing it first shows exactly what would go, and nothing is deleted until you
  confirm. If a submission to the system is still open - waiting for review, or in BETA - it is
  deleted too, and its contributor gets an email with the reason you give, so a reason is asked for
  then. Shared files the system used are kept; any that nothing else uses then show under **Unused
  files**. A deletion cannot be undone. Some systems cannot be deleted, and CRT says why instead of
  asking: one still listed in an older hardware and board list, which is kept unchanged for older
  CRT versions, or one with a file that another board also uses. If a deletion stops part-way, press **Delete**
  again to finish it.
- **API usage** - which CRT versions still use each part of the server, for deciding when an old part
  can be retired. Every route the server has is listed - the ones nobody used first in each group -
  with the CRT versions that called it in the last 30 days, 90 days or year, and, at the top, how
  many installations started each CRT version. The server counts the calls per day and CRT version
  only, with no address and no account. The routes every CRT ever released uses - the launch
  check-in, feedback, board views - are listed last and are never retired. A retired route is not
  simply removed: CRT versions still calling it are told to update.
- **Reset contribution data** - for going live: deletes everything sent and done through the
  contribution service while it was being tested. That is every submission and its uploaded files,
  every account except the administrators' - so every maintainer and invitation too - the history,
  the board views and the API usage counts. The BETA and the stable data are not touched. It first
  shows how much of each would go; type `RESET` and press the red button to do it. It cannot be
  undone, and it works only while the server has been set to allow it - until then the tool says how
  to switch that on. Contributors whose submissions are deleted see them as "No longer on the
  server" in CRT, and can send them again.

If you close CRT with changes in a table that you have not saved, you are asked whether to save them first.
