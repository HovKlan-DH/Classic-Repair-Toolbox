[Wiki Home](Home)

Review and publish the changes other people send in - for maintainers of the hardware data only.

---

The **Maintainer** tab is where a maintainer looks through the contributions sent in for the boards they look after, and publishes the good ones for everybody who uses CRT. It is hidden unless you turn it on, because it needs a maintainer account, and almost nobody using CRT has one - you can use every other part of CRT, and contribute data yourself, without it.

## Turning it on

In the [Configuration](Configuration-tab) tab, under **Contributing and maintaining data**, tick **Enable Maintainer tab**. The tab then appears in the row of tabs, just before Configuration. Untick it again to hide it; that does not sign you out.

Tick **Hide the Maintainer tab while no work is waiting for me** as well, and the tab only appears while its red number (below) shows something to do. A tab you are on, or one with unsaved changes in a table, stays until you leave it.

While the Maintainer tab is selected, the hardware and board list on the left and the worklog bar are put away, so its screens get the whole width of the window. They come back as soon as you choose another tab.

Once you are signed in, a red number on the tab tells you how much is waiting for you: the boards with a contribution for you to review, plus the boards in BETA waiting for you to publish them. It is the two numbers on the **Queue: Contributor submissions** and **Queue: Awaiting push from BETA to stable** tabs added up, and it goes away when nothing is left for you. CRT checks for new work every minute while its window is open, even when you are on another tab.

## Getting a maintainer account

You cannot create a maintainer account yourself. The administrator invites you by email, for the board or boards you will look after - or, if you created a new board yourself, the invitation comes by itself once that board is accepted. The email holds an invitation code.

To accept it, open the Maintainer tab and choose **I have an invitation**. The sign-in boxes make way for the invitation's own: paste the code, pick the name other maintainers and contributors will see and a password, and choose **Create my account**. Your account is made for the address the invitation was sent to, and you are taken back to signing in with that address and your password already filled in. **Cancel** takes you back to signing in without creating anything.

## Signing in

Sign in with your email address and password. If you have forgotten your password, choose **I forgot my password**: you are sent a code by email, which you paste into the tab together with the new password you want.

On Windows, CRT remembers that you are signed in, and signs you back in by itself when it starts, so the number on the tab is there without opening the tab first. What it keeps is encrypted to your Windows user account and cannot be used from another computer or another user. On Linux and macOS there is no such safe place to keep it, so you sign in each time you open CRT - the sign-in screen says so under **Sign in**.

With the tab turned on, CRT keeps in touch with the server while it runs, and that keeps you signed in: open CRT at least once every two weeks and you are not asked again. Leave it longer, and you may be asked to sign in again - at the latest 30 days after CRT last used your sign-in.

If the server ends your sign-in while a table holds a change you have not saved - you signed out on another computer, say - the tab does not throw the change away: it stops checking for new work, and the line above the list says the change can no longer be saved. Copy anything you need, then sign out and in again to carry on.

Under the list on the left, the tab always says who is signed in - **Signed in as:**, then your name in bold and your email address - so you can see which account you are acting as. **Sign out** is under **Account** > **My account** (below); it signs you out and forgets the remembered sign-in on this computer.

**While you are signed in, CRT uses your account's email address everywhere it asks for one.** The [Feedback tab](Feedback-tab) shows it, and so does the dialog that submits a draft from the Drafts tab - which also sends the submission with your account, so the maintainer reviewing it sees that the address is checked. Neither box can be changed while you are signed in - whatever the Configuration tab says, even with the Maintainer tab turned off or hidden; sign out to use another address, or change your account's address (below). CRT remembers the sign-in between launches until you sign out.

## My account

**Account** > **My account** is where you change your own details, and sign out:

- **Name or handle** - what other maintainers see, and what the emails call you. Type the new one and choose **Save name**.
- **Email address** - type the new email address and choose **Send code**. A code is sent to the new email address; when it arrives, choose **I have received a code**, type the code in and choose **Change address**. Your address does not change until the code is used, so a mistyped address cannot lock you out. You can go to another screen, or close CRT, in between - the code works for 24 hours. Your old address then gets an email saying the address was changed.
- **Password** - type the new password twice and choose **Change password**. You get an email saying it was changed.
- **Sign out** - signs you out and forgets the remembered sign-in on this computer. If a table holds changes you have not saved, you are asked about them first.

You are already signed in, so none of these asks for your current password.

When you change your email address or your password, every other computer where you are signed in is signed out - its Maintainer tab asks you to sign in again. The computer you made the change on stays signed in, and the Feedback tab and the submit dialog use the new address straight away. A new name shows on your other computers the next time CRT starts there with the Maintainer tab turned on.

## When CRT has to be updated

Now and then the server changes in a way an older CRT cannot follow, or stops serving CRT versions older than a given one. While the Maintainer tab is turned on (or you have drafts), CRT asks the server about this when it starts - or when you turn the tab on - and it notices it whenever the server answers that way. The whole Maintainer tab is then covered by **CRT has to be updated**, which cannot be closed - the tab cannot be used until CRT is updated, and stops asking the server for new work. You are not signed out. Everything else in CRT works as before.

If a table on the tab holds a change you have not sent, installing the update - or quitting CRT - asks first. It offers only **Discard edits** or **Cancel**, because the change cannot be sent until CRT is updated: cancel to keep CRT open and copy what you need.

The button on it is the way out. If CRT has already found a newer version - with **Check for new version at application launch** ticked on the [Configuration tab](Configuration-tab) - **Install update** downloads and installs it, as the update notice does. Otherwise **Open download page** opens the page listing every CRT release in your web browser.

## The four screens

The row of tabs along the top chooses what the Maintainer tab shows. Each is a list on the left and the item you choose in it on the right. Moving between them never closes anything, so a submission you were looking at is still there when you come back.

When you open the Maintainer tab and nothing is waiting for you in either queue, it goes to **Boards** and shows the board you looked at last - even after closing CRT - or otherwise the first one in the list. A submission that waits only for the other approver does not count as waiting for you. If something is open on a queue already - a submission you were looking at - the tab stays where you left it.

**Queue: Contributor submissions** and **Queue: Awaiting push from BETA to stable** open straight on something, so you do not have to pick from the list first: the entry you looked at last, if it is still there - even after closing CRT - or otherwise the first one in the list.

**Boards** - every board in the data. Under each name the list says how many maintain it, and otherwise only what needs attention: not published yet, not in the stable source yet, BETA ahead of the stable source, in the stable source but not in BETA, nobody assigned, closed to contributions - and anything off with CRT's hardware and board lists: a board listed only in the stable source's list and not in BETA's, or a board the stable source has that is missing from its list. A board that is in BETA and the stable source, as it should be, says nothing more. A last line says how the board's submissions went, for example "19 submissions in total; 2 rejected, 1 in BETA, 16 in stable" - and "1 waiting" too while one waits for review or for changes. A submission replaced by its contributor's newer one is not counted, and a rejection only when a maintainer made it.

Choose a board, and the switch under its name chooses how you look at it:

- **Board data** and **Files** have a **Data source: BETA | Stable** switch above them for a board that is in both. **BETA** is what the views describe below; **Stable** shows the board as the stable data holds it - what everybody using CRT gets - and can only be looked at, never changed there: a change is made on BETA's table and reaches stable when BETA is published to it. A board only one of them holds shows that one, with no switch. Moving between BETA and Stable never loses a change you are making in BETA's table.
- **Compare sources**, right after the two buttons, colours **Board data**'s table against the other source, the way a submission's table is coloured. With **BETA** chosen, everything BETA holds differently from the stable source is marked - green added, orange changed, red deleted - and pointing at a changed cell shows its **Stable source value**. With **Stable** chosen it is the same the other way round, with the **BETA source value**. Pointing at a changed picture or file shows both sources' versions. Like the switch, it is there only for a board in both sources, and while BETA's table holds a change you have not saved, the box cannot be changed - save or undo your change first. A board's table you can only look at says above it when it is compared with the stable source. If the stable source's board cannot be read, BETA's table says above it that it is not compared, and why - choose another view and come back to **Board data** to try again. CRT remembers the tick for the next time.
- **Board data** - the board's data as the BETA data holds it now, as a table - the same table a submission opens on, with the same checks, colour key and search. Unless **Compare sources** is ticked, nothing is coloured until you change something. If you maintain the board, you can correct it here: make your change and choose **Save changes**. CRT asks for a reason for the change - a sentence is enough, and it goes along with the change just like a contribution's description. If the change means some files are no longer used by anything, they are listed in the same window, because publishing removes them. Choose **Publish to BETA**, and the change goes straight into the BETA data, where you can try it in CRT with "Download data from the BETA source instead of the stable source" ticked in the Configuration tab. It then waits under **Queue: Awaiting push from BETA to stable** until it is published to the stable source.
  - While a board waits under **Queue: Awaiting push from BETA to stable**, the table here only lets you look - nobody can change it, you included. It has to be published to stable, pushed back or rejected first. An amber panel above the board's views says so, whichever view is open - from the moment the board is chosen - as it says why for a board you do not maintain, in the same place and look as the queue's reason why **Approve** is greyed out.
  - If publishing is refused - because BETA changed since you opened the table, say - your change stays in the table with the reason, and the next **Save changes** asks with the reason you typed already filled in.
  - The table follows BETA: when something is published to BETA for the board, pushed back or published to stable, the table (and the Files view) is read again by itself, so it never shows something BETA no longer has. The table is also read again when who maintains the board changes, since that decides who may change it. If you are in the middle of a change then, the table is left as it is and says the change can no longer be published.
  - Should the change not be published after all - because something else changed in BETA at the same moment - nothing is lost: it waits as a submission under **Queue: Contributor submissions**, CRT takes you there, and you can approve it like any other. While it waits, the table here only lets you look - decide that submission first, or make the new change in its own table.
  - If you do not maintain the board, you can look and search, but not change anything.
- **Files** - every file the board uses in BETA, as a folder tree: its own folder, and the shared files it uses, each with its size. Point at a file to see it - a picture has a thin frame round it and its size in pixels under it - and double-click it to open it.
- **Contributor** - everybody who has contributed to the board, with how many contributions were accepted, are waiting, were sent back for changes or were rejected, and when the last one came.
- **Maintainer** - who maintains the board. A new board also gets its place in CRT's hardware and board lists here, by the administrator or a maintainer of the board, under "Place it in the drop-down lists" - its hardware name, board name and hardware notes, and where it goes in the list: drag it there and choose **Save its place**. The hardware notes start with whatever the contributor wrote in "Create board"; what you save is what CRT shows on the Overview tab. It must have a place before it can be approved - the line under its name says so until it has one. Who maintains a board is changed by the administrator, under **Account** > **Maintainers**.
- **History** - everything that has happened to the board, newest first, grouped by month. Maintainers and contributors are named in bold, by the name on their account, or by their email address when there is no name. Each submission is one card: what the contributor wrote, where it stands now, when it was sent, changed by a maintainer and decided, what the contributor was told - and what it changed when it went into BETA: for each sheet, which rows were added, changed (and in which columns), removed or renamed, and which files were added, replaced or removed. Everything else - a publish to the stable source, a push back, a maintainer added or removed - is a line of its own: for example **Maintainer added: Anna**, with **by Dennis** under it. A submission published to BETA before CRT kept this record says that what it changed was not recorded.
- **Statistics** - how often the board is looked at in CRT, and a graph of its **views per day**. Choose **30 days**, **90 days** or **365 days** above the graph, and point at a day to see its date and how many views it had. Days are UTC days, and views from CRTs downloading the BETA source are left out, as from the counts above it.

Email addresses are shown only for the boards you maintain yourself (the administrator sees them on every board). On every other board you see the same views, with people named by the name on their account - a contributor who sent something without an account is shown as "A contributor without an account" - and the Contributor and Maintainer views say that the addresses are left out.

The view you pick stays picked when you choose another board. Choosing another board while the table holds a change you have not published asks first whether to publish it (with a reason, as above) or throw it away.

**Queue: Contributor submissions** - the queue of changes waiting for review, grouped by board. The number on its tab counts the boards waiting for you, and the queue checks for new submissions by itself every minute. Choose a submission, and the switch above it chooses how you look at it:

- **Board data** - the submission's changes as a table, one tab per sheet, with what changed coloured in. This is what a submission opens on. You can correct things in the table yourself before deciding. Above the table are the few changes a table cannot show - the component highlights on the schematics and the KiCad calibration points, each kind under a heading that counts them ("Component highlights have 1 change:"), one change a line (*Removed component [hest] from schematic "Board layout"*) - and any warnings from the automatic checks. The table is the same one contributors edit on the Drafts tab, with the same checks: a cell with a warning has an amber corner (hover it to read it), and clicking a count in the colour key - "Modified", "Warnings" and so on - shows only those rows, until you click it again. CRT remembers which counts you clicked for the next submission. The search box above the table finds text in it the same way, Shift+click or Ctrl+click selects several rows to delete in one go, and a long text is always shown in full, over several lines - double-click the edge of a column's heading to fit the column to its text. A row whose name alone changed - a credit's name, a component's board label - is shown as that row changed, not as one row deleted and another added. If your account is not a maintainer of the board, the table only lets you look - the server would not take a change from you - and the line above the buttons says so.
- **Files** - every file of the board as a folder tree, as the BETA data will be after approving the submission, with new, changed and removed files marked and each file's size. The number on the button counts the files the submission adds, replaces or removes, so you can see there is something to look at before you open it. This matters most for a file replaced under the same name, which does not show as a change in the table at all. Point at a file to see it, and double-click it to open it.
- **Contributor** - who sent it, whether it came from an account (so the email address is checked) or with an address simply typed in, how many other submissions the contributor has sent and how many of those were published to the stable source, published to BETA, are waiting or were rejected, and then each of those submissions: what it was, which board, how it went and what the contributor was told.

Then decide, with the buttons under the views: **Approve and publish to BETA** publishes it to the BETA data, for checking; **Request changes** sends it back to the contributor; **Reject** turns it down. **Request changes** and **Reject** need a comment of at least 10 characters in the box above the buttons - it is the only message the contributor receives. A comment is optional for **Approve**: when you write one, it goes into the mail that tells the contributor the submission is in BETA. A decision waits until changes you made in the table are saved - a cell you are still typing in counts too. If another maintainer changed the submission in the table after you opened it, your decision is refused and the submission is read again, so you see the change before you decide - the tab also reads it again by itself when its check notices the change. The decision buttons are there whichever of the three you are looking at, and moving between them never loses anything you changed in the table. If **Approve** is greyed out for a reason of its own - a new board with no place in the lists yet, or an earlier submission of the board still waiting in BETA - the reason is said in an amber panel above the switch - the same panel the **Boards** screen uses. The same panel warns when the contributor has since discarded the draft, and when the submission was made from older data than BETA has now ("This was made from older data than BETA has now - rows published since may show as removed. Compare carefully."): rows published since then show as removed, and approving would remove them. It also warns when the same contributor has sent a newer submission for the board that is waiting too ("The same contributor sent a newer submission for this board on ... (#57), which is waiting too."): a newer submission normally replaces the older one, but not one a maintainer has already changed in the table or given the first of two approvals, so the two wait side by side and approving the older one publishes the older work. You see these whichever view is open.

**Approving a new board makes its creator its maintainer.** A board sent signed in makes that account its maintainer; a board sent with a typed-in address that has no account mails that address an invitation to create one. A typed-in address that already has an account makes nobody maintainer - nothing proves the sender owns it. A board that is only missing from the BETA data - the stable data already has it - is not new, so approving it makes nobody maintainer. The message after approving says which happened, or why nothing could be done (that, or an account that is locked or not verified yet); the administrator can then add the maintainer under **Account**, **Maintainers**.

A submission that replaces a shared file other boards use needs two approvals: a maintainer of the board AND the administrator. The line above the buttons says so, who has approved so far, and which other boards use the file ("Also used by 2 other boards: ...") - every one of them gets the new file. In **Files**, the replaced file says the same beside it. A file the contributor sent unchanged is left as it is: when somebody else has replaced it since the submission was sent, approving does not put the old file back. The first approval only records it - the button then reads, for example, "Approve - the administrator must approve too", and the contributor sees "Approved, waiting to be published" - and the second one publishes it to BETA.

**Queue: Awaiting push from BETA to stable** - boards whose approved changes are in the BETA data and waiting to go out to everybody. Choose a board to see the contributions it carries - who sent each one, what it was and when it was accepted - then anything that needs a second look, and the files that would be copied. If a contributor has discarded the draft since the contribution was accepted, an amber panel right under the board's name says so - the same panel the other screens use. The contribution is not withdrawn, but it may no longer be wanted: consider pushing the board back to the queue and checking with the contributor first. Try the board in CRT with "Download data from the BETA source instead of the stable source" ticked in the Configuration tab, then:

- **Approve and publish to stable** - copies the board to the stable source. Tick "I have checked this board in CRT with the BETA data (Configuration: "Download data from the BETA source instead of the stable source"), and it is right." first.
- **Push back to queue** - puts the board's BETA data back to what the stable source has (a board never published to stable is taken out of BETA), and every submission published to BETA since then goes back to **Queue: Contributor submissions**. You are asked why, and every contributor named is told exactly that. Confirm with **Push back to queue** - or **Push back and remove from BETA** for a board never published to stable. The board's **History** then reads "Pushed back to the queue".
- **Reject** - the same, but those submissions are rejected instead of going back to the queue.

For now, only the administrator publishes to the stable source: for everybody else the publish button is greyed out and says so, each board's line in the list ends "with the administrator", and these boards do not count in your red number - but you can still look through them, push them back or reject them.

**Account** - for every maintainer:

- **My account** - your name, email address and password, and signing out - see [My account](#my-account) above. The screen opens on it.
- **Server version** - what is deployed on the server, as three lines: the server version, the API
  server version and the API application version. The two API versions are the API revision - a
  number that changes only when the server changes in a way an older CRT cannot follow - of the
  server and of this CRT. When they are the same, the two work together. When the server's is
  higher, this CRT is behind and has to be updated: the server turns it away from submitting and
  from this tab, saying so. When this CRT's is higher, the newer server has not been deployed yet.
  The server is asked every time you open this, so after an update the new version shows straight
  away. When the server has mails it could not send yet - the mail system on the server was down -
  a fourth line says how many. The server tries them again every few minutes for three days. A mail
  with a confirmation link or a code in it (signing up, a new password, a new address, an
  invitation) is not sent again: the person asks for a new one, or the administrator invites again.

Below those, the administrator also gets seven tools, each marked with a padlock. They are only for the administrator - nobody else sees them in the list, and the server turns everybody else away from them anyway:

- **Maintainers** - who maintains each board. The boards are listed in the same order as the **Boards** list. Choose a board to see its maintainers and the invitations nobody has used yet. Choose somebody who already has an account from the list and choose **Add as maintainer**, or type an email address and choose **Send invitation** to invite somebody new - an invitation code is mailed to that address. **Remove** takes a maintainer off the board, and **Withdraw** cancels an invitation. A board can have as many maintainers as you like. You can add yourself too, so other maintainers can see that you look after a board: you then get that board's emails about new submissions like its other maintainers - one email, not one as administrator and another as maintainer - and nothing else changes, as you can already review and publish every board.
- **Order of boards** - the order CRT shows the hardware and boards in its drop-down lists, with a **Hardware** and a **Board** column. Drag a board up or down, then choose **Save the order**. The order is written into the hardware and board list of both the BETA and the stable data, and CRT picks it up the next time it starts. **Cancel** drops the moves you have not saved.
- **Unused files** - files in the data that nothing uses any more, in the BETA or the stable data, shown as the same folder tree as a board's files, each with its size: point at a file to see it, and double-click it to open it. A file is listed only when no hardware and board list names it, no board's data file of any version - the older ones kept for older CRT versions included - cites it, and it is not in a "KiCad data" folder, the MiniPro IC tests or a file whose name starts with "!". If any data file cannot be read, nothing is listed at all. The BETA and the stable data are each checked on their own - pick one under **Data:** - so a file listed for BETA may still be used by the stable data, and the other way round. To delete the listed files from the server, tick "I have looked through this list, and every file on it can be removed from the server." and choose the remove button, which says how many files and how much space it removes. **Refresh** reads the list again.
- **Rebuild checksum manifests** - **Rebuild both manifests** rebuilds `dataChecksums.json` for both
  the BETA and the stable data. The server does this by itself whenever it publishes, promotes or rolls something back, so
  this is only needed when data files have been changed on the server by hand, which the service
  never sees - without a rebuild, those changes reach nobody. It reads every file again, so a file
  copied over by hand counts even when it kept the old file's size and date. It only reads the data
  and replaces one file, so it is safe to press again if something fails.
- **Delete a board** - removes a board from everywhere: its folder in the BETA and the stable
  data, its place in CRT's hardware and board lists, and everything the server knows about it - its
  submissions, maintainers, invitations and view statistics. Every board is listed - in the same order as the **Boards** list - with a
  **Delete** button. Pressing it first shows exactly what would go, and nothing is deleted until you
  confirm. If a submission to the board is still open - waiting for review, or in BETA - it is
  deleted too, and its contributor gets an email with the reason you give, so a reason is asked for
  then. Shared files the board used are kept; any that nothing else uses then show under **Unused
  files**. A deletion cannot be undone. Some boards cannot be deleted, and CRT says why instead of
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
