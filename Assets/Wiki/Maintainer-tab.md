[Wiki Home](Home)

Review and publish the changes other people send in - for maintainers of the hardware data only.

---

The **Maintainer** tab is where a maintainer looks through the contributions sent in for the systems they look after, and publishes the good ones for everybody who uses CRT. It is hidden unless you turn it on, because it needs a maintainer account, and almost nobody using CRT has one - you can use every other part of CRT, and contribute data yourself, without it.

## Turning it on

In the [Configuration](Configuration-tab) tab, under **Maintainer**, tick **Enable Maintainer tab**. The tab then appears in the row of tabs, just before Configuration. Untick it again to hide it; that does not sign you out.

While the Maintainer tab is selected, the hardware and board list on the left and the worklog bar are put away, so its screens get the whole width of the window. They come back as soon as you choose another tab.

## Getting a maintainer account

You cannot create a maintainer account yourself. The administrator invites you by email, for the system or systems you will look after. The email holds an invitation code.

To accept it, open the Maintainer tab and choose **I have an invitation**. Paste the code, and pick the name other maintainers and contributors will see and a password. Your account is made for the address the invitation was sent to, so you then sign in with that address.

## Signing in

Sign in with your email address and password. If you have forgotten your password, choose **I forgot my password**: you are sent a code by email, which you paste into the tab together with the new password you want.

On Windows, CRT remembers that you are signed in, so you are not asked again the next time you open the tab. What it keeps is encrypted to your Windows user account and cannot be used from another computer or another user. On Linux and macOS there is no such safe place to keep it, so you sign in each time you open CRT.

**Sign out**, at the bottom of the list on the left, signs you out and forgets the remembered sign-in on this computer.

## The four screens

The buttons at the top left choose what the tab shows. Each is a list on the left and the item you choose in it on the right. Moving between them never closes anything, so a submission you were looking at is still there when you come back.

**Systems** - every system in the data, who maintains it, who has contributed to it and how their contributions went, and how often the board is looked at. A new system also gets its place in CRT's hardware and board lists here, which it must have before it can be approved.

**Contributor Submissions** - the queue of changes waiting for review, grouped by board. Choose one and its changes open as a table, with what changed coloured in. You can correct things in the table yourself before deciding. Then approve it (which publishes it to the BETA data, for checking), ask the contributor for changes, or reject it - the contributor sees your comment either way. The number on the button counts the systems waiting for you. The queue checks for new submissions by itself every minute while the tab is on screen.

**Beta > Prod** - systems whose approved changes are in the BETA data and waiting to go out to everybody. Once you have checked a system in BETA, publish it; or push it back to the queue if it is not right.

**Admin** - only for the administrator: files in the data that nothing uses any more.

If you close CRT with changes in a table that you have not saved, you are asked whether to save them first.
