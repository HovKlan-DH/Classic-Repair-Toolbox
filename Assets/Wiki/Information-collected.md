[Wiki Home](Home)

What CRT sends home, and what it does not.

---

I want to be transparent here, and inform that I am gathering information about your setup, at every application launch, where the application does a mandatory "check-in" once its window has opened. It sends:

- CRT version
  - Ex. `CRT 3.0.0`
  - Used for knowing which versions are still in use
- Operating system and its version
  - Ex. `Windows` and `Microsoft Windows 10.0.19045`
  - Used for knowing if my rewrite to natively support Linux and macOS was worth it
- CPU architecture
  - Ex. `64-bit` or `ARM 64-bit`
  - Used for knowing which builds are worth making

The server stores this together with the date and your IP address (ex. `85.184.162.75`). The IP address is used for counting how many different installations there are, and for pinning countries on a worldmap: the server asks [ip-api.com](https://ip-api.com/), an outside service, which country the address belongs to, and stores the country with it.

I am allowing myself to gather this data for me to build the [CRT Fun facts](https://classic-repair-toolbox.dk/funfacts/) page, which is some statistics on usage. As a developer, this is a personal motivational point to see countries using my application and of course one always hope for that "upwards trend usage"... which never happens 🤣 I find this limited data a fair amount to "pay" for using this application, taking in consideration for the effort being put in to this.

## Board views

CRT also counts which boards are being used. Every time a published board has been on screen for at least 10 seconds, that counts as one view of that board - every time, so going to another board and back counts again. A board that exists only as your own draft is never counted. This is mandatory too, just like the check-in.

The views are collected on your computer and sent home in small batches, about a minute after they are counted. Without an internet connection they simply wait, and are sent the next time CRT can reach the server (views older than 30 days are dropped).

Each view sends:

- The board
  - Ex. `Commodore/C64/250407`
- When it was viewed
  - Ex. `2026-09-27 12:00:10` (UTC)
- The CRT version, operating system and CPU architecture
  - The same as the check-in above
- Whether CRT is downloading data from the BETA source
  - That is mostly maintainers checking their own work, so those views are kept apart in the statistics

When the views arrive, the server asks ip-api.com which country your IP address belongs to, and then forgets the IP address - it is not stored with the views. There is no user or installation ID either, so the views cannot be connected to you. Each batch carries a random number of its own, kept apart from the views for 60 days, only so that a batch sent twice is counted once.

I use this for the [CRT Fun facts](https://classic-repair-toolbox.dk/funfacts/) page, to see which boards are used and in which countries, and it shows the maintainers of each board how much their board is used.

## When you submit a contribution

Nothing you edit leaves your computer until you click **Submit** on the "Contribute" tab. A submission sends:

- The board's data and files, as your draft has them - your credit included, unless you chose to send none (below)
- Your email address
  - So the maintainers can reply, and so you are told how it went. When you are signed in to your account, its address is used.
- The CRT version

Your email address is shown only to the maintainers of that board and to the administrator. Maintainers of other boards never see it - at most the name on your account, if you were signed in when you sent it.

The server keeps the IP address a submission came from, only to limit how many submissions one address can send in a day.

Your first change to a board adds a credit for your work to its "Board: Credits" on the "Contribute" tab, framed in green: what kind of work it was, the name and contact you gave last time, and a main contribution for you to describe - and the board's credits are shown to everybody who uses CRT. The credit is a suggestion: the board's maintainers may change it before it is published. Write "Unknown" as the name to keep the credit without your name, or press "No credit for my changes" on it to send your changes with no credit at all - "Add a credit for my changes" brings it back. The main contribution is also what the maintainers read as the description of your submission; with no credit, they see the list of what changed instead.

## When you send in an oscilloscope model

**Submit**, under "Submit this model for everybody" on the "Oscilloscope" tab, sends:

- The model: its vendor, its name and every command and setting of it
- Your email address (your account's, when you are signed in), so you are told whether it was published
- Your name, if you give one, and your message
- Whether CRT's full test suite went through on the model
- The CRT version

Only the administrator sees it. The server keeps the IP address it came from - of an IPv6 address only its network part - for two things only: to limit how many models one internet connection can send in a day, and, when you are not signed in, to know that a newer copy of your model sent from the same connection and address replaces the one still waiting. A published model is part of CRT's data for everybody - without your name or address.

While a submission is still open, CRT asks the server how it is getting on - at launch, and every minute while CRT's window is open and not minimised - using its number and the private code CRT was given when you sent it.

## When you make an account

An account is not needed for anything above. You can make one with the account icon in the top right corner of CRT's window, or on the website's [account page](https://classic-repair-toolbox.dk/account/) - the same account in both. The server keeps:

- Your email address
  - To sign in with, and for the mails about your account. It is never shown to other people, only to the administrator - and to the maintainers of a board you send something to, as above.
- Your display name
  - Shown to the maintainers and the administrator beside what you send.
- Your password - only as a scrambled check value, never the password itself
- Each sign-in: when it was made, the IP address it came from and the name your CRT gives itself (ex. `CRT 4.1.0`)
  - So the sign-ins of one account can be told apart. A sign-in ends when you sign out, or by itself when it has not been used for 30 days.
- A record of what happened to the account - made, confirmed, signed in and out, a password that did not match, a changed address, name or password - with the IP address of each sign-in and of each password that did not match
  - Kept as the server's history, which only the administrator can read, to find out what happened if an account is ever misused.

While you are signed in, CRT uses your account's email address in the "Feedback" tab, the submit dialog and when you send an oscilloscope model.

To limit how often it can happen, the server also notes the IP address of each request that sends a mail (making an account, asking for a code or a new link) and of each failed sign-in. Those notes are deleted after a day.

## When you discard a draft you have submitted

If you discard a draft (**Discard draft** on the "Contribute" tab) while something you sent from it is still being reviewed - waiting, or published to the BETA source but not yet to everyone - CRT tells the server that you discarded your draft. The maintainers of that board then see it beside your contribution, so they can check with you before publishing it. It does not withdraw what you sent.

It sends only the number of that contribution and the private code CRT was given when you sent it (which proves the contribution is yours), nothing else. Without an internet connection it waits, and is sent the next time CRT starts and can reach the server.

## When you send feedback

The "Feedback" tab sends only what you give it:

- The email address you type, if any (or your account's, when you are signed in)
- Your text
- The CRT version
- Attachments, only if you ask for them: CRT's log and crash log with "Attach application logfile and crash reports", its settings and traces files with "Attach configuration and traces data files", and any files or folder you add yourself

It arrives as an email to me. CRT's own log, crash log, settings and traces files are shown in that email; any other attached files are kept on the server for me to look at. Your IP address is not stored.

## Checking for a new version

With "Check for new version at application launch" checked in the "Configuration" tab (it is by default), CRT asks GitHub at launch (and at once when you check it) whether a newer release exists. That request goes to GitHub, not to my server. Uncheck it, and CRT does not ask.

## Downloading the data

With "Check for new or updated data at application launch" checked in the "Configuration" tab (it is by default), CRT fetches the list of data files from `classic-repair-toolbox.dk` at launch and downloads what is new or changed. Those requests carry the CRT version and nothing else.

## Which CRT versions use which part of the server

Every time CRT talks to the server - the check-in, sending board views or feedback, submitting a contribution - it says which CRT version it is. The server counts, per day, how many times each CRT version used each part of the server. Nothing else is kept: no IP address, no account and no user or installation ID.

While you have drafts, or the "Maintainer" tab is turned on, CRT also asks the server once when it starts (or when either first comes into use) whether this CRT is still new enough for it, so those tabs can say so if CRT has to be updated. That request carries the CRT version and nothing else.

I use this to see when an old part of the server is no longer used by anybody, so it can be retired. A CRT version that still uses a retired part is told to update, instead of failing.
