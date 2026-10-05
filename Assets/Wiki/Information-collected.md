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

Nothing you edit leaves your computer until you click **Submit** on the "Drafts" tab. A submission sends:

- The board's data and files, as your draft has them
- The description you write
- Your email address
  - So the maintainers can reply, and so you are told how it went. When you are signed in on the "Maintainer" tab, your account's address is used.
- The CRT version

Your email address is shown only to the maintainers of that system and to the administrator. Maintainers of other systems never see it - at most the name on your account, if you were signed in when you sent it.

While a submission is still open, CRT asks the server how it is getting on - at launch, and every minute while CRT's window is open and not minimised - using its number and the private code CRT was given when you sent it.

## When you discard a draft you have submitted

If you discard a draft on the "Drafts" tab while something you sent from it is still being reviewed - waiting, or published to the BETA source but not yet to everyone - CRT tells the server that you discarded your draft. The maintainers of that board then see it beside your contribution, so they can check with you before publishing it. It does not withdraw what you sent.

It sends only the number of that contribution and the private code CRT was given when you sent it (which proves the contribution is yours), nothing else. Without an internet connection it waits, and is sent the next time CRT starts and can reach the server.

## When you send feedback

The "Feedback" tab sends only what you give it:

- The email address you type, if any (or your account's, when you are signed in on the "Maintainer" tab)
- Your text
- The CRT version
- Attachments, only if you ask for them: CRT's log and crash log with "Attach application logfile and crash reports", its settings and traces files with "Attach configuration and traces data files", and any files or folder you add yourself

It arrives as an email to me. CRT's own log, crash log, settings and traces files are shown in that email; any other attached files are kept on the server for me to look at. Your IP address is not stored.

## Checking for a new version

With "Check for new version at application launch" ticked in the "Configuration" tab (it is by default), CRT asks GitHub at launch (and at once when you tick it) whether a newer release exists. That request goes to GitHub, not to my server. Untick it, and CRT does not ask.

## Downloading the data

With "Check for new or updated data at application launch" ticked in the "Configuration" tab (it is by default), CRT fetches the list of data files from `classic-repair-toolbox.dk` at launch and downloads what is new or changed. Those requests carry the CRT version and nothing else.

## Which CRT versions use which part of the server

Every time CRT talks to the server - the check-in, sending board views or feedback, submitting a contribution - it says which CRT version it is. The server counts, per day, how many times each CRT version used each part of the server. Nothing else is kept: no IP address, no account and no user or installation ID.

I use this to see when an old part of the server is no longer used by anybody, so it can be retired. A CRT version that still uses a retired part is told to update, instead of failing.
