[Wiki Home](Home)

What CRT sends home, and what it does not.

---

I want to be transparent here, and inform that I am gathering information about your setup, at every application launch, where the application does a mandatory "check-in":

- IP address
  - Ex. `85.184.162.75`
  - Used for pinning countries on a worldmap
- Operating system version
  - Ex. `Microsoft Windows 10.0.19045`
  - Used for knowing if my rewrite to natively support Linux and macOS was worth it
- CPU architecture used (32-bit or 64-bit)
  - Ex. `64-bit`
  - Used for knowing how wide usage that pesky self-contained .NET6 has (this is legacy and not used any more)

I am allowing myself to gather this data for me to build the [CRT Fun facts](https://classic-repair-toolbox.dk/funfacts/) page, which is some statistics on usage. As a developer, this is a personal motivational point to see countries using my application and of course one always hope for that "upwards trend usage"... which never happens 🤣 I find this limited non-personal data a fair amount to "pay" for using this application, taking in consideration for the effort being put in to this.

## Board views

CRT also counts which boards are being used. Every time a board has been on screen for at least 10 seconds, that counts as one view of that board - every time, so going to another board and back counts again. This is mandatory too, just like the check-in.

The views are collected on your computer and sent home in small batches, about a minute after they are counted. Without an internet connection they simply wait, and are sent the next time CRT can reach the server.

Each view sends:

- The board
  - Ex. `Commodore/C64/250407`
- When it was viewed
  - Ex. `2026-09-27 12:00:10` (UTC)
- The CRT version, operating system version and CPU architecture
  - The same as the check-in above
- Whether CRT is downloading data from the BETA source
  - That is mostly maintainers checking their own work, so those views are kept apart in the statistics

When the views arrive, the server looks up which country your IP address belongs to, and then forgets the IP address - it is not stored with the views. There is no user or installation ID either, so the views cannot be connected to you, or to each other.

I use this for the [CRT Fun facts](https://classic-repair-toolbox.dk/funfacts/) page, to see which boards are used and in which countries, and it shows the maintainers of each board how much their board is used.

## When you discard a draft you have submitted

If you discard a draft on the "Drafts" tab while something you sent from it is still being reviewed - waiting, or published to the BETA source but not yet to everyone - CRT tells the server that you discarded your draft. The maintainers of that board then see it beside your contribution, so they can check with you before publishing it. It does not withdraw what you sent.

It sends only the number of that contribution and the private code CRT was given when you sent it (which proves the contribution is yours), nothing else. Without an internet connection it waits, and is sent the next time CRT starts and can reach the server.
