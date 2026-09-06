[Wiki Home](Home)

Move the data or workbooks folder elsewhere, and fake an update.

---

CRT has three commandline parameters:

* [--data-root](#--data-root)
* [--workbooks-root](#--workbooks-root)
* [--simulate-update](#--simulate-update)

## --data-root

Puts the downloaded data (schematics, board data, images - close to 1 GB) somewhere else, for example on another drive.

```
--data-root=D:\CRT-data
--data-root="D:\My Folder With Spaces"
--data-root=/mydata/crt
```

Default if you do not use it:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Data`
* Linux and macOS: `~/.local/share/Classic-Repair-Toolbox/Data`

Good to know:

* Use a full path, not a relative one.
* Do not end the path with `\` or `/`.
* The folder is created if it does not exist. The first start then takes a while, as the data is copied into it.
* Your settings, log and workbooks stay where they are - this moves the downloaded data only.

## --workbooks-root

Puts your workbooks (repair jobs) somewhere else - for example on a synced drive, so you have them on more than one machine.

```
--workbooks-root=D:\Repairs
--workbooks-root="D:\My Repair Jobs"
```

Default if you do not use it:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Workbooks`
* Linux and macOS: `~/.local/share/Classic-Repair-Toolbox/Workbooks`

Same rules as `--data-root` above.

## --simulate-update

Shows the "a new version is available" banner without a new version existing, so you can see what it looks like.

```
--simulate-update
--simulate-update=2.5.1-beta.2
```

Without a version number it pretends version `99.0.0` is available.

Clicking "Install" runs the progress bar from 0% to 100% and stops there - nothing is downloaded and the application does not restart.

## Which folders am I actually using?

The "Configuration" tab has three buttons - `Open data folder`, `Open workbooks folder` and `Open logs and settings folder` - and each opens the folder CRT is really using. So if you have set one of the parameters below and want to check it took effect, the button is the quickest answer: it opens where the data actually is, not where it would have been by default.

The log file also has a `Data root is [...]` line near the top.
