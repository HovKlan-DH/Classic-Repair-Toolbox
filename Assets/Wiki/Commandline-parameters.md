[Wiki Home](Home)

Move the data, workbooks or drafts folder elsewhere, and fake an update.

---

CRT has four commandline parameters:

* [--data-root](#--data-root)
* [--workbooks-root](#--workbooks-root)
* [--drafts-root](#--drafts-root)
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
* Linux: `~/.local/share/Classic-Repair-Toolbox/Data`
* macOS: `~/Library/Application Support/Classic-Repair-Toolbox/Data`

Good to know:

* Use a full path, not a relative one.
* Do not end the path with `\` or `/`.
* The folder is created if it does not exist. The first start then takes a while, as the data is copied (or downloaded) into it.
* Your settings, log, workbooks and drafts stay where they are - this moves the downloaded data only.

## --workbooks-root

Puts your workbooks (your repairs) somewhere else - for example on a synced drive, so you have them on more than one machine.

```
--workbooks-root=D:\Repairs
--workbooks-root="D:\My Repairs"
```

Default if you do not use it:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Workbooks`
* Linux: `~/.local/share/Classic-Repair-Toolbox/Workbooks`
* macOS: `~/Library/Application Support/Classic-Repair-Toolbox/Workbooks`

Same rules as `--data-root` above.

## --drafts-root

Puts your local, unpublished edits to hardware and board data somewhere else - the same idea as `--workbooks-root` above, but for drafted contributions rather than repairs.

```
--drafts-root=D:\CRT-drafts
--drafts-root="D:\My Draft Edits"
```

Default if you do not use it:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Drafts`
* Linux: `~/.local/share/Classic-Repair-Toolbox/Drafts`
* macOS: `~/Library/Application Support/Classic-Repair-Toolbox/Drafts`

Same rules as `--data-root` above.

## --simulate-update

Shows the "a new version is available" banner without a new version existing, so you can see what it looks like. The banner says `(simulated)` after the version.

```
--simulate-update
--simulate-update=3.0.1-beta.1
```

Without a version number it pretends version `99.0.0` is available.

Clicking "Install" runs the progress bar from 0% to 100% and stops there - nothing is downloaded and the application does not restart.

## Which folders am I actually using?

The "Configuration" tab has three buttons - `Open data folder`, `Open workbooks folder` and `Open logs and settings folder` - and each opens the folder CRT is really using. So if you have set one of the parameters above and want to check it took effect, the button is the quickest answer: it opens where the data actually is, not where it would have been by default.

There is no button for the drafts folder. The log file shows all three near the top: a `Data root is [...]` line, a `Drafts root is [...]` line, and a `Worklog loaded: [...] from [...]` line for the workbooks folder.
