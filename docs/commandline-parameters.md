# Commandline parameters

Go to [Documentation index](./README.md).

_CRT_ has two  commandline parameters:

* [--data-root](#--data-root)
* [--simulate-update](#--simulate-update)

## --data-root

Here you can specify which data folder you want to use. The data folder is where _CRT_ places all the data files it fetches from its
online source, and as this is a lot of data - close to 1 GB - then in some cases it can be useful to keep it somewhere else, for example on another drive.

If you do not use the parameter, then the data folder is:

* Windows: `%LocalAppData%\Classic-Repair-Toolbox\Data`
* Linux and macOS: `~/.local/share/Classic-Repair-Toolbox/Data`

It is deliberately placed outside the installation folder, because the installation folder is
replaced when _CRT_ updates itself - data kept there would not survive an update.

If the path does not exist, then _CRT_ will try and create it. Note that the first start with a new
data folder takes a while, as the data is copied into it from the installation folder.

Parameter examples:
- `--data-root=/mydata/crt`
- `--data-root="D:\My Folder With Spaces"`

A few things worth knowing:

* Use a full path. A relative path is resolved from wherever _CRT_ happens to be started from,
  which is rarely what you want.
* Do not end the path with a backslash or a slash - on Windows a trailing `\` in front of the
  closing quote is read as an escaped quote, and the path gets mangled.
* The parameter itself is not case-sensitive, and anything else you put on the commandline is
  ignored.
* If you are in doubt about which folder is actually being used, then look in the log file for the
  `Data root is [...]` line. The log and your settings always stay in
  `%LocalAppData%\Classic-Repair-Toolbox` (`~/.local/share/Classic-Repair-Toolbox`) no matter what
  you set here, and the `Open data/log/settings folder` button on the "Configuration" tab opens
  exactly that folder.

## --simulate-update

Here you can simulate as like a newer version is available online - to view how the UI for that would look like. You give it a version number like this:

- `----simulate-update=2.5.0`
- `----simulate-update=2.5.1-beta.2`

The UI will then show how this looks like, and clicking the "Install" button will simulate that, but without doing anything else, rather than go from 0% to 100% (no application restart or alike).

Go to [Documentation index](./README.md).