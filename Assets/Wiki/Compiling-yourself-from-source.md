[Wiki Home](Home)

Build CRT on Windows, Linux or macOS with the plain `dotnet` CLI.

---

CRT is a .NET 10 / Avalonia desktop application and builds with the plain `dotnet` CLI on all
three platforms. The build is identical everywhere - only **how you install the .NET 10 SDK** differs.

## Contents

- [Before you start](#before-you-start)
- [Quick start (any platform)](#quick-start-to-build-the-executable-any-platform)
- [Debug vs Release](#debug-vs-release)
- [Windows](#windows)
- [Linux](#linux)
- [macOS](#macos)
- [Self-contained builds](#self-contained-builds)
- [Running the test suite](#running-the-test-suite)
- [Where the hardware data comes from](#where-the-hardware-data-comes-from)
- [The MiniPro IC programmer](#the-minipro-ic-programmer)
- [Troubleshooting](#troubleshooting)

## Before you start

You need the **.NET 10 SDK** ([download](https://dotnet.microsoft.com/en-us/download/dotnet/10.0) or
your package manager), **Git**, and about **2 GB free disk space**. A newer SDK is fine - there is no
`global.json` pinning one. The project also sets `RollForward=LatestMajor`, so the built application
runs on a newer .NET runtime too.

Fork the [CRT GitHub repository](https://github.com/HovKlan-DH/Classic-Repair-Toolbox) first to your own GitHub repository and then from your development environment:

```
git clone https://github.com/<your-username>/Classic-Repair-Toolbox.git
cd Classic-Repair-Toolbox
dotnet --list-sdks
```

The last command must list a `10.x.x` entry.

## Quick start to build the executable (any platform)

```
dotnet restore Classic-Repair-Toolbox.slnx
dotnet build   Classic-Repair-Toolbox.slnx -c Release
dotnet test    Classic-Repair-Toolbox.slnx -c Release
```

The application executable lands in `src/CRT.App/bin/Release/net10.0/`; `Classic-Repair-Toolbox.exe` on Windows and
`Classic-Repair-Toolbox` on Linux and macOS.\
To run from the source tree:

```
dotnet run --project src/CRT.App/CRT.App.csproj -c Release
```

The solution holds more than CRT itself: `src/CRT.App` is CRT, `src/CRT.Data` is the board-data
library CRT and the server share, `src/CRT.Server` is the contribution service (submissions,
feedback, the launch check-in), and `tests/` holds a test project for each. You do not need the
server to build, test or run CRT - CRT talks to the one hosted at `classic-repair-toolbox.dk`.

> `build` and `test` target the **solution** (`.slnx`) so the tests come along. `run` and `publish`
> target the **project** (`.csproj`), or you would publish the server and the test projects too.

## Debug vs Release

**Both configurations behave identically** - same update check, same data sync, same diagnostics.
Build `Release` for anything you intend to use or measure; `Debug` is unoptimised and slower. Only a
`publish` with `-r` is also precompiled - see [Self-contained builds](#self-contained-builds).

## Windows

**Visual Studio 2026** - open `Classic-Repair-Toolbox.slnx`, set the configuration to **`Release`**, and then `Build` > `Build Solution`.

**Visual Studio Code** - with the **C# Dev Kit** extension you get the `build` task, the `watch` task
(rebuild-and-restart on save), and **F5** to debug. Both `watch` and F5 pass `--simulate-update`.

**Command line**

```
dotnet build Classic-Repair-Toolbox.slnx -c Release
src\CRT.App\bin\Release\net10.0\Classic-Repair-Toolbox.exe
```

## Linux

Install the .NET 10 SDK - this is the only distro-specific part:

| Distro | Command |
| --- | --- |
| Fedora | `sudo dnf install dotnet-sdk-10.0` |
| Debian / Ubuntu | `sudo apt install dotnet-sdk-10.0` |
| Arch | `sudo pacman -S dotnet-sdk` |
| Any distro | `curl -sSL https://dot.net/v1/dotnet-install.sh \| bash -s -- --channel 10.0` |

The install script needs no root and puts the SDK in `~/.dotnet` - add
`export PATH="$HOME/.dotnet:$PATH"` to your shell profile.

**Gentoo** keeps several SDKs side by side, so select one:

```
eselect dotnet list          # show available versions
eselect dotnet set 1         # pick the .NET 10 profile
. /etc/profile               # reload environment
```

Verify with `dotnet --list-sdks` that you got `10.x`, then:

```
dotnet build Classic-Repair-Toolbox.slnx -c Release
./src/CRT.App/bin/Release/net10.0/Classic-Repair-Toolbox
```

## macOS

Install the SDK with the `.pkg` from
[dotnet.microsoft.com](https://dotnet.microsoft.com/en-us/download/dotnet/10.0) or
`brew install --cask dotnet-sdk`, then:

```
dotnet build Classic-Repair-Toolbox.slnx -c Release
./src/CRT.App/bin/Release/net10.0/Classic-Repair-Toolbox
```

Two quirks:

- You get a **bare executable, not a `.app` bundle** - the `.app` in official releases is made by
  Velopack during packaging. Running from Terminal works fine.
- Local builds are **not signed or notarised**. If macOS refuses to start it:
  `xattr -dr com.apple.quarantine ./src/CRT.App/bin/Release/net10.0/Classic-Repair-Toolbox`

Use `osx-arm64` for Apple Silicon and `osx-x64` for Intel when publishing.

## Self-contained builds

Bundles the .NET runtime so it runs without .NET installed. Note this is `publish` on the **project**
file, and `--self-contained` requires `-r`:

```
dotnet publish src/CRT.App/CRT.App.csproj -c Release -f net10.0 -r <rid> --self-contained
```

| Target | `<rid>` | Output |
| --- | --- | --- |
| Windows x64 | `win-x64` | `src/CRT.App/bin/Release/net10.0/win-x64/publish/` |
| Linux x64 | `linux-x64` | `src/CRT.App/bin/Release/net10.0/linux-x64/publish/` |
| macOS Apple Silicon | `osx-arm64` | `src/CRT.App/bin/Release/net10.0/osx-arm64/publish/` |
| macOS Intel | `osx-x64` | `src/CRT.App/bin/Release/net10.0/osx-x64/publish/` |

Add `-o <folder>` to choose the output folder. Passing `-r` also enables **ReadyToRun** - bigger and
slower to build, noticeably quicker to launch. Cross-compiling works; only the macOS packaging step
in CI needs a Mac.

## Running the test suite

There are three xUnit test projects: `tests/CRT.App.Tests` (CRT itself, including headless UI tests
of its tabs and windows), `tests/CRT.Data.Tests` (the shared data library) and
`tests/CRT.Server.Tests` (the contribution service). No hardware, no network, no display and no
database - the whole run takes about two minutes in Release, build included:

```
dotnet test Classic-Repair-Toolbox.slnx -c Release
```

Every push runs the same suite on GitHub and a red suite blocks releases, so run it before sending a
pull request. If you add or change logic, add or update the tests in the same change.

## Where the hardware data comes from

The ~1 GB in `Assets/Data` is **not** copied into the build output - official installers bundle it, but a
source build does not. On first launch CRT creates its data folder and downloads from its stable source,
`classic-repair-toolbox.dk` - as long as "Check for new or updated data at application launch" is ticked in
the "Configuration" tab, which it is by default.

| Platform | Data folder |
| --- | --- |
| Windows | `%LOCALAPPDATA%\Classic-Repair-Toolbox\Data` |
| Linux | `~/.local/share/Classic-Repair-Toolbox/Data` |
| macOS | `~/Library/Application Support/Classic-Repair-Toolbox/Data` |

## The MiniPro IC programmer

Only the `win-x64` build bundles a `minipro` binary (committed at `Assets/MiniPro/win-x64/`, copied
next to the executable automatically).

On Linux and macOS, install `minipro` yourself, but please view [MiniPro programmer](MiniPro-programmer).

## Troubleshooting

* **`dotnet` cannot find a .NET 10 SDK**
  * Run `dotnet --list-sdks`. No `10.x` entry means it is not installed or not on `PATH`. On Gentoo,
check you ran `eselect dotnet set` and re-sourced `/etc/profile`.
* **`--simulate-update` shows no banner**
  * Tick "Check for new version at application launch" in the "Configuration" tab. Without it no update
check runs, so there is nothing for the simulation to answer. The log states this at startup.
* **Builds on Linux but exits immediately**
  * Avalonia needs fonts and ICU. On a minimal or container install add `fontconfig` and your distro's
`libicu` package. Check the log for the real error.
* **Application launches but there is no hardware data**
  * Expected on a first run from source - it is downloading. If nothing arrives, check that "Check for
new or updated data at application launch" is ticked. See
[Where the hardware data comes from](#where-the-hardware-data-comes-from).
