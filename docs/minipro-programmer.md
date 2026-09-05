# MiniPro programmer

Go to [Documentation index](./README.md).

_CRT_ can test logic ICs and the C64 PLA with a MiniPro programmer (TL866/T48/T56-class), directly
from the component info popup. The IC must be out of the board and seated in the ZIF socket -
nothing is tested in-circuit.

> [!NOTE]
> I want to highlight, to be fair, that this feature is fully implemented and contributed by [@theretroloft](https://github.com/theretroloft), as I personally have no experience with how a MiniPro works or what it can be used for. I have merely changed some UI, but [@theretroloft](https://github.com/theretroloft) has done the huge heavy-lifting and implemented it in _CRT_.
> Also note that the communication is through the GitLab project, [minipro](https://gitlab.com/DavidGriffith/minipro).

## Enable it in CRT

* Go to the "Configuration" tab:
  * Tick `Enable MiniPro programmer functionality`
  * There is no need for you to enable the simulated demo mode, as this is used for _CRT_ development
    only

## What to install on your local system?

If you use Windows, then the MiniPro application is distributed with the _CRT_ application and can be
used out-of-the-box. If you use Linux or macOS, then you will need to install or compile MiniPro
yourself and make the executable available to CRT.

For Linux and macOS then it looks for the `minipro` executable in your `PATH`, and additionally in these locations:

* `/usr/local/bin`
* `/opt/homebrew/bin`
* `/opt/homebrew/sbin`
* `/usr/local/sbin`
* `~/.local/bin`

Note that adding a directory to `PATH` in `~/.profile` is **not** enough on its own: a desktop
application is not started from a login shell, so it never sees that file. If you build MiniPro
yourself, then copy or symlink the executable into one of the directories above - that is what CRT
actually looks at.

Source for MiniPro on Linux and macOS:
* https://gitlab.com/DavidGriffith/minipro/-/releases (0.7.4 as-of this writing)

### Windows

Nothing special needed here, as the Windows binary is included.

### Linux

* `sudo apt install build-essential`
* `sudo apt install pkg-config`
* `sudo apt install libusb-1.0-0-dev`
* `sudo apt install zlib1g-dev`
* Build MiniPro, and then copy the `minipro` executable to `~/.local/bin` (or another of the
  directories listed above)
* Check from anywhere in a terminal session that you can execute the `minipro` application

### macOS

* `brew install minipro`

This installs into a location CRT already looks in, so nothing else is needed. If you need to build
it manually, then you can try these steps, and let me know what to change if it does not work (as a
help for others):

* `brew install pkg-config`
* `brew install libusb`
* `brew link libusb`
* Build MiniPro, and then copy the `minipro` executable to `/usr/local/bin`
* Check from anywhere in a terminal session that you can execute the `minipro` application

## Which ICs can be tested?

The `Test IC with MiniPro programmer` button only appears for components that are categorized as an
`IC` and that have a test in CRT's test catalogue - currently 29 of the 74-series logic ICs plus the
C64 PLA (`906114-01`). The catalogue is part of the online-synchronized data, so new tests can arrive
without a new version of CRT.

## How to test a logic IC in CRT?

* Insert the IC in the ZIF socket, with pin 1 of the IC aligned with pin 1 of the socket
* Go to a logic component (e.g. `7406` or `74LS139` or alike) and open the component information
  popup
  * Click the `Test IC with MiniPro programmer` button
  * Click the `Run test` button - it should finish almost instantly

## How to test a C64 PLA IC in CRT?

* Insert the IC in the ZIF socket, with pin 1 of the IC aligned with pin 1 of the socket
* Go to the PLA component and open the component information popup
  * Click the `Test IC with MiniPro programmer` button
  * Choose either `Quick (25 vectors)` or `Standard (512 vectors)` in the `Test depth` drop-down
  * Click the `Run test` button - it should finish almost instantly

## Important note for test results

You can for sure trust the result if the test reports a FAIL - then the IC is for sure broken - but
if the test reports a SUCCESS, then you need to pay attention to this fact:

* The truth-table is tested OK - no issues in that, so logically the IC is working
* Timing is **not** tested, so in some cases it could _potentially_ be problematic or fail in the
  real physical C64, if it switches faster than what is being tested in CRT

The reason for this is that the MiniPro hardware (TL866/T48/T56) is simply not fast enough for this
timing test! As a best guess, the IC will most likely work without any issues in the C64 if it tested
OK in CRT.

Go to [Documentation index](./README.md).