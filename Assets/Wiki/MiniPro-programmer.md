[Wiki Home](Home)

Test a logic IC or a C64 PLA out of the board.

---

CRT can test logic ICs and the C64 PLA with a MiniPro programmer (TL866/T48/T56-class), directly
from the component info popup. The IC must be out of the board and seated in the ZIF socket -
nothing is tested in-circuit.

> [!NOTE]
> I want to highlight, to be fair, that this feature is fully implemented and contributed by [@theretroloft](https://github.com/theretroloft), as I personally have no experience with how a MiniPro works or what it can be used for. I have merely changed some UI, but [@theretroloft](https://github.com/theretroloft) has done the huge heavy-lifting and implemented it in CRT.
> Also note that the communication is through the GitLab project, [minipro](https://gitlab.com/DavidGriffith/minipro).

## Enable it in CRT

* Go to the "Configuration" tab:
  * Make sure `Enable MiniPro programmer functionality` is ticked - it is on by default
  * There is no need for you to enable `Enable MiniPro programmer simulated demo mode (only
    required for CRT development)`, as this is used for CRT development only

## What to install on your local system?

If you use Windows, then the MiniPro application is distributed with the CRT application and can be
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
`IC` and that have a test in CRT's test catalogue. A component is matched by its "Technical name or
value" (e.g. `74LS139`; `7406` matches the `74LS06` test), not by its part number. A few ICs from
other logic families have the same pins and truth table as a 74LS IC, and use its test: `7402`,
`7406`, `7407`, `74F32` and `74HCU04`.

The catalogue currently holds 36 of the 74-series logic ICs, the `4066` bilateral switch, the `556`
timer and the C64 PLA (`906114-01`), and every one of them can be run. Among them are ICs that
remember something - the flip-flop `74LS74`, the latch `74LS373` and the counter `74LS193` - which
are tested as a sequence of steps rather than as a truth table. The catalogue is part of the
online-synchronized data, so new tests can arrive without a new version of CRT.

## How to test a logic IC in CRT?

* Insert the IC in the ZIF socket, with pin 1 of the IC aligned with pin 1 of the socket
* Go to a logic component (e.g. `7406` or `74LS139` or alike) and open the component information
  popup
  * Click the `Test IC with MiniPro programmer` button
  * The `Test depth` shows how many test vectors are run; there is nothing to choose
  * Click the `Run test` button - most ICs finish almost instantly; the `74LS153` and `74LS133`
    run thousands of test vectors and take a little longer, with the elapsed time shown beside the
    buttons

## How to test a C64 PLA IC in CRT?

* Insert the IC in the ZIF socket, with pin 1 of the IC aligned with pin 1 of the socket
* Go to the PLA component and open the component information popup
  * Click the `Test IC with MiniPro programmer` button
  * Choose `Quick (25)`, `Standard (512)` or `Full (65536)` in the `Test depth` drop-down - the
    number is how many test vectors are run. Quick and Standard finish almost instantly;
    Full tests every combination and takes longer, with the elapsed time shown beside the buttons.
  * Click the `Run test` button

While a test runs, `Cancel` stops it. Below the result, `Raw MiniPro output:` shows what the
programmer reported, which is useful when a test fails. `Close MiniPro Test` returns to the popup.

## Important note for test results

The result reads PASS or FAIL. You can for sure trust the result if the test reports a FAIL - then
the IC is for sure broken, and CRT names the failing pins when it can. Before you throw an IC away,
check that it sits fully in the socket with the lever closed, and run the test again - a badly
seated IC fails too. If the test reports a PASS, then you need to pay attention to this fact:

* The truth-table is tested OK - no issues in that, so logically the IC is working
* Timing is **not** tested, so in some cases it could _potentially_ be problematic or fail in the
  real physical C64, if it switches faster than what is being tested in CRT

The reason for this is that the MiniPro hardware (TL866/T48/T56) is simply not fast enough for this
timing test! As a best guess, the IC will most likely work without any issues in the C64 if it tested
OK in CRT.

For the `4066` bilateral switch, a PASS means that every switch opens and closes when it should, and
passes both a low and a high level in both directions. It does not measure how well a closed switch
passes an analog signal, nor a small leak in an open one - and that can still matter where the switch
carries an analog signal, such as the C64's paddle inputs.

For the `74LS74`, `74LS373` and `74LS193`, the test is a sequence of steps that goes through every
function in the datasheet - for example clear, preset and clocking for the flip-flop, holding a 0 and
a 1 for the latch, and counting up and down through all 16 states for the counter. A PASS means every
one of those functions works; no test can try every possible order of events.

For the `556` timer, a PASS means that each of its two timers sets, resets, holds, and drives its
output and discharge pins. The test does not measure timing, the exact trigger and threshold levels,
or how much current the output can deliver.
