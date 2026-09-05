# Classic Repair Toolbox

_Classic Repair Toolbox_ (or _CRT_ henceforward) is a cross-platform desktop application to assist hardware enthusiasts in diagnosing, troubleshooting, and repairing vintage computers and peripherals.

The project is a direct spin-off from an older project, _Commodore Repair Toolbox_ which now resides in a faint and distant memory only. The new _Classic_ project was realized as a complete rewrite, to be able to add natively support for **Linux** and **macOS**, but also to be able to support more hardware and not focus primarily on Commodore (Amstrad, ZX Spectrum and even different hardware types etc.).


## What can it do?

With CRT you can easily view technical schematics, zoom, identify components, view chip pinouts, use interactive (KiCad) traces or do manual circuit tracing, test ICs with a MiniPro programmer, study datasheets, view oscilloscope images, resources and various other information, helping you diagnosing and repairing old vintage hardware. You can also keep track on your repairs via the _Workbook_ feature.

It is (for now) having Commodore computers with the most documented systems, but it also has an Amstrad and a ZX Spectrum board (more systems will come for sure), but it can support any kind of hardware, as you can add your own data - e.g. other computers, radios, DIY electronics or whatever else you can imagine. It probably works the best, if the hardware is "simple" and have good documentation available, like schematics, and if it is something you need to revisit multiple times - then you can add the needed information yourself, and use it for easy future reference.


## Table of Contents

- [Requirements](#requirements)
- [Full documentation](#full-documentation)
- [Installation and usage](#installation-and-usage)
- [YouTube Quick-Help videos available](#youtube-quick-help-videos-available)
- [Built-in hardware and boards](#built-in-hardware-and-boards)
- [Data contributions being worked on currently](#data-contributions-being-worked-on-currently)
- [Supported oscilloscopes](#supported-oscilloscopes)
- [Help wanted](#help-wanted)
- [Contact developer](#contact-developer)
- [Screenshots](#screenshots)
- [License](#license)


## Requirements

- Operating systems supported:
  - **Windows 10** or newer (64-bit)
  - **macOS** on both **Apple Silicon** (arm64) and **Intel** (x64)
  - **Linux** (64-bit)
- Disk space needed: ~**2GB**

Note that the newest .NET LTS (Long-Term Support) is embedded in application, which means you do not need to have this installed. It also does mean that even if you have .NET installed on your computer, then it will still use the one embedded in application. Do note that it is the newest LTS at build time - it will _not_ get updated automatically and will stay as-is, until a newer version of CRT is released.


## Full documentation

Please go to the [Wiki](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki) to view a more comprehensive documentation, that is not suited for the project front page.


## Installation and usage

Download the newest normal (non-BETA) CRT version from [Releases](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/releases), and install it. The installation folder cannot be chosen by the user and is determined by the installation process. In the "Configuration" tab you can open the folder and see where the configuration and data files are stored.

If you want the actual data stored elsewhere than default, then this can be changed via a command-line parameter, view [Commandline parameters](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/wiki/Commandline-parameters).

Depending on your configuration settings, then CRT will check for newer data from its online source at application launch. It is recommended to have this enabled, as there will come many updates over time.

When a new version is released it will (per default, but can be configured) be shown to you in the application, and you can update directly from within the application.


## YouTube Quick-Help videos available

You can view the below _Quick Help_ videos for introduction to specific topics in CRT:

- [Short introduction](https://youtu.be/fwR018x39qg)
- [How to do manual traces](https://youtu.be/JUNXeCHsrME)
- [How to sync oscilloscope](https://youtu.be/CbTh1FFp3tU)
- [How to use component label editor](https://youtu.be/u-UkD-m4Z6o)
- [How to use interactive traces](https://youtu.be/Y55nC_gJbH4)


## Built-in hardware and boards

- **Amstrad CPC 664**
  - **MC0005A**
    - Oscilloscope baseline measurements
- **Commodore Plus/4**
  - **310163**
    - Oscilloscope baseline measurements for PAL and NTSC
- **Commodore VIC-20**
  - **324003**
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
  - **250403** (CR)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
- **Commodore 16**
  - **251789**
    - Oscilloscope baseline measurements for PAL and NTSC
- **Commodore 64**
  - **KU-14194HB**
    - Covers _all_ components
    - No oscilloscope baseline measurements - can you help?
    - Interactive (KiCad) traces and netlists
  - **250407** (long board)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
    - Interactive (KiCad) traces and netlists
  - **250425** (long board)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
  - **250466** (long board)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
  - **250469** (short board)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
    - Interactive (KiCad) traces and netlists
- **Commodore 128 and 128D** 
  - **310378** (C128 and C128D, plastic cabinet)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
    - Interactive (KiCad) traces and netlists
  - **250477** (C128DCR, metal cabinet)
    - Covers _all_ components
    - No oscilloscope baseline measurements yet - PAL is being worked on
    - Would appreciate help with:
      - Oscilloscope baseline for NTSC
- **ZX Spectrum 16K/48K**
  - **Issue 4B**
    - Interactive (KiCad) traces and netlists
    - Would appreciate help with:
      - Oscilloscope baseline for PAL and NTSC (not sure if they have differences for these regions?)


### Data contributions being worked on currently

Please let me know if you want to contribute with something, so it can be visualized here to avoid duplicate work.

**Commodore 128 and 128D** 
  - **250477** (C128DCR, metal cabinet)
    - As soon as time allows, then I will do a full PAL oscilloscope baseline


## Supported oscilloscopes

- **Keysight**
  - InfiniiVision 2000 X
  - InfiniiVision 6000 X
  - InfiniiVision 6000L
- **Rigol**
  - DHO800
  - DHO900
  - DHO1000
  - DHO4000
  - DS1000Z
  - DS2000A
  - MSO1000Z
  - MSO2000A
- **Rohde & Schwarz**
  - MXO 4
  - RTA4000
- **Siglent**
  - SDS800X HD
  - SDS1000CML+
  - SDS1000CNL+
  - SDS1000DL+
  - SDS1000X
  - SDS1000X+
  - SDS1000X-E
  - SDS1000X HD
  - SDS1202X-E
  - SDS2000X
  - SDS2000X HD
  - SDS2000X Plus
  - SDS3000X HD
  - SDS5000X
  - SDS6000 Pro
  - SDS6000A
  - SDS6000L
  - SDS7000A
  - SHS800X
  - SHS1000X


## Help wanted

I will keep adding and enhancing data, but if this is only me providing data, then it will take many years before this will reach a "premium level" - **if ever** 😁 So, I really do hope that the community will contribute, so it quickly can become a good source of information.

Data contribution can be almost anything - tiny and trivial updates (spelling mistakes, wrong or missing technical values or alike) or it can be huge new boards, but I really would like to get a massive amount of **quality** data, for the benefit of everyone using this. The goal is that it should have (most) relevant data in one place, so it would not be required to go and lookup for other data sources, but of course it also needs to be balanced a little, not overwelming with too much data 🤔

Contributing data is very easy - just go to the "Contribute" tab, select the component you want to edit and send your update - that's it.

You can help specifically with these topics:
- Do you have higher-quality images of the used schematics?
- Do you have (better) datasheets or pinouts for any of the components?
- Do you see missing components in either the component list or as a highlight?
- Can you improve any data or fill in more technical details anywhere?


## Contact developer

There are several ways to get in contact with the developer:

- Direct communication via [Retro Hardware Discord](https://discord.gg/kVTtdvZtzE) channel (accept invite on page)
- CRT "Feedback" tab
- GitHub [Issues](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/issues)


## Screenshots

Main schematics (with a component showing its direct connected traces):
<img width="900" height="539" alt="image" src="https://github.com/user-attachments/assets/c9e7a8e2-e17c-47da-aad7-dddfdb8a1775" />

Overview where a lot of component information is garthered:
<img width="900" height="539" alt="image" src="https://github.com/user-attachments/assets/4328b95c-b4ec-44e7-ae8c-4824b4023ae0" />

Resources relevant to the hardware and board:
<img width="900" height="539" alt="image" src="https://github.com/user-attachments/assets/784a279b-9522-4602-b6c5-cd3fc253a564" />

Oscilloscope configuration and test (it can auto-configure various settings per baseline image):
<img width="900" height="539" alt="image" src="https://github.com/user-attachments/assets/469866cd-94e0-4140-9eef-129eb4ab3aad" />

Configuration options:
<img width="900" height="539" alt="image" src="https://github.com/user-attachments/assets/168830f8-1f35-450a-bd0f-153d33493f41" />

One specific image from the oscilloscope baseline:
<img width="900" height="539" alt="image" src="https://github.com/user-attachments/assets/14632c12-2b88-4690-8cb8-fd20abb241ca" />

Some images can also have detailed explanation:
<img width="900" height="539" alt="image" src="https://github.com/user-attachments/assets/f0505e76-a7ce-4bb2-a854-1cccf06be372" />


## License

CRT is licensed under the **GNU General Public License v3** - see [LICENSE.txt](LICENSE.txt) for the
full license text.

Do note that this covers the application itself. The hardware reference data (schematics, datasheets,
component images and similar) is collected and contributed material, and the individual pieces of it
have their own origins and rights holders. For most important authors and contributors there is a link.
