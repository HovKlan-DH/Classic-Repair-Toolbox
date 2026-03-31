# Classic Repair Toolbox

_Classic Repair Toolbox_ (or **CRT** hence forward) is a cross-platform desktop application to assist hardware enthusiasts in diagnosing, troubleshooting, and repairing vintage computers and peripherals.

The project is a direct spin-off from an older project, **Commodore Repair Toolbox** which now resides in a faint and distant memory only. The new _Classic_ project (compared to _Commodore_) was realized as a complete rewrite, to be able to add natively support for **Linux** and **macOS**, but also to be able to support more hardware and not focus primarily on Commodore (Amstrad, Spectrum?).


## What is it?

With _CRT_ you can easily view technical schematics, zoom, identify components, view chip pinouts, do manual circuit tracing, study datasheets, view oscilloscope images, ressources and various other information, helping you diagnosing and repairing old vintage hardware.

It is (for now) primarily dedicated to Commodore, and have several built-in profiles for Commodore computers and it has a single Amstrad computer also, but it can support any kind of hardware, as you can add your own data - e.g. other computers, radios, DIY electronics or whatever else you can imagine. It probably works the best, if the hardware is "simple" and have good documentation available, like schematics, and if it is something you need to revisit multiple times - then you can add the needed information yourself, and use it for easy future reference.


## Table of Contents

- [Installation and usage](#installation-and-usage)
- [Built-in hardware and boards](#built-in-hardware-and-boards)
- [Data contributions being worked on currently](#data-contributions-being-worked-on-currently)
- [Supported oscilloscopes](#supported-oscilloscopes)
- [Requirements](#requirements)
- [Help wanted](#help-wanted)
- [Contact developer](#contact-developer)
- [Technical topics](#technical-topics)
  - [YouTube Quick-Help videos available](#youtube-quick-help-videos-available)
  - [Information automatically collected by CRT](#information-automatically-collected-by-crt)
  - [Commandline parameters](#commandline-parameters)
  - [How to contribute with data to CRT GitHub repository?](#how-to-contribute-with-data-to-crt-github-repository)
  - [Compiling yourself](#compiling-yourself)
  - [Controlling oscilloscope with keyboard/numpad](#controlling-oscilloscope-with-keyboardnumpad)
  - [Development tools used](#development-tools-used)
- [Inspiration for building this application](#inspiration-for-building-this-application)
- [Screenshots](#screenshots)


## Installation and usage

Download the newest normal (non-NETA) _CRT_ version from [Releases](https://github.com/HovKlan-DH/Commodore-Repair-Toolbox/releases), and install it afterwards. The installation folder cannot be set by the user, and is decided by installation system (Avalonia), but in the `Configuration` tab you can open the folder and view where it has its configuration and data files located.

If needed then the `data-root` folder can be changed via a commandline parameter, view [Commandline parameters](#commandline-parameters).

Depending on your configuration settings, then _CRT_ will check for newer data at application launch. It is recommended to have this enabled, as there will come many updates over time.

When a new version is released it will be shown to you in the application, and you can update directly from within the application.


## Built-in hardware and boards

- **Amstrad CPC 664**
  - MC0005A
- **Commodore VIC-20**
  - 250403 (CR)
    - Oscilloscope baseline measurements for PAL and NTSC
- **Commodore 64**
  - 250407 (long board)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
  - 250425 (long board)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
  - 250466 (long board)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
  - 250469 (short board)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
- **Commodore 128 and 128D** 
  - 310378 (C128 and C128D, plastic cabinet)
    - Covers _all_ components
    - Oscilloscope baseline measurements for PAL and NTSC
  - 250477 (C128DCR, metal cabinet)
    - Covers _all_ components
    - Would appreciate help with:
      - Oscilloscope baseline for NTSC


### Data contributions being worked on currently

- **@Rabs** is doing an oscilloscope baseline for **Amstrad CPC 664**
- Please let me know if you want to contribute with something, so it can be visualized here to avoid duplicate work.


## Supported oscilloscopes

- **Keysight**
  - InfiniiVision 2000 X
  - InfiniiVision 6000 X
  - InfiniiVision 6000L
- **Rigol**
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


## Requirements

- Operating systems supported:
  - **Windows 7** or newer (32-bit and 64-bit)
  - **macOS** (64-bit)
  - **Linux** (64-bit)

Note that .NET is embedded in application, which means you do not need to have this installed. It also does mean that even if you have .NET installed on your computer, then it will still use the one embedded in application. As .NET6 is the newest version supported on **Windows 7**, then this is the .NET version included with the Windows 32-bit installer.

> [!CAUTION]
> .NET6 used for Windows 32-bit has gone **End-of-Life in 2024**, and has not received any security hotfixes since then!

If possible then you should use the newer _CRT_ 64-bit installer, which embeds the newest available [.NET10 LTS](https://github.com/dotnet/core/blob/main/release-notes/10.0/README.md) (Long-Term-Support) at release date.


## Help wanted

I will keep adding and enhancing data, but if this is only me providing data, then it will take many years before this will reach a "premium level" - **if ever** 😁 So, I really do hope that the community will contribute, so it quickly can become a good source of information.

Data contribution can be almost anything - tiny and trivial updates (spelling mistakes, wrong or missing technical values or alike) or it can be huge new boards, but I really would like to get a massive amount of **quality** data, for the benefit of everyone using this. The goal is that it should have (most) relevant data in one place, so it would not be required to go and lookup for other data sources, but of course it also needs to be balanced a little, not overwelming with too much data 🤔

You can help specifically with these topics:
- Do you have higher-quality images of the used schematics?
- Do you have (better) datasheets or pinouts for any of the components?
- Do you see missing components in either the component list or as a highlight?
- Can you improve any data or fill in more technical details anywhere?


## Contact developer

There are several ways to get in contact with the developer:

- Direct communication via [Retro Hardware Discord](https://discord.gg/HDWct2vxem) channel (accept invite on page)
- CRT "Feedback" tab
- GitHub [Issues](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/issues)


## Technical topics

### YouTube Quick-Help videos available

You can view the below _Quick Help_ videos for introduction to specific topics in _CRT_:

- [Short introduction](https://youtu.be/FamNSLKJpdI)
- [How to do manual traces](https://youtu.be/JUNXeCHsrME)
- [How to sync oscilloscope](https://youtu.be/CbTh1FFp3tU)

### Information automatically collected by CRT

I want to be transparent here, and inform that I am gathering information about your setup, at every application launch, where the application does a mandatory "check-in":

- IP address
  - Ex. `85.184.162.75`
  - Used for pinning countries on a worldmap
- Operating system version
  - Ex. `Microsoft Windows 10.0.19045`
  - Used for knowing if my rewrite to natively support Linux and macOS was worth it
- CPU architechture used (32-bit or 64-bit)
  - Ex. `64-bit`
  - Used for knowing how wide usage that pesky self-contained .NET6 has

I am allowing myself to gather this data for me to build the [CRT Fun facts](https://classic-repair-toolbox.dk/funfacts/) page, which is some statistics on usage. As a developer, this is a personal motivational point to see countries using my application and of course one always hope for that "upwards trend usage"... which never happens 🤣 I find this limited non-personal data a fair amount to "pay" for using this application, taking in consideration for the effort being put in to this.


### Commandline parameters

_CRT_ supports currently only a single commandline parameter, where you can specify which data folder you want to use. The data folder is where it place all its data files that can be fetched from its online source, and as this can be a lot of data (+500MB), then maybe in some cases it could be useful to save this somewhere else.

If the path does not exists, it will try and create it.

Parameter examples:
- `--data-root=/mydata/crt`
- `--data-root="D:\My Folder With Spaces\"`


### How to contribute with data to CRT GitHub repository?

One possibility to contribute data is by submitting it directly to the GitHub repository, and in this way you will also be seen as a contributor. There are are some basic steps that you can follow, if you want to contribute data to CRT. It is quite easy, but it does require you have a GitHub account.

- Fork the _CRT_ GitHub repository
- Clone the fork to your local computer
- **Create a new branch** (important!)
- Do your own modifications:
  - Change existing files
  - Add new files
- Commit changes to your forked repo and the new branch you have created
- Create a `Pull Request`
  - Important - **make sure to validate your data before submitting this pull request, as bad data will be declined**
- Await review

There are of course more details to this, but please let me know if this does _not_ work for you.


### Installing in Linux

As per default the Linux package is a one-large binary package that can be run directly from whereever you have downloaded it - it will not install anything on system. If you want to have the _CRT_ application and icon available in your "Start" menu (not sure what this is called in Linux?), then you can install it with an application manager like e.g. **Gear Lever**. Just open **Gear Lever** and drag the _CRT_ file in to it, and afterwards you will be able to access it nice and easily:

<img width="902" height="578" alt="image" src="https://github.com/user-attachments/assets/13edb9d5-8b61-4259-bcc7-e0986d88ed51" />


### Compiling yourself

You can view the details in [BUILDING.md](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/blob/main/BUILDING.md)


### Controlling oscilloscope with keyboard/numpad

_CRT_ can control a network connected oscilloscope, if it is supported, with the keyboard/numpad. You can use these keys:

<img width="1051" height="358" alt="image" src="https://github.com/user-attachments/assets/8f339e2c-bf05-49bd-ab8d-9cad2a3b018b" />

If your oscilloscope is not in the list or it actually does not work properly, then please do investigate which **SCPI commands** for your specific oscilloscope model, as this varies quite a lot - even within same vendor. I do not know all oscilloscopes, so you will need to provide this data yourself. You can add and test the required data in the main Excel data file `Classic-Repair-Toolbox.xlsx` in the sheet `Oscilloscope`.


### Development tools used

_CRT_ has been developed in _Visual Studio Community 2026_. Where the old _Commodore_ project was primarily self-developed, then this new _Classic_ codebase has been developed primarily with GitHub Copilot, which is why I see myself more as a _conductor_ for this project, rather than the pure developer of this application - all credits to the people behind these LLM models 😁 As of March-2026 I have primarily used the _Gemini 3.1 Pro_ model, but also _Claude Sonnet 4.6_ and in some cases _GPT-5.3-Codex_ (these models will of course change for the future).

NuGet packages used:
- [Avalonia](https://avaloniaui.net/)
- [EPPlus](https://epplussoftware.com/)
- [Newtonsoft.Json](https://www.newtonsoft.com/json)
- [Velopack](https://github.com/velopack/velopack)


## Inspiration for building this application

I have been repairing Commodore 64/128 computers for some years, but I still consider myself as a _beginner_ in this world of hardware, and as you probably can guess (since I did this application) then I am more a software person. The hardware side of things is really relaxing for me, focussing on some physical hardware, troubleshooting, soldering, replacing and seeing a broken machine being revived is just so satisfying, so this is a _must-have_ for me to relax a little from all my software projects 😁

For my repairs I always forget _where_ and _what_ to check, and I struggle to find again all the relevant ressources and schematics to check, not to mention how to find the components in the schematics. I did often refer to the "Mainboards" section of [My Old Computer](https://myoldcomputer.nl/technical-info/mainboards/), and I noticed that Jeroen did have a prototype of an application named _Repair Help_, and it did have the easy layout I was looking for. However, it was never finalized from his side, so I took upon myself to create something similar, and a couple of years later (a lot of hiatus) I did come up with a very similar looking Windows application named **Commodore Repair Toolbox** (CRT).

After a year with _CRT_ and due to several questions about "_is it Windows only_", then I investigated if it was realistic for me to do a native porting to other systems. As I in the same time wanted to explore vibe-coding with the new LLM models, then I decided to give it a go... a complete rewrite based on a new platform (Avalonia), giving me a great opportunity to lurk out previous design flaws in the old project, which was almost completely "hand-written". So, here we are now with a completely new project and natively supporting **Windows**, **Linux** and **macOS** - nice.


## Screenshots

Main schematics:
<img width="902" height="555" alt="image" src="https://github.com/user-attachments/assets/ec67b241-2e08-46c8-ac27-21c17c795d1a" />

Overview where a lot of component information is garthered:
<img width="902" height="555" alt="image" src="https://github.com/user-attachments/assets/8db1fab6-55cc-45de-bfa0-3892f3145490" />

Resources relevant to the hardware and board:
<img width="902" height="555" alt="image" src="https://github.com/user-attachments/assets/bcb695c7-a7e3-414e-8ed0-cd7f43337ac4" />

Configuration options:
<img width="902" height="555" alt="image" src="https://github.com/user-attachments/assets/06f6a331-5d8a-4766-be01-ad9ee5be2d1a" />

Doing a few manual traces:
<img width="902" height="727" alt="image" src="https://github.com/user-attachments/assets/8291d990-5a20-4c8b-a537-fc5075a3235c" />

Component information popup:
<img width="902" height="580" alt="image" src="https://github.com/user-attachments/assets/ec811492-7c13-4542-bc80-a9479ba0d315" />
