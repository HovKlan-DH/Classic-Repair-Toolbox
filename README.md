# Classic Repair Toolbox

_Classic Repair Toolbox_ (or **CRT** hence forward) is a utility tool for repairing and diagnosing vintage computers and peripherals.

The project is a direct spin-off from an older project, **Commodore Repair Toolbox** which now resides in a faint and distant memory only. The new _Classic_ project (compared to _Commodore_) was realized as a complete rewrite, to be able to support **Linux** and **macOS** natively, but also to be able to support more hardware and not focus primarily on Commodore (Amstrad, Spectrum?).


## What is it?

With _CRT_ you can easily view technical schematics, zoom, identify components, view chip pinouts, do manual circuit tracing, study datasheets, view oscilloscope images, ressources and various other information, helping you diagnosing and repairing old vintage hardware.

It is (for now) primarily dedicated to Commodore, and have a few built-in profiles for Commodore computers, but it can support any kind of hardware, as you can add your own data - e.g. other computers, radios, DIY electronics or whatever else you can imagine. It probably works the best, if the hardware is "simple" and have good documentation available, like schematics, and if it is something you need to revisit multiple times - then you can add the needed information yourself, and use it for easy future reference.


## Table of Contents

- [Installation and usage](#installation-and-usage)
- [Built-in hardware and boards](#built-in-hardware-and-boards)
- [Data contributions being worked on currently](#data-contributions-being-worked-on-currently)
- [Requirements](#requirements)
- [Help wanted](#help-wanted)
- [Contact developer](#contact-developer)
- [Technical topics](#technical-topics)
- [Information automatically collected by CRT](#information-automatically-collected-by-crt)
- [How to contribute with data to CRT GitHub repository?](#how-to-contribute-with-data-to-crt-github-repository)
- [Compiling for Linux](#compiling-for-linux)
- [Development tools used](#development-tools-used)
- [Inspiration for building this application](#inspiration-for-building-this-application)
- [Screenshots](#screenshots)


## Installation and usage

Download the newest normal (non-NETA) _CRT_ version from [Releases](https://github.com/HovKlan-DH/Commodore-Repair-Toolbox/releases), and install it afterwards. The installation folder cannot be set by the user, and is decided by installation system (Avalonia), but in the `Configuration` tab you can open the folder where it has its configuration and data files located.

If needed then the `data-root` folder can be changed via a commandline parameter, view [Commandline parameters](#commandline-parameters).

Depending on your configuration settings, then _CRT_ will check for newer data at application launch.

When a new version is released, you can update directly from within the application.


## Built-in hardware and boards

- **Commodore VIC-20**
  - 250403 (CR)
    - Would appreciate help with:
      - Oscilloscope baseline for PAL and NTSC
      - More data
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

- Amstrad CPC 664
  - "Rabs" is working on this new board (updated 2026-March-7)
    

## Requirements

- Operating systems supported:
  - **Windows 7** or newer (32-bit and 64-bit)
  - **macOS** (64-bit)
  - **Linux** (64-bit)

Note that .NET is embedded in application, which means you do not need to have this installed. It also does mean that even if you have .NET installed on your computer, then it will still use the one embedded in application. As .NET6 is the newest version supported on **Windows 7**, then this is the .NET version included with the Windows 32-bit installer.

> [!CAUTION]
> .NET6 has gone **End-of-Life in 2024**, and has not received any security hotfixes since then!

If possible then you should use the newer _CRT_ 64-bit installer, which embeds the newest available [.NET10 LTS](https://github.com/dotnet/core/blob/main/release-notes/10.0/README.md) (Long-Term-Support) at release date.


## Help wanted

I will for sure keep adding and enhancing data, but if this is only me providing data, then it will take many years before this will reach a "premium level" - if ever 😁 So, I really do hope that the community will contribute, so it quickly can become a good source of information.

Data contribution can be almost anything - tiny and trivial updates (spelling mistakes, wrong technical values or alike) or it can be huge new boards, but I really would like to get a massive amount of **quality** data, for the benefit of everyone using this. The goal is that it should have (most) relevant data in one place, so it would not be required to go and lookup for other data sources, but of course it also needs to be balanced a little, not overwelming with too much data 🤔

You can help specifically with these topics:
- Do you have higher-quality images of the used schematics?
- Do you have (better) datasheets or pinouts for any of the components?
- Do you see missing components in either the component list or as a highlight?
- Can you improve any data or fill in more technical details anywhere?


## Contact developer

There are several ways to get in contact with the developer:

- Direct communication via _Retro Hardware Discord_ channel, [https://discord.gg/HDWct2vxem](https://discord.gg/HDWct2vxem) (accept the invite on page)
- CRT "Feedback" tab
- GitHub [Issues](https://github.com/HovKlan-DH/Classic-Repair-Toolbox/issues)


## Technical topics

### Information automatically collected by CRT

I want to be transparent here, and inform that I am gathering information about your setup, at every application launch, where the application does a mandatory "check-in":

- IP address
  - Ex. `85.184.162.75`
  - Used for pinning countries on a worldmap
- Operating system version
  - Ex. `Microsoft Windows 10.0.19045`
  - Used for knowing where to put the most effort
- CPU architechture used (32-bit or 64-bit)
  - Ex. `64-bit`
  - Used for knowing how wide usage that pesky self-contained .NET6 has

I am allowing myself to gather this data for me to build the [CRT Fun facts](https://commodore-repair-toolbox.dk/funfacts/) page, which is some statistics on usage. As a developer, this is a personal motivational point to see countries using my application and of course one always hope for that "upwards trend usage"... which never happens 🤣 I find this limited non-personal data a fair amount to "pay" for using this application, taking in consideration of the effort being put in to this.


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


### Compiling for Linux

Here is a rudimentary guide to compile _CRT_ on Linux - the below is based on Fedora 43, and there can probably be variations depending on your setup.

- Fork the _CRT_ GitHub repository
- Clone the fork to your local computer
- Make sure .NET 10 SDK is installed
  - Download from here, https://dotnet.microsoft.com/en-us/download/dotnet/10.0
- Compile RELEASE build
  - `dotnet publish -c Release -f net10.0 --self-contained`
- Run application:  
  - `./bin/Debug/net10.0/Classic-Repair-Toolbox`


### Development tools used

_CRT_ has been developed in _Visual Studio Community 2026_. Its codebase has been developed primarily with GitHub Copilot, which is why I see myself more as a _conductor_ rather than the pure developer of this application - all credits to the people behind these LLM models 😁

NuGet packages used:
- [Avalonia](https://avaloniaui.net/)
- [EPPlus](https://epplussoftware.com/)
- [Newtonsoft.Json](https://www.newtonsoft.com/json)
- [Velopack](https://github.com/velopack/velopack)


## Inspiration for building this application

I have been repairing Commodore 64/128 computers for some years, but I still consider myself as a _beginner_ in this world of hardware, and as you probably can guess (since I did this application) then I am more a software person. I always forget _where_ and _what_ to check, and I struggle to find again all the relevant ressources and schematics to check, not to mention how to find the components in the schematics. I did often refer to the "Mainboards" section of [My Old Computer](https://myoldcomputer.nl/technical-info/mainboards/), and I noticed that Jeroen did have a prototype of an application named _Repair Help_, and it did have the easy layout I was looking for. However, it was never finalized from his side, so I took upon myself to create something similar, and a couple of years later (a lot of hiatus) I did come up with a very similar looking Windows application named **Commodore Repair Toolbox** (CRT).

After a year with _CRT_ and due to several questions about "_is it Windows only_", then I investigated if it was realistic for me to do a native porting to other systems. As I in the same time wanted to explore vibe-coding with the new LLM models, then I decided to give it a go... a complete rewrite based on a new platform (Avalonia), giving me a great opportunity to lurk out previous design flaws in the old project, which was almost completely "hand-written". So, here we are now with a completely new project and natively supporting **Windows**, **Linux** and **macOS** - nice.


## Screenshots

To come when version **1.0.0** gets released ...
