[Wiki Home](Home)

What CRT is built with.

---

CRT is written in C# with .NET 10, using [Avalonia](https://avaloniaui.net/) for the user interface, and is developed in _Visual Studio Community 2026_ and with _Visual Studio Code_ for the agentic approach.

Where the old _Commodore_ project was primarily self-developed, this _Classic_ codebase has been developed primarily with AI assistance, which is why I see myself more as a _conductor_ for this project rather than the pure developer of this application - all credits to the people behind these LLM models 😁

I started out with GitHub Copilot, and now (2026-September) use Claude Code.

NuGet packages used:

- [Avalonia](https://avaloniaui.net/) - the user interface
- [EPPlus](https://epplussoftware.com/) - reads and writes the Excel data files
- [QuestPDF](https://www.questpdf.com/) - the workbook PDF export
- [Velopack](https://github.com/velopack/velopack) - the in-application updater

Bundled with the application:

- [minipro](https://gitlab.com/DavidGriffith/minipro) - the command-line programmer behind the IC
  logic tests, by David Griffith. CRT bundles a Windows build made from
  [this fork](https://github.com/KevinWelton/minipro-win), which talks to the programmer through
  WinUSB directly so there is no driver to install. It is a separate program that CRT calls, not a
  part of CRT itself.
- [Font Awesome 7 Free](https://fontawesome.com/) - the icons used throughout the application and
  in the exported PDF.

---

## Licensing

CRT itself is under the **GNU General Public License v3**.

Most of the above is under permissive licenses that ask for little more than credit, but two of
them are worth knowing about if you ever want to build on CRT yourself:

- **EPPlus** is free for noncommercial use only. It is not open source.
- **QuestPDF** is free for individuals, for smaller organisations, and for open-source projects
  such as this one, but not for everybody.

None of that affects you if you are simply using CRT to repair your own machines - that is exactly
what it is free for. It matters only if you plan to take the source and sell something built from
it, in which case you would need your own EPPlus license or a replacement for it.

The full details, and the license of everything CRT ships, are in `THIRD-PARTY-NOTICES.md` and
`LICENSE.txt` in the [repository](https://github.com/HovKlan-DH/Classic-Repair-Toolbox).
