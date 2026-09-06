# Third-party notices

Classic Repair Toolbox (CRT) is licensed under the GNU General Public License v3 - see
[LICENSE.txt](LICENSE.txt), and the "Additional permission under GNU GPL version 3 section 7"
clause at the end of it, which is what makes the two source-available components below usable
in a GPL-3.0 application at all.

This file lists what CRT ships or bundles, and under what terms. It covers the application only.
The hardware reference data (schematics, datasheets, component images and similar) is contributed
material with its own origins and rights holders - see the License section of
[README.md](README.md).

---

## NuGet packages

| Package | License | Ships in the build |
| --- | --- | --- |
| [Avalonia](https://avaloniaui.net/) (plus `Avalonia.Desktop`, `Avalonia.Themes.Fluent`, `Avalonia.Fonts.Inter`, `Avalonia.Controls.ColorPicker`) | MIT | Yes |
| [AvaloniaUI.DiagnosticsSupport](https://github.com/AvaloniaUI/Avalonia) | MIT | Yes |
| [Velopack](https://github.com/velopack/velopack) | MIT | Yes |
| [EPPlus](https://epplussoftware.com/) | **Polyform Noncommercial 1.0.0** | Yes - see below |
| [QuestPDF](https://www.questpdf.com/) | **Dual-licensed; used under the Community License** | Yes - see below |
| [SonarAnalyzer.CSharp](https://www.sonarsource.com/) | LGPL-3.0 | No - `PrivateAssets="all"`, build-time analyzer only |

### EPPlus - Polyform Noncommercial 1.0.0

Reads and writes the Excel data files ([Handlers/Data/BoardDataReader.cs](Handlers/Data/BoardDataReader.cs),
[BoardDataWriter.cs](Handlers/Data/BoardDataWriter.cs), [DataManager.cs](Handlers/Data/DataManager.cs)).

**This is not an open-source license.** Only noncommercial purposes are permitted. CRT uses it
under a personal, noncommercial grant, declared in code as
`ExcelPackage.License.SetNonCommercialPersonal(...)`.

**What this means if you fork CRT:** the GPL-3.0 grants you the right to use and distribute this
software commercially. EPPlus's own license does not. Those rights are not CRT's to give, because
CRT is not the copyright holder of EPPlus. If you intend any commercial use of a fork, you must
either obtain a commercial EPPlus license from EPPlus Software, or replace EPPlus with a
permissively licensed reader/writer (ClosedXML and NPOI are the usual alternatives).

License text: <https://polyformproject.org/licenses/noncommercial/1.0.0>

### QuestPDF - Community License

Generates the workbook PDF export ([Handlers/Data/WorkbookPdfExporter.cs](Handlers/Data/WorkbookPdfExporter.cs)).
`QuestPDF.Settings.License` is set to `LicenseType.Community` once at startup in
[Main/App.axaml.cs](Main/App.axaml.cs).

QuestPDF is dual-licensed. The Community License is free for individuals, for organisations with
annual gross revenue under USD 1,000,000, and for charitable, academic and eligible open-source
projects - which is what CRT qualifies under. Public-sector entities and publicly traded companies
are not eligible regardless of revenue.

The Community License states that it "is not an OSI-approved open-source license", permits
redistribution only as a compiled component of an application, and does not grant a right to
redistribute QuestPDF in source form. QuestPDF's proprietary notices must not be removed or
obscured.

**What this means if you fork CRT:** if you do not meet the eligibility criteria above, you need
your own commercial license from QuestPDF. Again, that is not a right CRT can pass on.

Licensing overview: <https://www.questpdf.com/pricing>

---

## Bundled binaries

### minipro

`Assets/MiniPro/win-x64/minipro.exe` - the command-line programmer used for the logic/PLA vector
tests. Invoked as a **separate subprocess** by
[Handlers/MiniPro/MiniproProcessRunner.cs](Handlers/MiniPro/MiniproProcessRunner.cs); it is not
linked into CRT, so this is mere aggregation under GPL-3.0 section 5.

- **Upstream:** <https://gitlab.com/DavidGriffith/minipro> (David Griffith)
- **Windows fork built from:** <https://github.com/KevinWelton/minipro-win>
- **License:** GPL-3.0-or-later ("either version 3 of the License, or (at your option) any later
  version"), which is compatible with CRT's own GPL-3.0.

**The bundled binary is modified.** It is cross-compiled with a native WinUSB backend (no libusb
DLL, no Zadig) and carries two patches:

1. `device->vector_count` widened from `uint8_t` to `uint32_t`. Logic tests with more than 255
   vectors otherwise wrap modulo 256 to zero and silently report a pass - a correctness bug with
   safety consequences for the user's conclusions about a chip.
2. `<Shlobj.h>` changed to `shlobj.h`, a case shim for mingw-w64.

**Source availability (GPL-3.0 section 6).** Because a modified GPL binary is distributed, the
corresponding source must be offered to recipients. The upstream and fork URLs above, together
with the two patches described here, are that offer. See
[Assets/MiniPro/win-x64/minipro.txt](Assets/MiniPro/win-x64/minipro.txt) for the same information
alongside the binary itself.

---

## Bundled fonts

### Font Awesome 7 Free

`Assets/Fonts/Font Awesome 7 Free-Regular-400.otf` and
`Assets/Fonts/Font Awesome 7 Free-Solid-900.otf`, both version 7.3.1 and both the unmodified
files from the official desktop package, compiled into the assembly as `AvaloniaResource` and
used for icon glyphs throughout the UI and in the exported PDF.

- **License:** SIL Open Font License 1.1 - this is the license that applies to Font Awesome Free's
  **font files**. (Font Awesome Free is tri-licensed: CC BY 4.0 covers icons packaged as SVG/JS,
  and MIT covers non-font, non-icon files. Neither of those is used here.)
- **Copyright:** Fonticons, Inc. - <https://fontawesome.com>
- **License text:** [Assets/Fonts/LICENSE.txt](Assets/Fonts/LICENSE.txt), shipped next to the fonts.

SIL OFL 1.1 is GPL-compatible. The license text is included because the OTF files themselves carry
no embedded license or licenseURL name-table entry, so the notice has to travel beside them.

---

## The .NET runtime

Self-contained builds bundle the .NET 10 runtime, which is MIT licensed
(<https://github.com/dotnet/runtime>), with its own third-party notices at
<https://github.com/dotnet/runtime/blob/main/THIRD-PARTY-NOTICES.TXT>.
