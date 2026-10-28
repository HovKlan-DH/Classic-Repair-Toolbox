using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE NAME A FILE ADDED IN THE COMPONENT PANE GETS IN ITS BOARD (CRT 4.0 step 2, slice 3b - the
    // plan: "No 'File location' box. Every file belongs to its board").
    //
    // In its kind's usual folder (ContributionFileLocations.DefaultFolderFor - where every converted
    // board keeps that kind), and NEVER a name the board already has (BoardFileNames.FreeName): given
    // one, the file would replace that file for every row citing it - another component's picture.
    //
    //   - a picture or a file keeps the picked file's own name: "Component images/HotCPU.png";
    //   - a scope baseline is named as the published ones are, "<label>_<pin>_<signal>" with the
    //     picked file's type: "Scope baseline/IC107_1_PA3.jpg" - whatever the screenshot was called.
    //
    // Shaped so the board can keep it on every system and the server takes it: no sign a file system
    // refuses (a fixed set, not this machine's - a name made on Linux must open on Windows), and no
    // leading dot (the server refuses a hidden name, file.hidden_name).
    // ###########################################################################################
    internal static class ComponentPaneFileNames
    {
        private const string Refused = "\\/:*?\"<>|";

        public static string ForPicture(string sourcePath, IEnumerable<string> taken) =>
            ComponentPaneFileNames.InFolder(ContributionFileLocations.FileRowKind.ComponentPicture, ComponentPaneFileNames.Shaped(Path.GetFileName(sourcePath)), taken);

        public static string ForFile(string sourcePath, IEnumerable<string> taken) =>
            ComponentPaneFileNames.InFolder(ContributionFileLocations.FileRowKind.ComponentFile, ComponentPaneFileNames.Shaped(Path.GetFileName(sourcePath)), taken);

        // A picture of the board (its Pictures part, step 3): at the board's top, as every live board
        // keeps its pictures ("Board Layout 250407 NTSC.png") - the "Schematic images" window's place.
        public static string ForBoardPicture(string sourcePath, IEnumerable<string> taken) =>
            BoardFileNames.FreeName(ComponentPaneFileNames.Shaped(Path.GetFileName(sourcePath)), taken);

        // ###########################################################################################
        // A photo of one of the user's devices, in its "device_{id}" folder (CRT 4.0 step 4): a plain
        // name of its own - the picked file's, shaped as a board's is. Never a separator: on Linux and
        // macOS a "\" is part of a file's name, and turned into "/" by FreeName it took the copy out
        // of the device's folder (code review, 2026-10-28).
        // ###########################################################################################
        public static string ForDevicePhoto(string sourcePath, IEnumerable<string> taken) =>
            BoardFileNames.FreeName(ComponentPaneFileNames.Shaped(Path.GetFileName(sourcePath)), taken);

        // A file of the board itself (the board's "Files and links", slice 4a): "Board local files/<file>".
        public static string ForBoardFile(string sourcePath, IEnumerable<string> taken) =>
            ComponentPaneFileNames.InFolder(ContributionFileLocations.FileRowKind.BoardFile, ComponentPaneFileNames.Shaped(Path.GetFileName(sourcePath)), taken);

        public static string ForBaseline(string label, string pin, string? signal, string sourcePath, IEnumerable<string> taken)
        {
            string stem = string.Join(
                "_",
                new[] { label, pin, signal }
                    .Select(part => part?.Trim() ?? string.Empty)
                    .Where(part => part.Length > 0)
                    .Select(ComponentPaneFileNames.Safe));

            return ComponentPaneFileNames.InFolder(
                ContributionFileLocations.FileRowKind.ScopeBaseline,
                ComponentPaneFileNames.Shaped(stem + Path.GetExtension(sourcePath)),
                taken);
        }

        private static string InFolder(ContributionFileLocations.FileRowKind kind, string fileName, IEnumerable<string> taken) =>
            BoardFileNames.FreeName($"{ContributionFileLocations.DefaultFolderFor(kind)}/{fileName}", taken);

        // A file's name as the board can keep it: refused signs made "_", no leading dots, never empty.
        private static string Shaped(string? fileName)
        {
            string shaped = ComponentPaneFileNames.Safe(fileName ?? string.Empty).Trim().TrimStart('.').Trim();

            return shaped.Length == 0 ? "file" : shaped;
        }

        private static string Safe(string part)
        {
            var builder = new StringBuilder(part.Length);

            foreach (char character in part)
            {
                builder.Append(char.IsControl(character) || ComponentPaneFileNames.Refused.Contains(character) ? '_' : character);
            }

            return builder.ToString();
        }
    }
}
