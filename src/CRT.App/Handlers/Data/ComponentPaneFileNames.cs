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
