using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // What files dropped on the Contribute tab are: pictures of the board, its KiCad project, or neither.
    internal enum BoardDropKind
    {
        Nothing,
        Pictures,
        KiCad,
    }

    // ###########################################################################################
    // FILES DROPPED ON THE CONTRIBUTE TAB (CRT 4.0 step 3, slice 3e - the plan: "Pictures: today's
    // Schematic images content, moved inline ... KiCad: drop the project folder and get a summary").
    // What was dropped decides what is done with it - pictures go into the Pictures part as its "Add
    // pictures" adds them, a KiCad project opens the "KiCad data" window importing it:
    //
    //   - A FOLDER holding a KiCad file CRT reads is the KiCad project. One holding none, but pictures
    //     at its top (a folder of board scans), is those pictures. One holding neither is NOTHING, said
    //     for what it is - an older KiCad's or another program's project told what to do (code review,
    //     2026-10-28: tried as the KiCad project, a published board got a draft made for nothing).
    //   - A KiCad FILE on its own is no project (code review, 2026-10-28): taken as its folder, a file
    //     from Downloads imported every project under Downloads, and a sheet from Pages/ only Pages/.
    //   - PICTURE FILES (the types CRT shows, never a hidden one - macOS's "._IMG_0001.JPG",
    //     ContributionPackaging.IsAddablePictureFile) are pictures.
    //   - Nothing of these: said, and nothing opens.
    //
    // *** EVERYTHING DROPPED COUNTS, AND WHAT IS LEFT OUT IS SAID (code review, 2026-10-28). *** The
    // first folder used to decide alone: two loose photos dropped with a folder of scans were skipped
    // without a word. Now a KiCad project folder is taken on its own - one project at a time - and the
    // rest is said to be left (Note); else every picture dropped, loose or at a folder's top, goes in
    // together, and whatever held no picture is said. The file system is asked through the functions
    // handed in, so the rule is tested without one.
    // ###########################################################################################
    internal sealed record BoardDrop(BoardDropKind Kind, IReadOnlyList<string> Paths, string Problem = "", string Note = "")
    {
        public const string NothingUsable =
            "Drop pictures of the board (PNG, JPG, GIF, BMP or WEBP files) or its KiCad project folder to add them.";

        public const string KiCadFileNotFolder =
            "Drop the folder that holds the KiCad project, not one of its files - CRT takes the project's files from it, its sub-folders included.";

        // Said beside a KiCad project dropped with other things - one project at a time, pictures apart.
        public static string LeftBesideKiCad(int others) =>
            others == 1
                ? "The KiCad project was taken - the other item dropped with it was left. Drop pictures on their own to add them."
                : $"The KiCad project was taken - the other {others.ToString(System.Globalization.CultureInfo.InvariantCulture)} items dropped with it were left. Drop pictures on their own to add them.";

        // Said beside pictures taken: the items dropped that held none.
        public static string LeftWithoutPictures(int count) =>
            count == 1
                ? "1 item dropped holds no picture CRT can show and was left out."
                : $"{count.ToString(System.Globalization.CultureInfo.InvariantCulture)} items dropped hold no picture CRT can show and were left out.";

        public static BoardDrop Of(
            IReadOnlyList<string> dropped,
            Func<string, bool> isFolder,
            Func<string, bool> holdsKiCad,
            Func<string, IReadOnlyList<string>> filesAtTop,
            Func<string, string> whyNothingIn)
        {
            ArgumentNullException.ThrowIfNull(dropped);

            List<string> paths = dropped.Where(path => !string.IsNullOrWhiteSpace(path)).ToList();
            List<string> folders = paths.Where(isFolder).ToList();

            if (folders.FirstOrDefault(holdsKiCad) is { } project)
            {
                int others = paths.Count - 1;
                return new BoardDrop(BoardDropKind.KiCad, [project], Note: others > 0 ? BoardDrop.LeftBesideKiCad(others) : string.Empty);
            }

            var pictures = new List<string>();
            int leftOut = 0;

            foreach (string path in paths)
            {
                List<string> found = folders.Contains(path)
                    ? filesAtTop(path).Where(ContributionPackaging.IsAddablePictureFile).ToList()
                    : ContributionPackaging.IsAddablePictureFile(path) ? [path] : [];

                if (found.Count == 0)
                {
                    leftOut++;
                }

                pictures.AddRange(found.Where(picture => !pictures.Contains(picture, StringComparer.OrdinalIgnoreCase)));
            }

            if (pictures.Count > 0)
            {
                return new BoardDrop(BoardDropKind.Pictures, pictures, Note: leftOut > 0 ? BoardDrop.LeftWithoutPictures(leftOut) : string.Empty);
            }

            if (folders.Count > 0)
            {
                return new BoardDrop(BoardDropKind.Nothing, [], whyNothingIn(folders[0]));
            }

            return paths.Any(ComponentListBuilder.IsSupportedKiCadRawFile)
                ? new BoardDrop(BoardDropKind.Nothing, [], BoardDrop.KiCadFileNotFolder)
                : new BoardDrop(BoardDropKind.Nothing, [], BoardDrop.NothingUsable);
        }

        // ###########################################################################################
        // The rule on the real file system. The folder is walked ONCE (KiCadRawFileScanner.Probe -
        // second code review, 2026-10-28: a folder holding nothing usable was walked in full twice):
        // stopping at its first KiCad file - the import walks it in full itself - and else keeping only
        // its .sch and .brd files, to say why nothing can be imported: an older KiCad's or another
        // program's project told what to do.
        // ###########################################################################################
        public static BoardDrop OnDisk(IReadOnlyList<string> dropped)
        {
            var probes = new Dictionary<string, KiCadFolderProbe>(StringComparer.Ordinal);
            KiCadFolderProbe ProbeOf(string folder) =>
                probes.TryGetValue(folder, out KiCadFolderProbe? probe) ? probe : probes[folder] = KiCadRawFileScanner.Probe(folder);

            return BoardDrop.Of(
                dropped,
                Directory.Exists,
                folder => ProbeOf(folder).HoldsKiCad,
                folder =>
                {
                    try
                    {
                        return Directory.GetFiles(folder).OrderBy(path => path, StringComparer.OrdinalIgnoreCase).ToList();
                    }
                    catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                    {
                        return [];
                    }
                },
                folder =>
                {
                    KiCadFolderProbe probe = ProbeOf(folder);
                    List<string> why = KiCadImportSummary.CannotRead(probe.SchematicsAndBoards, probe.OlderKiCad, []);

                    return why.Count > 0
                        ? "No KiCad files CRT can read in that folder. " + string.Join(" ", why)
                        : "That folder holds neither KiCad files nor pictures CRT can show. " + BoardDrop.NothingUsable;
                });
        }
    }
}
