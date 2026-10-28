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
    // What was dropped decides what is done with it - the "Schematic images" or the "KiCad data"
    // window opens with it being added, as its own drop target would add it:
    //
    //   - A FOLDER holding a KiCad file CRT reads is the KiCad project. One holding none, but pictures
    //     at its top (a folder of board scans), is those pictures. One holding neither is NOTHING, said
    //     for what it is - an older KiCad's or another program's project told what to do (code review,
    //     2026-10-28: tried as the KiCad project, a published board got a draft made for nothing).
    //   - A KiCad FILE on its own is no project (code review, 2026-10-28): taken as its folder, a file
    //     from Downloads imported every project under Downloads, and a sheet from Pages/ only Pages/.
    //   - PICTURE FILES (the types CRT shows, never a hidden one - macOS's "._IMG_0001.JPG",
    //     ContributionPackaging.IsAddablePictureFile) are pictures; anything else among them is left be.
    //   - Nothing of these: said, and nothing opens.
    //
    // Of several folders the first is taken - one KiCad project at a time. The file system is asked
    // through the functions handed in, so the rule is tested without one.
    // ###########################################################################################
    internal sealed record BoardDrop(BoardDropKind Kind, IReadOnlyList<string> Paths, string Problem = "")
    {
        public const string NothingUsable =
            "Drop pictures of the board (PNG, JPG, GIF, BMP or WEBP files) or its KiCad project folder to add them.";

        public const string KiCadFileNotFolder =
            "Drop the folder that holds the KiCad project, not one of its files - CRT takes the project's files from it, its sub-folders included.";

        public static BoardDrop Of(
            IReadOnlyList<string> dropped,
            Func<string, bool> isFolder,
            Func<string, bool> holdsKiCad,
            Func<string, IReadOnlyList<string>> filesAtTop,
            Func<string, string> whyNothingIn)
        {
            ArgumentNullException.ThrowIfNull(dropped);

            List<string> paths = dropped.Where(path => !string.IsNullOrWhiteSpace(path)).ToList();

            if (paths.FirstOrDefault(isFolder) is { } folder)
            {
                if (holdsKiCad(folder))
                {
                    return new BoardDrop(BoardDropKind.KiCad, [folder]);
                }

                List<string> scans = filesAtTop(folder).Where(ContributionPackaging.IsAddablePictureFile).ToList();

                return scans.Count > 0
                    ? new BoardDrop(BoardDropKind.Pictures, scans)
                    : new BoardDrop(BoardDropKind.Nothing, [], whyNothingIn(folder));
            }

            List<string> pictures = paths.Where(ContributionPackaging.IsAddablePictureFile).ToList();

            if (pictures.Count > 0)
            {
                return new BoardDrop(BoardDropKind.Pictures, pictures);
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
