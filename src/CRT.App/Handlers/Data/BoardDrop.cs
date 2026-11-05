using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Handlers.DataHandling
{
    // What files dropped on "Board: Schematic images" are: images of the board, or nothing it can take.
    internal enum BoardDropKind
    {
        Nothing,
        Pictures,
    }

    // ###########################################################################################
    // FILES DROPPED ON "BOARD: SCHEMATIC IMAGES" (CRT 4.0 step 3, slice 3e; only that page since
    // 2026-11-03, TabDrafts.Drop.cs). What was dropped decides what is taken - as "Add schematic image(s)"
    // adds them:
    //
    //   - PICTURE FILES (the types CRT shows, never a hidden one - macOS's "._IMG_0001.JPG",
    //     ContributionPackaging.IsAddablePictureFile) are pictures, and so are those at the top of a
    //     folder (a folder of board scans).
    //   - A FOLDER holding a KiCad file CRT reads is a KiCad project, which goes on "KiCad: Files" - it is
    //     no folder of scans, whatever images lie in it. Dropped alone it is told where to go (2026-11-03).
    //   - A folder holding neither is NOTHING, said for what it is - an older KiCad's or another program's
    //     project told what to do.
    //   - A KiCad FILE on its own is told the same way (code review, 2026-10-28).
    //   - Either dropped WITH images is told too (KiCadLeftOut; code review, 2026-11-04: it was counted as
    //     an item holding no image - of a folder that may well hold some, and never said where it goes).
    //   - Nothing of these: said, and nothing is added.
    //
    // *** EVERYTHING DROPPED COUNTS, AND WHAT IS LEFT OUT IS SAID (code review, 2026-10-28). *** The
    // first folder used to decide alone: two loose photos dropped with a folder of scans were skipped
    // without a word. Every picture dropped, loose or at a folder's top, goes in together, and whatever
    // held no picture is said (Note). The file system is asked through the functions handed in, so the
    // rule is tested without one.
    // ###########################################################################################
    internal sealed record BoardDrop(BoardDropKind Kind, IReadOnlyList<string> Paths, string Problem = "", string Note = "")
    {
        // Named as the page names them - "schematic images" (code review, 2026-10-30: the part was renamed
        // on 2026-10-29, and these still said "pictures").
        public const string NothingUsable =
            "Drop schematic images of the board (PNG, JPG, GIF, BMP or WEBP files) here to add them.";

        public const string KiCadGoesElsewhere =
            "That is a KiCad project - add it under \"KiCad: Files\" on the left. Only schematic images can be dropped here.";

        // Said beside pictures taken: the KiCad files dropped with them.
        public const string KiCadLeftOut =
            "The KiCad files dropped with the images were left out - add them under \"KiCad: Files\" on the left.";

        // Said beside pictures taken: the items dropped that held none.
        public static string LeftWithoutPictures(int count) =>
            count == 1
                ? "1 item dropped holds no image CRT can show and was left out."
                : $"{count.ToString(System.Globalization.CultureInfo.InvariantCulture)} items dropped hold no image CRT can show and were left out.";

        private static string NoteOf(int leftOut, bool kiCadLeftOut)
        {
            var note = new List<string>(2);

            if (leftOut > 0)
            {
                note.Add(BoardDrop.LeftWithoutPictures(leftOut));
            }

            if (kiCadLeftOut)
            {
                note.Add(BoardDrop.KiCadLeftOut);
            }

            return string.Join(" ", note);
        }

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
            List<string> projects = folders.Where(holdsKiCad).ToList();

            var pictures = new List<string>();
            int leftOut = 0;
            bool kiCadLeftOut = false;

            foreach (string path in paths)
            {
                if (projects.Contains(path) || (!folders.Contains(path) && ComponentListBuilder.IsSupportedKiCadRawFile(path)))
                {
                    kiCadLeftOut = true;
                    continue;
                }

                List<string> found = folders.Contains(path) ? filesAtTop(path).Where(ContributionPackaging.IsAddablePictureFile).ToList()
                    : ContributionPackaging.IsAddablePictureFile(path) ? [path] : [];

                if (found.Count == 0)
                {
                    leftOut++;
                }

                pictures.AddRange(found.Where(picture => !pictures.Contains(picture, StringComparer.OrdinalIgnoreCase)));
            }

            if (pictures.Count > 0)
            {
                return new BoardDrop(BoardDropKind.Pictures, pictures, Note: BoardDrop.NoteOf(leftOut, kiCadLeftOut));
            }

            if (projects.Count > 0 || paths.Any(ComponentListBuilder.IsSupportedKiCadRawFile))
            {
                return new BoardDrop(BoardDropKind.Nothing, [], BoardDrop.KiCadGoesElsewhere);
            }

            return folders.Count > 0
                ? new BoardDrop(BoardDropKind.Nothing, [], whyNothingIn(folders[0]))
                : new BoardDrop(BoardDropKind.Nothing, [], BoardDrop.NothingUsable);
        }

        // ###########################################################################################
        // The rule on the real file system. The folder is walked ONCE (KiCadRawFileScanner.Probe -
        // second code review, 2026-10-28: a folder holding nothing usable was walked in full twice):
        // stopping at its first KiCad file, and else keeping only its .sch and .brd files, to say why
        // nothing can be taken: an older KiCad's or another program's project told what to do.
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
                        : "That folder holds neither KiCad files nor images CRT can show. " + BoardDrop.NothingUsable;
                });
        }
    }
}
