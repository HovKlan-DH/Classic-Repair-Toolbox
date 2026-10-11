using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A BOARD'S SCHEMATIC IMAGES RENAMED - WHAT THIS MACHINE KEEPS UNDER THEIR NAMES FOLLOWS
    // (owner decisions, 2026-10-10: renaming an image, "And fix all", a published rename "must
    // also change for everyone updating to this data", and the image file kept with it "to be
    // unambiguous").
    //
    // CRT keeps things of the user's own that name a board's images, and they would quietly lose
    // the image when it is renamed:
    //   - every worklog's marked area (WorklogEntryRecord.SchematicName), and each workbook's own
    //     order of the board's images (WorkbookRecord.SchematicOrder);
    //   - the user's thumbnail order on the Schematics tab (UserSettings.SchematicsOrderByBoard);
    //   - the image last looked at, opened again with the board (LastSchematicByBoard);
    //   - the images hidden on the Configuration tab ("hardware|board|schematic" catalogue keys).
    //
    // *** CAUGHT UP WHEN THE BOARD IS SHOWN (CatchUp, from Main's board load), BY THE IMAGE FILE. ***
    // CRT remembers each board's images - name and file - as it last showed them
    // (UserSettings.SchematicImagesByBoard). When the board comes back different, every name is
    // followed to wherever its FILE is now (SchematicRenameLog.Resolver): a rename published, the
    // user's own draft rename, the draft discarded or rejected, "view boards as officially
    // published", a rename pushed back out of BETA, two images swapping names, a name later given to
    // another image. Without a memory - the first time this CRT shows the board - the board's
    // published list (its .json's "Schematic renames") says where names went instead.
    //
    // Each change is in place - nothing else in a worklog or a setting changes.
    // ###########################################################################################
    internal static class LocalSchematicRename
    {
        // ###########################################################################################
        // Brings this machine's worklogs and settings for the board in line with its images as
        // shown now (`board`), then remembers them. Nothing is read or moved when the board is
        // shown exactly as last time - the board's published list (`readPublished`, its .json) is
        // read only when it can be needed (code review, 2026-10-11: it was parsed on every board
        // load, on the UI thread) - or, the first time, when that list names no rename.
        //
        // *** A BOARD SHOWN WITH NO IMAGES CHANGES NOTHING (code review, 2026-10-11). *** It says
        // nothing about where its images went - its workbook could not be read, or a new board's
        // draft is shown "as officially published", which has no board - and remembering it forgot
        // the images, so the next rename could not be followed by its file.
        //
        // *** NOTHING IS REMEMBERED UNTIL EVERYTHING HAS MOVED. *** When a worklog cannot be written,
        // the worklogs are put back as they were (RenameInWorklogs) and the board is not remembered
        // as shown, so the next showing tries again from the same place.
        // ###########################################################################################
        public static void CatchUp(
            string hardwareName,
            string boardName,
            IReadOnlyList<SchematicImage> board,
            Func<IReadOnlyList<SchematicRenameRecord>> readPublished)
        {
            if (board.Count == 0)
            {
                return;
            }

            string boardKey = CatalogueVisibility.BuildKey(hardwareName, boardName);
            IReadOnlyList<SchematicImage>? remembered = UserSettings.GetSchematicImages(boardKey);

            if (remembered is not null && SchematicRenameLog.SameImages(remembered, board))
            {
                return;
            }

            IReadOnlyList<SchematicRenameRecord> published = readPublished();

            if (remembered is not null || published.Count > 0)
            {
                Func<string, string?> resolve = SchematicRenameLog.Resolver(board, remembered, published);

                if (!LocalSchematicRename.RenameInWorklogs(boardKey, resolve))
                {
                    Logger.Warning($"The worklogs of [{boardKey}] could not follow its renamed schematic images - tried again when the board is next shown");
                    return;
                }

                LocalSchematicRename.RenameInSettings(hardwareName, boardName, resolve);
            }

            UserSettings.SetSchematicImages(boardKey, board);
        }

        // ###########################################################################################
        // `names` with each name `resolve` moves replaced where it stood - or null when it moves
        // none. A name already in the list is kept once, where it first stood.
        // ###########################################################################################
        internal static List<string>? Renamed(IReadOnlyList<string>? names, Func<string, string?> resolve)
        {
            if (names is null)
            {
                return null;
            }

            bool moved = false;
            var result = new List<string>();

            foreach (string name in names)
            {
                string? now = resolve(name);
                string kept = now ?? name;

                moved |= now is not null;

                if (!result.Any(existing => string.Equals(existing.Trim(), kept.Trim(), StringComparison.OrdinalIgnoreCase)))
                {
                    result.Add(kept);
                }
            }

            return moved ? result : null;
        }

        // The user's settings for the board: thumbnail order, the image last looked at, and the
        // images hidden on the Configuration tab.
        internal static void RenameInSettings(string hardwareName, string boardName, Func<string, string?> resolve)
        {
            string boardKey = CatalogueVisibility.BuildKey(hardwareName, boardName);

            if (LocalSchematicRename.Renamed(UserSettings.GetSchematicsOrder(boardKey), resolve) is { } order)
            {
                UserSettings.SetSchematicsOrder(boardKey, order);
            }

            if (UserSettings.GetLastSchematicForBoard(boardKey) is { } lastName && resolve(lastName) is { } lastNow)
            {
                UserSettings.SetLastSchematicForBoard(boardKey, lastNow);
            }

            // All at once: two hidden images swapping names stay hidden, both of them.
            string prefix = CatalogueVisibility.BuildKey(hardwareName, boardName, string.Empty);

            var hiddenMoves = UserSettings.CatalogueUncheckedKeysSnapshot
                .Where(key => CatalogueVisibility.IsSchematicKey(key) && key.StartsWith(prefix, StringComparison.OrdinalIgnoreCase))
                .Select(key => (Key: key, Now: resolve(key[prefix.Length..])))
                .Where(move => move.Now is not null)
                .ToList();

            foreach ((string key, _) in hiddenMoves)
            {
                UserSettings.SetCatalogueKeyChecked(key, true);
            }

            foreach ((_, string? now) in hiddenMoves)
            {
                UserSettings.SetCatalogueKeyChecked(CatalogueVisibility.BuildKey(hardwareName, boardName, now!), false);
            }
        }

        // ###########################################################################################
        // Every workbook of the board: its worklogs marked on an image that moved, and its own image
        // order. True when all of it was written.
        //
        // *** ALL OR NOTHING (code review, 2026-10-11). *** A worklog that could not be written kept
        // its old name while the board was remembered as shown - so nothing followed it again, and
        // its marked area showed on no image. Now a write that fails puts back every worklog and
        // order this call had already moved, and CatchUp remembers nothing: the next showing starts
        // again from the same place. (Carrying on from half-way would not do - with two images
        // swapping names, a worklog already moved would be moved back.)
        //
        // Each workbook's worklogs are read once and each moved one written once
        // (WorklogManager.UpdateEntries) - it read the whole workbook again for every worklog.
        // ###########################################################################################
        internal static bool RenameInWorklogs(string boardKey, Func<string, string?> resolve)
        {
            var undo = new List<Action>();

            foreach (WorkbookRecord workbook in WorklogManager.GetWorkbooksForBoard(boardKey))
            {
                int workbookId = workbook.Id;
                var moved = new List<(WorklogEntryRecord Entry, string OldName)>();

                foreach (WorklogEntryRecord entry in WorklogManager.GetEntries(workbookId))
                {
                    if (string.IsNullOrWhiteSpace(entry.SchematicName) || resolve(entry.SchematicName) is not { } now)
                    {
                        continue;
                    }

                    moved.Add((entry, entry.SchematicName));
                    entry.SchematicName = now;
                }

                int written = moved.Count == 0 ? 0 : WorklogManager.UpdateEntries(workbookId, moved.ConvertAll(item => item.Entry));

                if (written > 0)
                {
                    List<(WorklogEntryRecord Entry, string OldName)> back = moved.GetRange(0, written);

                    undo.Add(() =>
                    {
                        foreach ((WorklogEntryRecord entry, string oldName) in back)
                        {
                            entry.SchematicName = oldName;
                        }

                        WorklogManager.UpdateEntries(workbookId, back.ConvertAll(item => item.Entry));
                    });
                }

                if (written < moved.Count)
                {
                    LocalSchematicRename.Undo(undo);
                    return false;
                }

                if (LocalSchematicRename.Renamed(workbook.SchematicOrder, resolve) is { } order)
                {
                    var orderBefore = new List<string>(workbook.SchematicOrder);

                    if (!WorklogManager.UpdateWorkbookSchematicOrder(workbookId, order))
                    {
                        LocalSchematicRename.Undo(undo);
                        return false;
                    }

                    undo.Add(() => WorklogManager.UpdateWorkbookSchematicOrder(workbookId, orderBefore));
                }
            }

            return true;
        }

        // Puts back what was written, the last first.
        private static void Undo(List<Action> undo)
        {
            for (int i = undo.Count - 1; i >= 0; i--)
            {
                undo[i]();
            }
        }
    }
}
