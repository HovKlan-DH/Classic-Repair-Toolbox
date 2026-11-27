using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Handlers.DataHandling;

namespace Handlers.Collections
{
    // ###########################################################################################
    // A COLLECTION ATTACHED TO A WORKLOG (owner request, 2026-11-27: "I would like to be able to attach a collection to a
    // worklog"; asked, "A copy of its captures"): every capture COPIED into the worklog as a photo - in the page's order, its
    // comment where it was taken, the settings and the baseline's beside them, what it should read and the note on it
    // (CollectionSharing.WorklogCommentOf) - and, while the collection's page is up, that page as a link of interest. The
    // worklog then stands alone, whatever happens to the collection later.
    //
    // The writes are the worklog's own (WorklogAttachmentWriter - the oscilloscope capture's "Attach image to worklog" files
    // into an entry the same way); the collection window's "Attach to worklog" picks the worklog (Main.Collections.cs).
    // ###########################################################################################
    public static class CollectionWorklogAttach
    {
        // `Added`: the captures now in the worklog; `NotAdded`: those whose file was gone or could not be copied; `Linked`: the
        // page's link is in its links of interest.
        public sealed record Outcome(int Added, int NotAdded, bool Linked);

        // The captures to attach, in the page's order, each with its file on this computer - one with no file left out.
        public static IReadOnlyList<(LocalCapture Capture, string Path)> CapturesOf(LocalCollection collection, Func<string, string> capturePath)
        {
            ArgumentNullException.ThrowIfNull(collection);
            ArgumentNullException.ThrowIfNull(capturePath);

            return CollectionSharing.InPageOrder(collection.Captures)
                .Where(capture => DataTreeLayout.IsSha256(capture.Sha256))
                .Select(capture => (capture, capturePath(capture.Sha256)))
                .Where(pair => File.Exists(pair.Item2))
                .ToList();
        }

        // ###########################################################################################
        // Into an existing worklog of `workbookId` - each capture read fresh and written back, as AttachToEntry does.
        // ###########################################################################################
        public static Outcome ToEntry(int workbookId, int entryId, LocalCollection collection, Func<string, string> capturePath)
        {
            ArgumentNullException.ThrowIfNull(collection);

            IReadOnlyList<(LocalCapture Capture, string Path)> captures = CollectionWorklogAttach.CapturesOf(collection, capturePath);
            int added = 0;

            foreach ((LocalCapture capture, string path) in captures)
            {
                WorklogAttachmentWriter.AttachOutcome outcome = WorklogAttachmentWriter.AttachToEntry(
                    workbookId, entryId, path, WorklogAttachmentStorage.PhotoFilePrefix, CollectionSharing.WorklogCommentOf(capture));

                if (outcome == WorklogAttachmentWriter.AttachOutcome.Added)
                    added++;
            }

            bool linked = collection.IsOnline && collection.PageUrl is { } url &&
                          WorklogAttachmentWriter.AddLinkToEntry(workbookId, entryId, CollectionWorklogAttach.LinkHeadline(collection), url) == WorklogAttachmentWriter.AttachOutcome.Added;

            return new Outcome(added, collection.Captures.Count - added, linked);
        }

        // ###########################################################################################
        // Into a worklog being made in its editor ("Create new worklog") - `attach` and `link` are the editor's own (a draft is
        // written at its Save, or not at all).
        // ###########################################################################################
        public static Outcome ToEditor(LocalCollection collection, Func<string, string> capturePath, Func<string, string, bool> attach, Action<string, string> link)
        {
            ArgumentNullException.ThrowIfNull(collection);
            ArgumentNullException.ThrowIfNull(attach);
            ArgumentNullException.ThrowIfNull(link);

            int added = CollectionWorklogAttach.CapturesOf(collection, capturePath)
                .Count(pair => attach(pair.Path, CollectionSharing.WorklogCommentOf(pair.Capture)));

            bool linked = collection.IsOnline && collection.PageUrl is not null;

            if (linked)
                link(CollectionWorklogAttach.LinkHeadline(collection), collection.PageUrl!);

            return new Outcome(added, collection.Captures.Count - added, linked);
        }

        // The page as the worklog's link of interest names it.
        public static string LinkHeadline(LocalCollection collection) => $"Oscilloscope collection: {CollectionSharing.NameOf(collection)}";

        // What the collection window says once it is attached.
        public static string Said(Outcome outcome, string worklog)
        {
            ArgumentNullException.ThrowIfNull(outcome);

            string said = outcome.Added == 0
                ? $"No capture could be attached to {worklog}"
                : $"{outcome.Added} capture{(outcome.Added == 1 ? string.Empty : "s")} attached to {worklog} as photos";

            if (outcome.Linked)
                said += ", with the page's link under \"Links of interest\"";

            if (outcome.Added > 0 && outcome.NotAdded > 0)
                said += $" - {outcome.NotAdded} could not be, their files gone from this computer";

            return said + ".";
        }

        // What the attach dialog says goes in.
        public static string Summary(LocalCollection collection, int captures) =>
            $"{captures} capture{(captures == 1 ? string.Empty : "s")} of [{CollectionSharing.NameOf(collection)}], each as a photo with " +
            "its pin, settings and note" + (collection.IsOnline ? ", and the page's link" : string.Empty);

        public const string NothingToAttach = "None of its captures is on this computer any more - there is nothing to attach.";

        // The workbook has no worklog yet, and a new one needs the collection's board shown (its schematic image).
        public static string NoWorklogYet(LocalCollection collection) =>
            $"Show {collection.HardwareName} - {collection.BoardName} in CRT to attach the collection to a new worklog - its workbook " +
            "has no worklog to pick yet.";

        // "Create new worklog" picked, and the collection's board no longer shown by then.
        public static string BoardNoLongerShown(LocalCollection collection) =>
            $"{collection.HardwareName} - {collection.BoardName} is no longer shown in CRT - show it again to attach the collection to a new worklog.";

        // No workbook worked in on the collection's board to attach into.
        public static string NoWorkbook(LocalCollection collection) =>
            $"There is no workbook for {collection.HardwareName} - {collection.BoardName} to attach it to. Show that board in CRT, " +
            "press \"Add workbook\", then attach the collection.";
    }
}
