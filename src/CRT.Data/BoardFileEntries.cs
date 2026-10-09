using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A BOARD'S FILES, AND WHAT THE NEXT STEP DOES TO EACH (owner request, 2026-09-28): "it would
    // be nice to see a file-structure for all existing files in BETA or PROD, and then a
    // highlighting of files changed (added, removed, changed) ... A new system should just show how
    // the new file-system will look like ... this should include the shared folders also".
    //
    // Both of the Maintainer tab's queues show one:
    //
    //   - CONTRIBUTOR SUBMISSIONS: the BETA data after approving, against BETA now. The server
    //     builds it (ForApproval) - only it can see the board's folder and tell whether the files
    //     the approval writes from the table would change.
    //   - BETA > PROD: production after publishing, against production now. The maintainer
    //     application builds it (ForPromotion) from the production plan, which already hashed
    //     every file of the board.
    //
    // "The board's files" is its own folder plus everything outside it the board uses - shared
    // files at whatever level they sit - which is what both the approval and the promotion write.
    // On the wire as these records; the enums travel as names.
    // ###########################################################################################
    [JsonConverter(typeof(JsonStringEnumConverter<BoardFileChange>))]
    public enum BoardFileChange
    {
        Unchanged,
        Added,
        Changed,
        Removed
    }

    // Where the bytes a maintainer opens are read from: the BETA or production data (public, as CRT
    // downloads them), the submission's own upload (by hash), or nowhere yet - a file the approval
    // will write that does not exist in any form today.
    [JsonConverter(typeof(JsonStringEnumConverter<BoardFileSource>))]
    public enum BoardFileSource
    {
        Beta,
        Production,
        Submission,
        NotWrittenYet
    }

    // `Sha256` is the submitted bytes when OpenFrom is Submission. `WrittenOnApproval` marks the
    // workbook and the highlight file, which the approval generates from the table rather than
    // copying - so what opens today is BETA's current copy, or nothing for a new one.
    //
    // `SizeBytes` (owner request, 2026-10-04: "Ideally the 'Files' actually also states its size
    // (everywhere)") is the size of what opens: BETA's or production's file, or the submission's
    // upload. Null for a file not written yet, and from a server older than 4.4.0 - the tree then
    // shows no size rather than a wrong one.
    public sealed record BoardFileEntry(
        string Path,
        BoardFileChange Change,
        BoardFileSource OpenFrom,
        string? Sha256 = null,
        bool WrittenOnApproval = false,
        long? SizeBytes = null);

    // A file the approval generates: where, whether BETA has it now, and whether writing it would
    // change it.
    public sealed record GeneratedFile(string Path, bool ExistsInBeta, bool Changes);

    public static class BoardFileEntries
    {
        // ###########################################################################################
        // PRODUCTION AFTER PUBLISHING, from the production plan: its copies (added or replaced),
        // what it removes, and what it found already the same. Everything that exists after the
        // promotion is BETA's bytes, so that is where it opens from; a removed file only exists in
        // production now.
        // ###########################################################################################
        public static IReadOnlyList<BoardFileEntry> ForPromotion(
            IReadOnlyList<PromotionFile>? copies,
            IReadOnlyCollection<string>? removals,
            IReadOnlyCollection<string>? unchanged)
        {
            var entries = new Dictionary<string, BoardFileEntry>(StringComparer.Ordinal);

            foreach (string path in unchanged ?? [])
                BoardFileEntries.Put(entries, new BoardFileEntry(path, BoardFileChange.Unchanged, BoardFileSource.Beta));

            foreach (string path in removals ?? [])
                BoardFileEntries.Put(entries, new BoardFileEntry(path, BoardFileChange.Removed, BoardFileSource.Production));

            foreach (PromotionFile file in copies ?? [])
            {
                BoardFileEntries.Put(entries, new BoardFileEntry(
                    file.Path,
                    file.Change == PromotionChange.Added ? BoardFileChange.Added : BoardFileChange.Changed,
                    BoardFileSource.Beta));
            }

            return BoardFileEntries.Ordered(entries);
        }

        // ###########################################################################################
        // THE BETA DATA AFTER APPROVING, against BETA now.
        //
        //   submitted - every file the submission carries: every file its rows cite, own or shared,
        //               with the hash BETA has at that path now (SubmittedFileFacts);
        //   removals  - what the approval removes (FileRemovalPreview, the list it is sent back);
        //   ownInBeta - every file in the board's own BETA folder now;
        //   workbook, sidecar - the two files the approval writes from the table.
        //
        // Later sources win for a path: a generated file over the submission's entry for it, a
        // submitted file over a removal, anything over "just there". A file in the board's folder
        // that nothing above names is there before and after - unchanged.
        // ###########################################################################################
        public static IReadOnlyList<BoardFileEntry> ForApproval(
            IReadOnlyList<SubmittedFileFact>? submitted,
            IReadOnlyCollection<string>? removals,
            IReadOnlyCollection<string>? ownInBeta,
            GeneratedFile? workbook,
            GeneratedFile? sidecar)
        {
            var entries = new Dictionary<string, BoardFileEntry>(StringComparer.Ordinal);

            foreach (string path in ownInBeta ?? [])
                BoardFileEntries.Put(entries, new BoardFileEntry(path, BoardFileChange.Unchanged, BoardFileSource.Beta));

            foreach (string path in removals ?? [])
                BoardFileEntries.Put(entries, new BoardFileEntry(path, BoardFileChange.Removed, BoardFileSource.Beta));

            foreach (SubmittedFileFact fact in submitted ?? [])
            {
                BoardFileEntry entry = fact.IsUnchanged
                    ? new BoardFileEntry(fact.Path, BoardFileChange.Unchanged, BoardFileSource.Beta)
                    : new BoardFileEntry(
                        fact.Path,
                        fact.ExistsOnServer ? BoardFileChange.Changed : BoardFileChange.Added,
                        BoardFileSource.Submission,
                        fact.Sha256);

                BoardFileEntries.Put(entries, entry);
            }

            foreach (GeneratedFile? generated in new[] { workbook, sidecar })
            {
                if (generated is null)
                    continue;

                BoardFileChange change = !generated.ExistsInBeta
                    ? BoardFileChange.Added
                    : generated.Changes ? BoardFileChange.Changed : BoardFileChange.Unchanged;

                BoardFileEntries.Put(entries, new BoardFileEntry(
                    generated.Path,
                    change,
                    generated.ExistsInBeta ? BoardFileSource.Beta : BoardFileSource.NotWrittenYet,
                    WrittenOnApproval: true));
            }

            return BoardFileEntries.Ordered(entries);
        }

        // ###########################################################################################
        // A BOARD'S FILES AS BETA HOLDS THEM NOW, with nothing changing (owner request, 2026-10-03:
        // the Boards screen's Files view "should not show changed files - just list all files").
        //
        //   own   - every file in the board's own BETA folder;
        //   cited - every file its board cites, wherever it sits - the shared files are the ones
        //           outside the folder.
        //
        // Each once, all opened from BETA's public address - or, for the stable source's files
        // (`source` Production, 2026-10-04), from its own. The same "own folder plus what it uses
        // elsewhere" the two other trees are built from.
        // ###########################################################################################
        public static IReadOnlyList<BoardFileEntry> ForBoard(
            IReadOnlyCollection<string>? own,
            IReadOnlyCollection<string>? cited,
            BoardFileSource source = BoardFileSource.Beta)
        {
            var entries = new Dictionary<string, BoardFileEntry>(StringComparer.Ordinal);

            foreach (string path in (own ?? []).Concat(cited ?? []))
                BoardFileEntries.Put(entries, new BoardFileEntry(path, BoardFileChange.Unchanged, source));

            return BoardFileEntries.Ordered(entries);
        }

        // ###########################################################################################
        // The entries with their sizes (2026-10-04): `sizeOf` answers each entry's size, or null to
        // leave it without one. The builders above say what each file becomes; where its bytes are
        // - and so how big it is - only the server can see (CRT.Server's TreeFileSizes), or the
        // production plan's FileSizes says.
        // ###########################################################################################
        public static IReadOnlyList<BoardFileEntry> WithSizes(
            IReadOnlyList<BoardFileEntry>? entries,
            Func<BoardFileEntry, long?> sizeOf)
        {
            ArgumentNullException.ThrowIfNull(sizeOf);

            return (entries ?? [])
                .Select(entry => sizeOf(entry) is long size && size >= 0 ? entry with { SizeBytes = size } : entry)
                .ToList();
        }

        // ###########################################################################################
        // Whether a generated file is the board's HIGHLIGHT file (the `.json` beside its workbook)
        // rather than the workbook itself (code review, 2026-10-01). The two GeneratedFile arguments
        // of ForApproval already say which is which, but the entry that reaches the tree does not
        // carry it, and the wording used to re-derive it from the extension in a private helper of
        // its own. It is named HERE, beside the code that decides which path is the sidecar, so a
        // change to what the sidecar is called is made in one place - and the tree's sentence is
        // never left guessing from a path.
        // ###########################################################################################
        public static bool IsHighlightFile(string? path) =>
            !string.IsNullOrEmpty(path) && path.EndsWith(".json", StringComparison.OrdinalIgnoreCase);

        private static void Put(Dictionary<string, BoardFileEntry> entries, BoardFileEntry entry)
        {
            string path = (entry.Path ?? string.Empty).Replace('\\', '/').Trim('/');

            if (path.Length > 0)
                entries[path] = entry with { Path = path };
        }

        private static IReadOnlyList<BoardFileEntry> Ordered(Dictionary<string, BoardFileEntry> entries) =>
            entries.Values.OrderBy(entry => entry.Path, StringComparer.Ordinal).ToList();
    }
}
