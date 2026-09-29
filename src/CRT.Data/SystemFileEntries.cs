using System;
using System.Collections.Generic;
using System.Linq;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // A SYSTEM'S FILES, AND WHAT THE NEXT STEP DOES TO EACH (owner request, 2026-09-28): "it would
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
    //     every file of the system.
    //
    // "The system's files" is its own folder plus everything outside it the board uses - shared
    // files at whatever level they sit - which is what both the approval and the promotion write.
    // On the wire as these records; the enums travel as names.
    // ###########################################################################################
    [JsonConverter(typeof(JsonStringEnumConverter<SystemFileChange>))]
    public enum SystemFileChange
    {
        Unchanged,
        Added,
        Changed,
        Removed
    }

    // Where the bytes a maintainer opens are read from: the BETA or production data (public, as CRT
    // downloads them), the submission's own upload (by hash), or nowhere yet - a file the approval
    // will write that does not exist in any form today.
    [JsonConverter(typeof(JsonStringEnumConverter<SystemFileSource>))]
    public enum SystemFileSource
    {
        Beta,
        Production,
        Submission,
        NotWrittenYet
    }

    // `Sha256` is the submitted bytes when OpenFrom is Submission. `WrittenOnApproval` marks the
    // workbook and the highlight file, which the approval generates from the table rather than
    // copying - so what opens today is BETA's current copy, or nothing for a new one.
    public sealed record SystemFileEntry(
        string Path,
        SystemFileChange Change,
        SystemFileSource OpenFrom,
        string? Sha256 = null,
        bool WrittenOnApproval = false);

    // A file the approval generates: where, whether BETA has it now, and whether writing it would
    // change it.
    public sealed record GeneratedFile(string Path, bool ExistsInBeta, bool Changes);

    public static class SystemFileEntries
    {
        // ###########################################################################################
        // PRODUCTION AFTER PUBLISHING, from the production plan: its copies (added or replaced),
        // what it removes, and what it found already the same. Everything that exists after the
        // promotion is BETA's bytes, so that is where it opens from; a removed file only exists in
        // production now.
        // ###########################################################################################
        public static IReadOnlyList<SystemFileEntry> ForPromotion(
            IReadOnlyList<PromotionFile>? copies,
            IReadOnlyCollection<string>? removals,
            IReadOnlyCollection<string>? unchanged)
        {
            var entries = new Dictionary<string, SystemFileEntry>(StringComparer.Ordinal);

            foreach (string path in unchanged ?? [])
                SystemFileEntries.Put(entries, new SystemFileEntry(path, SystemFileChange.Unchanged, SystemFileSource.Beta));

            foreach (string path in removals ?? [])
                SystemFileEntries.Put(entries, new SystemFileEntry(path, SystemFileChange.Removed, SystemFileSource.Production));

            foreach (PromotionFile file in copies ?? [])
            {
                SystemFileEntries.Put(entries, new SystemFileEntry(
                    file.Path,
                    file.Change == PromotionChange.Added ? SystemFileChange.Added : SystemFileChange.Changed,
                    SystemFileSource.Beta));
            }

            return SystemFileEntries.Ordered(entries);
        }

        // ###########################################################################################
        // THE BETA DATA AFTER APPROVING, against BETA now.
        //
        //   submitted - every file the submission carries: every file its rows cite, own or shared,
        //               with the hash BETA has at that path now (SubmittedFileFacts);
        //   removals  - what the approval removes (FileRemovalPreview, the list it is sent back);
        //   ownInBeta - every file in the system's own BETA folder now;
        //   workbook, sidecar - the two files the approval writes from the table.
        //
        // Later sources win for a path: a generated file over the submission's entry for it, a
        // submitted file over a removal, anything over "just there". A file in the board's folder
        // that nothing above names is there before and after - unchanged.
        // ###########################################################################################
        public static IReadOnlyList<SystemFileEntry> ForApproval(
            IReadOnlyList<SubmittedFileFact>? submitted,
            IReadOnlyCollection<string>? removals,
            IReadOnlyCollection<string>? ownInBeta,
            GeneratedFile? workbook,
            GeneratedFile? sidecar)
        {
            var entries = new Dictionary<string, SystemFileEntry>(StringComparer.Ordinal);

            foreach (string path in ownInBeta ?? [])
                SystemFileEntries.Put(entries, new SystemFileEntry(path, SystemFileChange.Unchanged, SystemFileSource.Beta));

            foreach (string path in removals ?? [])
                SystemFileEntries.Put(entries, new SystemFileEntry(path, SystemFileChange.Removed, SystemFileSource.Beta));

            foreach (SubmittedFileFact fact in submitted ?? [])
            {
                SystemFileEntry entry = fact.IsUnchanged
                    ? new SystemFileEntry(fact.Path, SystemFileChange.Unchanged, SystemFileSource.Beta)
                    : new SystemFileEntry(
                        fact.Path,
                        fact.ExistsOnServer ? SystemFileChange.Changed : SystemFileChange.Added,
                        SystemFileSource.Submission,
                        fact.Sha256);

                SystemFileEntries.Put(entries, entry);
            }

            foreach (GeneratedFile? generated in new[] { workbook, sidecar })
            {
                if (generated is null)
                    continue;

                SystemFileChange change = !generated.ExistsInBeta
                    ? SystemFileChange.Added
                    : generated.Changes ? SystemFileChange.Changed : SystemFileChange.Unchanged;

                SystemFileEntries.Put(entries, new SystemFileEntry(
                    generated.Path,
                    change,
                    generated.ExistsInBeta ? SystemFileSource.Beta : SystemFileSource.NotWrittenYet,
                    WrittenOnApproval: true));
            }

            return SystemFileEntries.Ordered(entries);
        }

        private static void Put(Dictionary<string, SystemFileEntry> entries, SystemFileEntry entry)
        {
            string path = (entry.Path ?? string.Empty).Replace('\\', '/').Trim('/');

            if (path.Length > 0)
                entries[path] = entry with { Path = path };
        }

        private static IReadOnlyList<SystemFileEntry> Ordered(Dictionary<string, SystemFileEntry> entries) =>
            entries.Values.OrderBy(entry => entry.Path, StringComparer.Ordinal).ToList();
    }
}
