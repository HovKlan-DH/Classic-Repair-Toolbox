using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHEN A LOCAL DRAFT HAS DONE ITS JOB AND SHOULD GO (owner request, 2026-09-23).
    //
    // *** THE HALF THAT WAS DESCRIBED BUT NEVER BUILT. *** TabDrafts' own header says a draft is
    // left alone on submit and "goes away when the work is actually published and the synced data
    // starts carrying it". The first sentence was implemented; the second never was, so
    // DraftManager.DiscardDraft had exactly one caller - the manual Discard button - and a
    // contributor whose work had shipped was left with a draft folder they had to remember to
    // clean up themselves. Reported after the first real publish.
    //
    // *** THE TEST IS EQUALITY WITH THE PUBLISHED BOARD, NOT THE SERVER'S SAY-SO. *** A state of
    // "published" is necessary but nowhere near sufficient, and treating it as sufficient would
    // be a silent data-loss bug:
    //
    //   - the sync that brings the published board down may not have run yet, so the local
    //     "published" copy can still be the OLD one;
    //   - the contributor may have kept working in the draft after submitting, and those edits
    //     have been published by nobody;
    //   - a partially-applied publish, or a maintainer who edited the submission before merging,
    //     means what shipped is not what was sent.
    //
    // In every one of those cases the draft still holds work that exists nowhere else. So the
    // question asked here is the only one that is actually safe: *is there anything in this draft
    // that the published board does not already say?* If the answer is no, the folder is a
    // duplicate of synced data and deleting it destroys nothing. If anything at all differs, the
    // draft stays and the contributor deals with it.
    //
    // *** A NEW SYSTEM'S DRAFT IS RETIRED LIKE ANY OTHER, ONCE CRT CAN SHOW THE PUBLISHED ONE
    // (owner request, 2026-09-25): "People will either not know they can/should remove this
    // or they forget, so better clean-up when we can." *** It used to be refused outright. What
    // protects it is the published copy: until the system is published, listed in the master
    // workbook and synced, there is no published workbook on this machine, IsRetirable answers no,
    // and the draft - the only copy - stays. CRT downloads a board workbook only once the master
    // lists it, so "the published workbook is here" also means "CRT lists the published board":
    // removing the draft can never make the contributor's board disappear from their CRT.
    //
    // PURE except for reading the two folders - workbooks, sidecars and files - which is what the
    // comparison IS. Nothing here deletes anything - the caller does that, so the decision can be
    // unit tested on its own.
    // ###########################################################################################
    public static class DraftRetirement
    {
        // ###########################################################################################
        // Has the server said this submission is actually LIVE?
        //
        // *** "approved" AND "accepted" DELIBERATELY DO NOT COUNT. *** Both are past the decision
        // and both show the contributor a cheerful green row, but neither means the data tree has
        // been written - an approved submission is still waiting to be published. Retiring a draft
        // at that point would delete the work before anything carried it, and the sync would then
        // bring down a board that still lacks the change.
        //
        // ###########################################################################################
        // *** ONLY PRODUCTION COUNTS - "merged" (BETA) NO LONGER DOES (owner decision,
        // 2026-09-27). *** "It will not remove anything from the user's Draft tab until the data has
        // been migrated to production." A submission in BETA can still be pushed back to the queue
        // by a maintainer, and a draft retired at the BETA stage was then gone from the
        // contributor's machine - the one copy they would correct it from. So a draft stays until
        // its work reaches the production data everyone downloads ("published").
        // ###########################################################################################
        public static bool IsPublishedState(string? state)
        {
            if (string.IsNullOrWhiteSpace(state))
            {
                return false;
            }

            return state.Trim().ToLowerInvariant() switch
            {
                "published" => true,

                // In BETA only - it can still be taken back out (see above).
                "merged" => false,

                // "returned" was merged and has been taken back OUT of BETA (code review,
                // 2026-09-27) - its work is in the published tree no longer, and the draft may be
                // the only copy the contributor has to correct it from.
                "returned" => false,

                _ => false
            };
        }

        // ###########################################################################################
        // Is this draft now a duplicate of the published board, and therefore safe to delete?
        //
        // Returns false for every uncertainty - no draft, no published workbook, an unreadable
        // workbook on either side, or a draft-only system. Failing closed is the only acceptable
        // direction: the cost of keeping a redundant folder is that the contributor sees a stale
        // row, and the cost of deleting a live one is work that exists nowhere else.
        //
        // *** UNCACHED ON BOTH SIDES, deliberately. *** The contributor may have edited the draft
        // workbook in Excel since the application last read it, and a cached comparison could call
        // two boards identical on the strength of a copy taken before those edits - which is
        // exactly the deletion this method exists to prevent.
        //
        // *** ROWS ARE NOT THE WHOLE DRAFT (code review, 2026-09-25). *** A draft folder also holds
        // the KiCad calibrations (in the sidecar, where BoardData has no section for them) and the
        // bytes of every file it references. This used to compare rows alone, so a contributor who
        // re-calibrated a KiCad overlay, or replaced "Sheet1.png" with a corrected scan under the
        // same name, after submitting had that work deleted without being asked the moment the
        // published rows caught up. All three are now compared, and ALL must match.
        // ###########################################################################################
        public static bool IsRetirable(DraftStatus? status)
        {
            if (status is null)
            {
                return false;
            }

            // A NEW system is not refused here any more (2026-09-25): until it is published and
            // synced, the published-workbook check below keeps it, since there is nothing to
            // compare against. See the class header.
            if (string.IsNullOrWhiteSpace(status.WorkbookPath) || !File.Exists(status.WorkbookPath))
            {
                return false;
            }

            if (string.IsNullOrWhiteSpace(status.PublishedWorkbookPath)
                || !File.Exists(status.PublishedWorkbookPath))
            {
                // The publish may be real and the sync simply not have run yet. Keep the draft.
                return false;
            }

            BoardData? draft = BoardDataReader.ReadWorkbookUncached(status.WorkbookPath);
            if (draft is null)
            {
                return false;
            }

            BoardData? published = BoardDataReader.ReadWorkbookUncached(status.PublishedWorkbookPath);
            if (published is null)
            {
                return false;
            }

            // The rows: does the draft say anything the published board does not? Comparing in this
            // direction (published as the base, draft as the proposal) is the same orientation the
            // Drafts tab's own "N rows changed" uses.
            if (BoardDataDiffer.CountChanges(published, draft) != 0)
            {
                return false;
            }

            // The calibrations, which live only in the sidecar - compared through ReviewSummary,
            // the same comparison a maintainer is shown, so "unchanged" means one thing everywhere.
            if (!DraftRetirement.CalibrationsMatch(status, draft, published))
            {
                return false;
            }

            // And the bytes of every file the draft folder holds.
            return DraftRetirement.FilesMatchPublished(status);
        }

        // ###########################################################################################
        // Whether the draft's KiCad calibrations are exactly the published ones.
        // ###########################################################################################
        private static bool CalibrationsMatch(DraftStatus status, BoardData draft, BoardData published)
        {
            IReadOnlyList<KiCadCalibrationEntry> draftCalibrations =
                DraftBoardSource.CollectCalibrations(status.WorkbookPath, draft);

            IReadOnlyList<KiCadCalibrationEntry> publishedCalibrations =
                DraftBoardSource.CollectCalibrations(status.PublishedWorkbookPath, published);

            ReviewChangeSummary summary = ReviewSummary.Compare(
                published,
                draft,
                renames: null,
                publishedCalibrations: publishedCalibrations,
                submittedCalibrations: draftCalibrations);

            return summary.Sections
                .Where(section => string.Equals(section.Section, ReviewSummary.SectionKiCadCalibrations, StringComparison.Ordinal))
                .All(section => !section.HasChanges);
        }

        // ###########################################################################################
        // Whether every file in the draft folder exists in the published folder with the SAME
        // BYTES.
        //
        // A file with no published counterpart - a new image dropped in, an Excel lock file
        // because the workbook is open right now - fails it, which is the safe direction: a lock
        // file means the draft is in use, and a file nobody has published yet is exactly the work
        // this check exists to protect.
        //
        // The workbook and its sidecar are skipped here: their CONTENT has already been compared
        // above, and comparing their bytes would fail on nothing more than a different writer's
        // formatting. The draft marker is local bookkeeping and never published.
        // ###########################################################################################
        private static bool FilesMatchPublished(DraftStatus status)
        {
            string? draftFolder = Path.GetDirectoryName(status.WorkbookPath);
            string? publishedFolder = Path.GetDirectoryName(status.PublishedWorkbookPath);

            if (string.IsNullOrWhiteSpace(draftFolder) || string.IsNullOrWhiteSpace(publishedFolder)
                || !Directory.Exists(draftFolder))
            {
                return false;
            }

            string draftSidecar = BoardComponentHighlightStorage.GetJsonPath(status.WorkbookPath);

            try
            {
                foreach (string draftFile in Directory.EnumerateFiles(draftFolder, "*", SearchOption.AllDirectories))
                {
                    if (DraftRetirement.SamePath(draftFile, status.WorkbookPath)
                        || DraftRetirement.SamePath(draftFile, draftSidecar)
                        || DraftFolderLayout.IsDraftOnlyFile(draftFile))
                    {
                        continue;
                    }

                    string relative = Path.GetRelativePath(draftFolder, draftFile);
                    string publishedFile = DraftRetirement.PublishedCounterpart(status, publishedFolder, relative);

                    if (!DraftRetirement.SameBytes(draftFile, publishedFile))
                    {
                        return false;
                    }
                }

                return true;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // A file that cannot be read might hold anything. Keep the draft.
                CrtLog.Warning($"Could not compare the draft files for [{status.SystemKey}] - [{ex.Message}]");
                return false;
            }
        }

        // ###########################################################################################
        // Where a draft file is published. The board's own files sit in the board's published
        // folder; a SHARED file the contributor attached sits in the draft under its whole path
        // (DraftFileResolver.BuildDraftFileDestination) and is published in that shared folder,
        // from the data root - comparing it with a copy inside the board's folder could never
        // match, and the draft would never be retired (2026-09-25).
        // ###########################################################################################
        private static string PublishedCounterpart(DraftStatus status, string publishedFolder, string relative)
        {
            string[] system = status.SystemKey.Split('/', StringSplitOptions.RemoveEmptyEntries);
            string path = relative.Replace(Path.DirectorySeparatorChar, '/');

            SubmissionFileScope scope = system.Length > 1
                ? SubmissionFileScopes.Classify(system[0], hardware: null, board: null, path)
                : SubmissionFileScope.Foreign;

            if (scope is not (SubmissionFileScope.ManufacturerShared or SubmissionFileScope.GenericShared))
                return Path.Combine(publishedFolder, relative);

            // The data root is the published folder minus the system's own folder segments.
            string dataRoot = publishedFolder;

            for (int i = 0; i < system.Length - 1; i++)
                dataRoot = Path.GetDirectoryName(dataRoot) ?? dataRoot;

            return Path.Combine(dataRoot, relative);
        }

        private static bool SamePath(string a, string b) =>
            string.Equals(Path.GetFullPath(a), Path.GetFullPath(b), StringComparison.OrdinalIgnoreCase);

        private static bool SameBytes(string first, string second)
        {
            var firstInfo = new FileInfo(first);
            var secondInfo = new FileInfo(second);

            if (!secondInfo.Exists || firstInfo.Length != secondInfo.Length)
            {
                return false;
            }

            const int BufferSize = 64 * 1024;

            using FileStream a = new(first, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, BufferSize);
            using FileStream b = new(second, FileMode.Open, FileAccess.Read, FileShare.ReadWrite, BufferSize);

            byte[] bufferA = new byte[BufferSize];
            byte[] bufferB = new byte[BufferSize];

            while (true)
            {
                int readA = a.ReadAtLeast(bufferA, BufferSize, throwOnEndOfStream: false);
                int readB = b.ReadAtLeast(bufferB, BufferSize, throwOnEndOfStream: false);

                if (readA != readB)
                {
                    return false;
                }

                if (readA == 0)
                {
                    return true;
                }

                if (!bufferA.AsSpan(0, readA).SequenceEqual(bufferB.AsSpan(0, readB)))
                {
                    return false;
                }
            }
        }

        // ###########################################################################################
        // A cheap fingerprint of everything in a draft folder: each file's relative path, length
        // and last-write time, hashed together.
        //
        // *** WHY IT EXISTS: THE CHECK AND THE DELETE ARE NOT AT THE SAME MOMENT. *** The
        // comparison parses two workbooks and reads every file, so it runs off the UI thread and
        // takes a while - long enough for the contributor to save an edit into that very draft
        // before the delete runs. Taken BEFORE the comparison and taken again just before the
        // delete, any write in between shows up as a different stamp and the draft is kept.
        // ###########################################################################################
        public static string FolderStamp(string? folder)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
            {
                return string.Empty;
            }

            try
            {
                var builder = new StringBuilder();

                foreach (string file in Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories)
                             .OrderBy(path => path, StringComparer.Ordinal))
                {
                    var info = new FileInfo(file);

                    builder
                        .Append(Path.GetRelativePath(folder, file)).Append('|')
                        .Append(info.Length).Append('|')
                        .Append(info.LastWriteTimeUtc.Ticks).Append('\n');
                }

                return Convert.ToHexString(SHA256.HashData(Encoding.UTF8.GetBytes(builder.ToString())));
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                // Never equal to a real stamp, so an unreadable folder is never deleted.
                return Guid.NewGuid().ToString("N");
            }
        }

        // ###########################################################################################
        // Whether a draft found retirable is still exactly as it was when the check began - the
        // last thing to ask before deleting it. See FolderStamp.
        // ###########################################################################################
        public static bool IsUnchangedSince(RetirableDraft draft)
        {
            ArgumentNullException.ThrowIfNull(draft);

            return draft.Stamp.Length > 0
                && string.Equals(DraftRetirement.FolderStamp(draft.Folder), draft.Stamp, StringComparison.Ordinal);
        }

        // ###########################################################################################
        // The systems whose drafts have shipped and can now be removed.
        //
        // Takes the receipts so the caller does not have to pair them with drafts itself, and
        // returns system keys rather than performing the deletion, so the rule stays testable with
        // no filesystem writes.
        //
        // A system is named at most ONCE even when several of its submissions are published, which
        // is ordinary: a contributor who submits, gets published, edits again and submits again
        // has two published receipts for one system.
        // ###########################################################################################
        public static IReadOnlyList<string> FindRetirableSystems(
            IEnumerable<SubmissionReceipt>? receipts,
            Func<string, DraftStatus?> resolveStatus) =>
            DraftRetirement.FindRetirableDrafts(receipts, resolveStatus)
                .Select(draft => draft.ExcelDataFile)
                .ToList();

        // ###########################################################################################
        // THE SEARCH THE APPLICATION RUNS: receipts name their system by id, and each is matched
        // to its board's workbook among `excelDataFiles` (the boards the app knows) - see
        // DraftStatusReader.ResolveForSystem. The caller passes values, so no key translation is
        // left at the call site to get wrong; that translation being wrong is why no draft was ever
        // retired until 2026-09-25.
        // ###########################################################################################
        public static IReadOnlyList<RetirableDraft> FindRetirableDrafts(
            IEnumerable<SubmissionReceipt>? receipts,
            string dataRoot,
            string draftsRoot,
            IEnumerable<string>? excelDataFiles)
        {
            List<string> known = (excelDataFiles ?? []).ToList();

            return DraftRetirement.FindRetirableDrafts(
                receipts,
                systemId => DraftStatusReader.ResolveForSystem(dataRoot, draftsRoot, systemId, known));
        }

        // ###########################################################################################
        // The same search, carrying what the caller needs to delete SAFELY: the folder, and its
        // stamp taken BEFORE the comparison ran - see FolderStamp and IsUnchangedSince.
        // ###########################################################################################
        public static IReadOnlyList<RetirableDraft> FindRetirableDrafts(
            IEnumerable<SubmissionReceipt>? receipts,
            Func<string, DraftStatus?> resolveStatus)
        {
            ArgumentNullException.ThrowIfNull(resolveStatus);

            var retirable = new List<RetirableDraft>();

            if (receipts is null)
            {
                return retirable;
            }

            var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (SubmissionReceipt receipt in receipts)
            {
                if (receipt is null || string.IsNullOrWhiteSpace(receipt.SystemId))
                {
                    continue;
                }

                if (!DraftRetirement.IsPublishedState(receipt.LastKnownState))
                {
                    continue;
                }

                if (!seen.Add(receipt.SystemId))
                {
                    continue;
                }

                DraftStatus? status = resolveStatus(receipt.SystemId);
                string folder = Path.GetDirectoryName(status?.WorkbookPath ?? string.Empty) ?? string.Empty;

                // BEFORE the comparison, so an edit made while it runs changes the stamp.
                string stamp = DraftRetirement.FolderStamp(folder);

                // Carries the draft's WORKBOOK key, not the receipt's system id: everything that
                // follows - the unsaved-edits guard, the delete, the cache clear - is keyed by it.
                if (DraftRetirement.IsRetirable(status))
                {
                    retirable.Add(new RetirableDraft(status!.SystemKey, folder, stamp));
                }
            }

            return retirable;
        }
    }

    // ###########################################################################################
    // One draft found safe to retire: its board's workbook path ("Commodore/C64/250407/Data C64
    // 250407 v2.0.0.xlsx" - the key DraftManager, the table editor and the board cache all use), its
    // folder, and the folder's stamp from before the comparison. The caller deletes it only if
    // IsUnchangedSince still holds.
    // ###########################################################################################
    public sealed record RetirableDraft(string ExcelDataFile, string Folder, string Stamp);
}
