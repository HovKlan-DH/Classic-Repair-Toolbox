using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Threading;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // THE LOCAL-ONLY RECORD THAT MAKES A BOARD FOLDER A DRAFT
    // (NewContributeStrategy.md Phase 6 - owner request, 2026-09-23).
    //
    // A draft folder is otherwise indistinguishable from a published board: same workbook, same
    // sidecar, same images at the same relative paths. This file is the one thing that says
    // "these are unpublished local edits", and it holds only what CANNOT be derived from the
    // folder's own contents.
    //
    // *** IT DELIBERATELY CARRIES NO ROW DIFF. *** The obvious design is a ledger of what has
    // changed, and it is wrong here: the contributor can edit the workbook in Excel with the
    // application closed, so any ledger drifts out of date the moment they do. The difference is
    // DERIVED instead, by BoardDataDiffer comparing the draft workbook against the published one.
    // That is what makes editing in the app and editing in Excel genuinely interchangeable rather
    // than merely both possible - neither route has to remember to update anything.
    //
    // So what is left is the two facts a comparison cannot recover:
    //
    //   - BaseRevision: WHICH published revision this draft was taken from. Nothing in the draft
    //     folder records that, and without it the drift warning cannot tell "the published board
    //     moved on underneath you" from "you edited these rows yourself".
    //   - NewBoard: the registration for a board that exists ONLY as a draft. There is no
    //     published counterpart to derive a manufacturer/hardware/board identity from.
    //
    // *** NEVER SUBMITTED, NEVER PUBLISHED. *** See DraftFolderLayout.DraftMarkerFileName for the
    // two independent mechanisms that keep it out of a submission.
    // ###########################################################################################
    public sealed record DraftMarker
    {
        // The board's ExcelDataFile identity, carried so a marker found on disk can be checked
        // against the folder it was found in - a folder copied by hand to a new name would
        // otherwise claim to be a draft of the board it was copied FROM.
        //
        // Written as "SystemKey", the name it had before "system" became "board" everywhere
        // (owner decision, 2026-10-09): every draft folder already on a contributor's disk carries
        // that name, and a marker that no longer reads would turn the draft into an unknown folder.
        [JsonPropertyName("SystemKey")]
        public string BoardKey { get; init; } = string.Empty;

        // ###########################################################################################
        // The published RevisionDate this draft was seeded from.
        //
        // EMPTY for a draft-only board, deliberately: it has no official counterpart, and stamping
        // today's date here would make a later drift check compare against a revision that never
        // existed. The same reasoning BoardDraft.BaseRevision carried.
        // ###########################################################################################
        public string BaseRevision { get; init; } = string.Empty;

        // Set only on a board that exists purely as a draft. Null for the ordinary case - a draft
        // seeded from a published board, whose identity is already established. Written as
        // "NewSystem" for the reason BoardKey keeps "SystemKey".
        [JsonPropertyName("NewSystem")]
        public NewBoardRegistration? NewBoard { get; init; }

        // When the draft was created, so a contributor with several can tell which is which. UTC,
        // round-trip format, for the same reason receipts use it: a file that travels between time
        // zones must not reorder itself.
        public string CreatedUtc { get; init; } = string.Empty;

        // ###########################################################################################
        // THE REVISION THE CONTRIBUTOR SAID "I HAVE LOOKED AT THIS" TO (code review, 2026-10-13),
        // for a draft with no base to bring up to date (one started before CRT kept a base). It used
        // to OVERWRITE BaseRevision - which the submission tells the server, so the maintainer was
        // told a stale draft was made from the newest data. BaseRevision now always says what the
        // draft was really copied from; this only quiets the warning.
        // ###########################################################################################
        public string DriftSeenRevision { get; init; } = string.Empty;

        // ###########################################################################################
        // THE PUBLISHED CONTENT "I HAVE LOOKED AT THIS" WAS SAID TO (code review, 2026-10-13) - the
        // content hash (BaseContentHash's kind) of the published files at that moment. A draft that
        // keeps a content hash is judged by CONTENT (DraftDriftCheck), which never read
        // DriftSeenRevision - so for a draft that could not be brought up to date after all (its
        // base unreadable, the published board unreadable), the warning could not be dismissed.
        // Empty when nothing was seen; cleared when the draft is brought up to date.
        // ###########################################################################################
        public string DriftSeenContentHash { get; init; } = string.Empty;

        // ###########################################################################################
        // WHAT THE PUBLISHED FILES HELD WHEN THE DRAFT WAS COPIED OR LAST BROUGHT UP TO DATE (code
        // review, 2026-10-13): one SHA-256 over the published workbook's and sidecar's bytes
        // (DraftBaseStore.PublishedContentHash). A revision is a DATE, so two publishes on one day
        // read as no change; the content cannot. Empty for a draft from before it existed - those
        // fall back to comparing revisions.
        // ###########################################################################################
        public string BaseContentHash { get; init; } = string.Empty;

        // ###########################################################################################
        // THE CREDITS OFFERED TO A NEW BOARD SO FAR (owner request, 2026-11-02; NewBoardCredits.Offer),
        // each as NewBoardCredits.KeyOf says it - so one the contributor deleted stays deleted, and this
        // survives CRT being closed. Nothing a board publishes: the marker never leaves the machine.
        // ###########################################################################################
        public IReadOnlyList<string>? CreditsOffered { get; init; }

        // ###########################################################################################
        // THE CONTRIBUTOR WANTS NO CREDIT FOR THIS DRAFT'S WORK (owner request, 2026-11-24: "I want a possibility for the user
        // to opt-out of the credit ... when I contribute something, then I do not want to have a new line generated") - the
        // credit CRT added removed, so none is added again (SubmissionCredit.NoCreditAfter, DraftSession.NoCreditForWork). Null
        // unless so, so every other marker keeps its bytes. Nothing a board publishes: the marker never leaves the machine.
        // ###########################################################################################
        public bool? NoCreditForWork { get; init; }

        // True when this draft has no published counterpart at all.
        [JsonIgnore]
        public bool IsNewBoard => this.NewBoard != null;
    }

    // ###########################################################################################
    // Reads and writes the marker. Separate from the record itself so the record stays a plain
    // DTO with no I/O, matching how DraftDataStore sat beside BoardDraft.
    // ###########################################################################################
    public static class DraftMarkerStore
    {
        private static readonly JsonSerializerOptions WriteOptions = new()
        {
            WriteIndented = true,
            DefaultIgnoreCondition = JsonIgnoreCondition.WhenWritingNull,
        };

        private static readonly JsonSerializerOptions ReadOptions = new()
        {
            PropertyNameCaseInsensitive = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
            AllowTrailingCommas = true,
        };

        // ###########################################################################################
        // Loads the marker from a draft folder, or null when there is none.
        //
        // *** A MISSING MARKER MEANS "NOT A DRAFT", which is a legitimate answer rather than an
        // error. *** Callers walk the drafts tree and ask this of every folder they find, so
        // "there is nothing here" is the ordinary case and must not throw.
        //
        // An UNREADABLE marker - truncated by a crash, hand-edited into invalid JSON - also answers
        // null, with a warning. That is the safe direction: treating a folder as not-a-draft leaves
        // its files untouched on disk for the contributor to recover, whereas throwing would take
        // down whatever was enumerating the tree.
        // ###########################################################################################
        public static DraftMarker? Load(string markerPath)
        {
            if (string.IsNullOrWhiteSpace(markerPath) || !File.Exists(markerPath))
            {
                return null;
            }

            try
            {
                string json = File.ReadAllText(markerPath);

                return string.IsNullOrWhiteSpace(json)
                    ? null
                    : JsonSerializer.Deserialize<DraftMarker>(json, DraftMarkerStore.ReadOptions);
            }
            catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not read draft marker [{markerPath}] - [{ex.Message}]");

                return null;
            }
        }

        // ###########################################################################################
        // Writes the marker, creating the folder if needed.
        //
        // ATOMIC: written to a temporary file and moved into place, so a crash mid-write leaves the
        // previous marker rather than a truncated one. A half-written marker would make a real
        // draft look like an ordinary folder, and the contributor's edits would stop being
        // recognised as theirs - the same reasoning AtomicJsonFile applies to the app's own state.
        // ###########################################################################################
        public static void Save(string markerPath, DraftMarker marker) =>
            DraftMarkerStore.Save(markerPath, marker, Thread.Sleep);

        // ###########################################################################################
        // READS, CHANGES AND WRITES A MARKER AS ONE STEP (code review, 2026-11-24). What changes one field of a draft's marker
        // while the draft is in use - the wish for no credit (DraftSession, with a save, on the pool), "I have looked at this"
        // (DraftWorkbookStore.MarkDriftSeen), the new board's credits offered (TabDrafts) - goes through here, one at a time
        // per marker: two at once each read the marker before the other wrote it, and the later write dropped the other's
        // field. What writes a marker whole as it makes, moves or converts its folder (DraftSeeder, DraftBoardRename,
        // DraftConversion, DraftFolderImport) does not - a field one of those carries over while the draft is in use goes
        // through here first. The changed marker, or null when there is none (its write's IOException is the caller's).
        // ###########################################################################################
        public static DraftMarker? Update(string markerPath, Func<DraftMarker, DraftMarker> change)
        {
            ArgumentNullException.ThrowIfNull(change);

            // One gate per marker, so a marker held open for a moment (Save waits for it) holds up no other draft's.
            object gate = DraftMarkerStore.UpdateGates.GetOrAdd(
                string.IsNullOrWhiteSpace(markerPath) ? string.Empty : Path.GetFullPath(markerPath),
                _ => new object());

            lock (gate)
            {
                if (DraftMarkerStore.Load(markerPath) is not { } marker)
                {
                    return null;
                }

                DraftMarker changed = change(marker);
                DraftMarkerStore.Save(markerPath, changed);

                return changed;
            }
        }

        private static readonly ConcurrentDictionary<string, object> UpdateGates = new(StringComparer.OrdinalIgnoreCase);

        internal static void Save(string markerPath, DraftMarker marker, Action<TimeSpan> pause)
        {
            if (string.IsNullOrWhiteSpace(markerPath))
            {
                throw new ArgumentException("A marker path is required.", nameof(markerPath));
            }

            ArgumentNullException.ThrowIfNull(marker);

            string? folder = Path.GetDirectoryName(markerPath);
            if (!string.IsNullOrWhiteSpace(folder))
            {
                Directory.CreateDirectory(folder);
            }

            string temporary = markerPath + ".tmp";

            File.WriteAllText(
                temporary,
                JsonSerializer.Serialize(marker, DraftMarkerStore.WriteOptions));

            // A marker held for a moment is waited for (2026-11-02: a rename of a new board was refused about
            // once in 64 runs) - a virus scanner opens what was just written or moved.
            FileReplacer.MoveOver(temporary, markerPath, (from, to) => File.Move(from, to, overwrite: true), pause);
        }
    }
}
