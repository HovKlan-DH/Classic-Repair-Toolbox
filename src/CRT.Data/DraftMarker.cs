using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

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
    //   - NewSystem: the registration for a system that exists ONLY as a draft. There is no
    //     published counterpart to derive a manufacturer/hardware/board identity from.
    //
    // *** NEVER SUBMITTED, NEVER PUBLISHED. *** See DraftFolderLayout.DraftMarkerFileName for the
    // two independent mechanisms that keep it out of a submission.
    // ###########################################################################################
    public sealed class DraftMarker
    {
        // The system's ExcelDataFile identity, carried so a marker found on disk can be checked
        // against the folder it was found in - a folder copied by hand to a new name would
        // otherwise claim to be a draft of the system it was copied FROM.
        public string SystemKey { get; init; } = string.Empty;

        // ###########################################################################################
        // The published RevisionDate this draft was seeded from.
        //
        // EMPTY for a draft-only system, deliberately: it has no official counterpart, and stamping
        // today's date here would make a later drift check compare against a revision that never
        // existed. The same reasoning BoardDraft.BaseRevision carried.
        // ###########################################################################################
        public string BaseRevision { get; init; } = string.Empty;

        // Set only on a system that exists purely as a draft. Null for the ordinary case - a draft
        // seeded from a published board, whose identity is already established.
        public NewSystemRegistration? NewSystem { get; init; }

        // When the draft was created, so a contributor with several can tell which is which. UTC,
        // round-trip format, for the same reason receipts use it: a file that travels between time
        // zones must not reorder itself.
        public string CreatedUtc { get; init; } = string.Empty;

        // True when this draft has no published counterpart at all.
        [JsonIgnore]
        public bool IsNewSystem => this.NewSystem != null;
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
        public static void Save(string markerPath, DraftMarker marker)
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

            File.Move(temporary, markerPath, overwrite: true);
        }
    }
}
