using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A PUBLISHED BOARD'S FILE, READ ONCE (code review, 2026-10-26) - for an approval, which parsed
    // the same multi-megabyte board file for its list of files three times (the plan's view of what
    // is published, the removal preview, the change summary), for its renames once and for its
    // calibrations once per schematic - all under the publish lock every other publish waits on -
    // and then once more to write it. Read here once, at the start of the publish, and handed to
    // each of them and to PublishExecutor.
    //
    // The answers are the separate reads': `Files` is PublishedBoardFiles.ByName's (EMPTY for a
    // board the tree does not hold, NULL for one whose file cannot be read - never an empty list,
    // which would turn every file into an addition), `Renames` LoadSchematicRenames' and
    // CalibrationsFor LoadKiCadCalibrations'. `Document` is null for either case, and a board file
    // that cannot be read is read again IN FULL by the executor, which then stops the publish.
    // ###########################################################################################
    public sealed class PublishedBoardRead
    {
        private static readonly IReadOnlyDictionary<string, BoardDocument.StoredFileRow> NoFiles =
            new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

        private readonly BoardContent? thisContent;

        public PublishedBoardRead(string path, bool exists, BoardDocument? document)
        {
            this.Path = path ?? string.Empty;
            this.Exists = exists;
            this.Document = exists ? document : null;
            this.Files = exists ? PublishedBoardFiles.ByName(this.Document) : PublishedBoardRead.NoFiles;
            this.thisContent = this.Document is null ? null : BoardDocumentMapping.ToContent(this.Document);
        }

        // The board `boardId` of the tree at `dataTreeRoot`, its board file read.
        public static PublishedBoardRead Of(string? dataTreeRoot, string? boardId)
        {
            PublishedBoardLocation location = PublishedBoardLocator.LocateBoard(dataTreeRoot, boardId);

            return location.Exists
                ? new PublishedBoardRead(location.WorkbookPath, exists: true, BoardFiles.ReadDocument(location.WorkbookPath))
                : new PublishedBoardRead(location.WorkbookPath, exists: false, null);
        }

        // The board file's path ("" for a board the tree does not hold).
        public string Path { get; }

        public bool Exists { get; }

        public BoardDocument? Document { get; }

        public IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? Files { get; }

        // The board's "Schematic renames", the entries naming both names - as LoadSchematicRenames.
        public IReadOnlyList<SchematicRenameRecord> Renames =>
            this.thisContent?.Renames.Where(rename => rename.From.Trim().Length > 0 && rename.To.Trim().Length > 0).ToList() ?? [];

        // The calibrations filed for these schematic names - as LoadKiCadCalibrations.
        public List<KiCadCalibrationEntry> CalibrationsFor(IEnumerable<string?> schematicNames) =>
            this.thisContent is null ? [] : BoardComponentHighlightStorage.CalibrationsNamed(this.thisContent.Calibrations, schematicNames);

        // Whether this is the board file at `path` - what the executor writes.
        public bool IsAt(string path) =>
            this.Path.Length > 0 &&
            string.Equals(System.IO.Path.GetFullPath(this.Path), System.IO.Path.GetFullPath(path), StringComparison.Ordinal);
    }
}
