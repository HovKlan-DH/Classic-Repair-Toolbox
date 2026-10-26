using Handlers.DataHandling;

namespace CRT.Server.Handlers.Submissions
{
    // ###########################################################################################
    // A PUBLISHED BOARD'S FILES IN THE CRT 4.0 TREE (2026-10-25) - the files its board file lists,
    // by the board's own names, and where each one's bytes are stored.
    //
    // A board names every file it has ("Component images/7406.png") with the file's SHA-256 and
    // size (BoardDocument.Files); the bytes lie once under their checksum,
    // "Files/<ab>/<sha256>.<ext>" (DataTreeLayout.StoredFilePath). So "what is published under this
    // name, and is it the same file?" is a look in the board's list - no file is hashed to answer
    // it - and the bytes are found from the checksum.
    //
    // An I/O boundary, thin on purpose: the board file is read through BoardFiles (CRT.Data's one
    // reader of it), and every place in the tree is resolved through SubmissionPathRules, the one
    // containment rule this service has - the board id and the names come from the database and
    // from what a submission sent.
    // ###########################################################################################
    public static class PublishedBoardFiles
    {
        // ###########################################################################################
        // The board's files by name - compared without regard to capitals, as a board's names are
        // (BoardFileNames). EMPTY for a board the tree does not hold (a new board has no files yet);
        // NULL when its board file is there and cannot be read - never an empty list, which would
        // read as "the board has no files" and turn every file into an addition.
        // ###########################################################################################
        public static IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? ByName(string? dataTreeRoot, string? boardId)
        {
            PublishedBoardLocation location = PublishedBoardLocator.LocateBoard(dataTreeRoot, boardId);

            if (!location.Exists)
                return new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

            return PublishedBoardFiles.ByName(BoardFiles.ReadDocument(location.WorkbookPath));
        }

        // The same, from a board file already read; null for none.
        public static IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? ByName(BoardDocument? document)
        {
            if (document is null)
                return null;

            var files = new Dictionary<string, BoardDocument.StoredFileRow>(StringComparer.OrdinalIgnoreCase);

            foreach (BoardDocument.StoredFileRow file in document.Files)
            {
                if (!string.IsNullOrWhiteSpace(file.Name))
                    files.TryAdd(file.Name.Trim(), file);
            }

            return files;
        }

        // "Files/3f/3f7a...e1.png" for a listed file, or "" when its checksum is not a usable one.
        public static string StoredRelative(BoardDocument.StoredFileRow file)
        {
            ArgumentNullException.ThrowIfNull(file);

            string sha = file.Sha256?.Trim().ToLowerInvariant() ?? string.Empty;

            return DataTreeLayout.IsSha256(sha) ? DataTreeLayout.StoredFilePath(sha, file.Name) : string.Empty;
        }

        // ###########################################################################################
        // Where the tree at `dataTreeRoot` keeps the bytes of a file - the full path, contained - or
        // "" when it cannot say. Whether the file is THERE is the caller's question (IsStored).
        // ###########################################################################################
        public static string StoredPath(string? dataTreeRoot, string? sha256, string? name)
        {
            string sha = sha256?.Trim().ToLowerInvariant() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(dataTreeRoot) || !DataTreeLayout.IsSha256(sha))
                return string.Empty;

            return SubmissionPathRules.TryResolve(dataTreeRoot, DataTreeLayout.StoredFilePath(sha, name ?? string.Empty), out string path, out _)
                ? path
                : string.Empty;
        }

        // Whether the tree holds the stored file at this data-root-relative path.
        public static bool IsStored(string? dataTreeRoot, string? storedRelative)
        {
            if (string.IsNullOrWhiteSpace(dataTreeRoot) || string.IsNullOrWhiteSpace(storedRelative))
                return false;

            return SubmissionPathRules.TryResolve(dataTreeRoot, storedRelative, out string path, out _) && File.Exists(path);
        }

        // ###########################################################################################
        // Where the tree holds bytes with this checksum, whatever name they were stored under, or
        // "" when it does not. The stored file of `name` first ("x.png" stores as ".png"); then any
        // other with the checksum - the same bytes stored for a name of another extension.
        // ###########################################################################################
        public static string FindStored(string? dataTreeRoot, string? sha256, string? name)
        {
            string exact = PublishedBoardFiles.StoredPath(dataTreeRoot, sha256, name);

            if (exact.Length == 0)
                return string.Empty;

            if (File.Exists(exact))
                return exact;

            string? folder = Path.GetDirectoryName(exact);
            string sha = sha256!.Trim().ToLowerInvariant();

            try
            {
                if (folder is null || !Directory.Exists(folder))
                    return string.Empty;

                return Directory
                    .EnumerateFiles(folder, sha + "*")
                    .FirstOrDefault(path =>
                    {
                        string file = Path.GetFileName(path);
                        return file.Length == sha.Length || (file.Length > sha.Length && file[sha.Length] == '.');
                    }) ?? string.Empty;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                return string.Empty;
            }
        }
    }
}
