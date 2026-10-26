using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text.Json;

namespace Handlers.DataHandling
{
    // One file the draft copied from the published board, as it was then. Path is inside the board
    // folder, with forward slashes ("Component images/U8.png").
    public sealed class DraftBaseFile
    {
        public string Path { get; init; } = string.Empty;

        public string Sha256 { get; init; } = string.Empty;

        public long Length { get; init; }

        // The DRAFT copy's write time (UTC ticks) when it was known to hold exactly these bytes - so
        // a copy nobody touched is recognised by its length and write time, without reading it.
        // 0 when not known; the bytes are then hashed.
        public long DraftWriteUtcTicks { get; init; }
    }

    // ###########################################################################################
    // THE PUBLISHED BOARD AS IT WAS COPIED INTO A DRAFT (code review, 2026-10-13). Kept in the draft
    // folder as DraftFolderLayout.DraftBaseFileName, written when the draft is seeded and again each
    // time it is brought up to date (DraftRebaser).
    //
    // A draft is a complete copy of the board, and without a copy of what it was copied FROM nothing
    // can tell the contributor's changes from the changes published since - which is why a draft
    // started before somebody else's work was published used to send that work back as removals.
    // With this, DraftRebase can carry the contributor's changes onto the newer board, and the
    // Drafts tab can tell "a file the contributor replaced" from "a file the update replaced".
    //
    // JSON rather than a second workbook: it is read without parsing Excel, holds the calibrations
    // (which no sheet does) and the file hashes, and nobody is meant to open it.
    // ###########################################################################################
    public sealed class DraftBase
    {
        // Raised if the shape ever changes; a base of an unknown format is not used.
        public const int CurrentFormat = 1;

        public int Format { get; init; } = DraftBase.CurrentFormat;

        public string Revision { get; init; } = string.Empty;

        public BoardData Board { get; init; } = new();

        public List<KiCadCalibrationEntry> Calibrations { get; init; } = [];

        public List<DraftBaseFile> Files { get; init; } = [];
    }

    public static class DraftBaseStore
    {
        private static readonly JsonSerializerOptions Json = new()
        {
            PropertyNameCaseInsensitive = true,
            AllowTrailingCommas = true,
            ReadCommentHandling = JsonCommentHandling.Skip,
        };

        // Content hashes of published files, remembered by the file's length and write time - the
        // drift check asks for one per drafted board on every refresh of the Drafts tab.
        private static readonly ConcurrentDictionary<string, (string Stamp, string Hash)> ContentHashes =
            new(StringComparer.OrdinalIgnoreCase);

        // The draft's base, or null when it has none (a draft started before bases existed, a new
        // board, or a file that cannot be read - every caller then falls back to the old behaviour).
        public static DraftBase? Load(string draftsRoot, string excelDataFile) =>
            DraftBaseStore.LoadFile(DraftFolderLayout.GetBasePath(draftsRoot, excelDataFile));

        // The base in the file at `path` - Load's rule, for a base not in a draft's own place (the
        // CRT 4.0 draft conversion reads it from the folder the Excel-era draft was set aside in).
        public static DraftBase? LoadFile(string path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return null;

            try
            {
                string json = File.ReadAllText(path);

                // A CRT 4.0 draft keeps its base as a board file (SaveDocument).
                if (DraftBaseStore.IsBoardDocument(json))
                    return DraftBaseStore.FromDocument(DataTreeJson.ReadBoard(json));

                DraftBase? loaded = JsonSerializer.Deserialize<DraftBase>(json, DraftBaseStore.Json);
                return loaded is { Format: DraftBase.CurrentFormat } ? loaded : null;
            }
            catch (InvalidDataException ex)
            {
                CrtLog.Warning($"Could not read the draft base [{path}] - [{ex.Message}]; set aside");
                DraftBaseStore.SetAside(path);
                return null;
            }
            catch (Exception ex) when (ex is JsonException or NotSupportedException)
            {
                // *** SET ASIDE, SO THE DRAFT READS AS HAVING NO BASE (code review, 2026-10-13). ***
                // Left in place, the Drafts row (DraftStatus.HasBase - the file exists) said "Bring
                // your draft up to date" while the window, finding no base it could use, offered
                // only "I have looked at this". Kept beside, as bookkeeping (".crt-"), for whoever
                // wants to see what went wrong.
                CrtLog.Warning($"Could not read the draft base [{path}] - [{ex.Message}]; set aside");
                DraftBaseStore.SetAside(path);
                return null;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not read the draft base [{path}] - [{ex.Message}]");
                return null;
            }
        }

        // The unreadable base moved to "<name>.unreadable" (replacing an older one), or left where it
        // is when even that fails - it is then merely read again next time.
        private static void SetAside(string path)
        {
            try
            {
                File.Move(path, path + ".unreadable", overwrite: true);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not set aside the unreadable draft base [{path}] - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // *** A CRT 4.0 DRAFT'S BASE IS THE PUBLISHED BOARD FILE AS IT WAS COPIED (2026-10-24). *** Kept
        // as a board file (DataTreeJson's bytes), read back as a DraftBase for the code that merges and
        // compares against it: its rows, calibrations, and each file by name with its checksum. The
        // checksum is given in upper-case hex there, the form DraftFingerprint compares files by.
        // ###########################################################################################
        public static void SaveDocument(string draftsRoot, string excelDataFile, BoardDocument published)
        {
            ArgumentNullException.ThrowIfNull(published);

            string path = DraftFolderLayout.GetBasePath(draftsRoot, excelDataFile);

            if (path.Length == 0)
                throw new IOException($"No draft folder for [{excelDataFile}].");

            BoardFiles.Write(path, published);
        }

        // The base as the board file it was kept as, or null (no base, an Excel-era one, or unreadable).
        public static BoardDocument? LoadDocument(string draftsRoot, string excelDataFile)
        {
            string path = DraftFolderLayout.GetBasePath(draftsRoot, excelDataFile);

            if (path.Length == 0 || !File.Exists(path))
                return null;

            try
            {
                string json = File.ReadAllText(path);
                return DraftBaseStore.IsBoardDocument(json) ? DataTreeJson.ReadBoard(json) : null;
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not read the draft base [{path}] - [{ex.Message}]");
                return null;
            }
        }

        private static bool IsBoardDocument(string json)
        {
            try
            {
                using JsonDocument parsed = JsonDocument.Parse(json);
                return parsed.RootElement.ValueKind == JsonValueKind.Object && parsed.RootElement.TryGetProperty("formatVersion", out _);
            }
            catch (JsonException)
            {
                return false;
            }
        }

        private static DraftBase FromDocument(BoardDocument document)
        {
            BoardContent content = BoardDocumentMapping.ToContent(document);

            return new DraftBase
            {
                Revision = document.RevisionDate,
                Board = content.Board,
                Calibrations = content.Calibrations.ToList(),
                Files = document.Files.Select(file => new DraftBaseFile
                {
                    Path = file.Name,
                    Sha256 = file.Sha256.ToUpperInvariant(),
                    Length = file.Size,
                }).ToList(),
            };
        }

        // Whether the draft has a base, without reading it.
        public static bool Exists(string draftsRoot, string excelDataFile)
        {
            string path = DraftFolderLayout.GetBasePath(draftsRoot, excelDataFile);
            return path.Length > 0 && File.Exists(path);
        }

        // Written through a temporary file moved into place, so a failure leaves the previous base.
        public static void Save(string draftsRoot, string excelDataFile, DraftBase draftBase)
        {
            ArgumentNullException.ThrowIfNull(draftBase);

            string path = DraftFolderLayout.GetBasePath(draftsRoot, excelDataFile);

            if (path.Length == 0)
                throw new IOException($"No draft folder for [{excelDataFile}].");

            string temporary = DraftFolderLayout.TemporaryPathFor(path);
            File.WriteAllText(temporary, JsonSerializer.Serialize(draftBase, DraftBaseStore.Json));
            File.Move(temporary, path, overwrite: true);
        }

        // ###########################################################################################
        // The base for a draft copied from `published`, the board read from the published workbook
        // at `dataRoot`: its rows and highlights, its calibrations, and the hash of every file of the
        // board's OWN folder its rows reference - exactly the files DraftSeeder copies. A file the
        // published tree lacks is left out (the draft has no copy of it either).
        //
        // `draftsRoot`, when given, records each DRAFT copy's write time for a copy that holds the
        // same bytes (DraftBaseFile.DraftWriteUtcTicks): `draftCopiesAreFresh` says the copies were
        // just made from these very files (the seeder), so they are not read to prove it.
        // ###########################################################################################
        public static DraftBase Capture(
            string dataRoot,
            string excelDataFile,
            BoardData published,
            string? draftsRoot = null,
            bool draftCopiesAreFresh = false)
        {
            ArgumentNullException.ThrowIfNull(published);

            string publishedWorkbook = DraftBoardSource.PublishedPathOf(dataRoot, excelDataFile);
            var files = new List<DraftBaseFile>();

            foreach (string relative in SubmissionManifestBuilder.CollectReferencedFiles(published))
            {
                string? withinBoard = DraftFolderLayout.RelativeToBoardFolder(excelDataFile, relative);

                if (withinBoard is null)
                    continue; // Another board's file - never copied into a draft.

                string source = Path.Combine(dataRoot, relative.Replace('/', Path.DirectorySeparatorChar));

                if (!File.Exists(source))
                    continue;

                string sha256 = DraftFingerprint.CachedFileHash(source);
                long length = new FileInfo(source).Length;
                long draftTicks = 0;

                string draftCopy = string.IsNullOrWhiteSpace(draftsRoot)
                    ? string.Empty
                    : DraftFolderLayout.GetReferencedFilePath(draftsRoot, excelDataFile, relative);

                if (draftCopy.Length > 0 && File.Exists(draftCopy))
                {
                    var copy = new FileInfo(draftCopy);

                    if (copy.Length == length &&
                        (draftCopiesAreFresh || string.Equals(DraftFingerprint.CachedFileHash(draftCopy), sha256, StringComparison.Ordinal)))
                    {
                        draftTicks = copy.LastWriteTimeUtc.Ticks;
                    }
                }

                files.Add(new DraftBaseFile
                {
                    Path = withinBoard,
                    Sha256 = sha256,
                    Length = length,
                    DraftWriteUtcTicks = draftTicks,
                });
            }

            return new DraftBase
            {
                Revision = published.RevisionDate ?? string.Empty,
                Board = published,
                Calibrations = DraftBoardSource.CollectCalibrations(publishedWorkbook, published).ToList(),
                Files = files,
            };
        }

        // ###########################################################################################
        // One SHA-256 over the published workbook's bytes and its sidecar's (empty for none) - what
        // DraftMarker.BaseContentHash remembers. "" when the workbook does not exist or cannot be
        // read, which the drift check treats as "cannot tell".
        // ###########################################################################################
        public static string PublishedContentHash(string publishedWorkbookPath)
        {
            if (string.IsNullOrWhiteSpace(publishedWorkbookPath) || !File.Exists(publishedWorkbookPath))
                return string.Empty;

            string sidecar = BoardComponentHighlightStorage.GetJsonPath(publishedWorkbookPath);
            string stamp = BoardDataReader.FileStamp(publishedWorkbookPath) + "|" + BoardDataReader.FileStamp(sidecar);

            // Trusted only once both writes have settled (StampSettling): two same-length writes
            // inside one clock tick share a stamp.
            if (DraftBaseStore.ContentHashes.TryGetValue(publishedWorkbookPath, out var cached) && cached.Stamp == stamp &&
                StampSettling.IsSettled(File.GetLastWriteTimeUtc(publishedWorkbookPath)) &&
                (!File.Exists(sidecar) || StampSettling.IsSettled(File.GetLastWriteTimeUtc(sidecar))))
            {
                return cached.Hash;
            }

            try
            {
                using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
                DraftBaseStore.AppendFile(hash, publishedWorkbookPath);
                DraftBaseStore.AppendFile(hash, sidecar);

                string result = Convert.ToHexString(hash.GetHashAndReset());
                DraftBaseStore.ContentHashes[publishedWorkbookPath] = (stamp, result);
                return result;
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not read [{publishedWorkbookPath}] to see whether it changed - [{ex.Message}]");
                return string.Empty;
            }
        }

        // The file's bytes with its length in front, or a lone marker for a file that is not there.
        private static void AppendFile(IncrementalHash hash, string path)
        {
            if (!File.Exists(path))
            {
                hash.AppendData(BitConverter.GetBytes(-1L));
                return;
            }

            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            hash.AppendData(BitConverter.GetBytes(stream.Length));

            byte[] buffer = new byte[81920];
            int read;

            while ((read = stream.Read(buffer, 0, buffer.Length)) > 0)
                hash.AppendData(buffer, 0, read);
        }
    }
}
