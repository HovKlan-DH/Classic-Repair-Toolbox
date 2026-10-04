using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT A DRAFT HOLDS, AS ONE VALUE (owner request, 2026-10-03: "When a contributor has just
    // submitted, then the "Submit" button should be disabled, as the submitted is identical to what
    // is in draft now. It should not be possible to submit the same data again").
    //
    // The receipt keeps this value from the moment of sending (SubmissionReceipt.DraftFingerprint),
    // and the Drafts tab greys Submit out while the draft still gives the same one
    // (SubmissionReceiptPresenter.IsAlreadySent). It is made of everything the contributor can
    // change that a submission carries:
    //
    //   - the workbook's VALUES - every sheet's cells as BoardWorkbookSchema writes them, plus the
    //     caption and revision date - not its bytes. Excel rewrites a file it merely saved (another
    //     date, other bytes), and a cell changed and changed back is the board that was sent
    //     (cases 3 and 4 agreed with the project owner);
    //   - the JSON beside it (highlights, KiCad calibrations), by its bytes - only CRT writes it;
    //   - every other file in the draft's own folder, by path and bytes: pictures, documents and
    //     the KiCad data. A picture replaced under its own name is a change (case 5).
    //
    // *** WHAT IS LEFT OUT, and why each matters. *** Files the draft uses from CRT's DOWNLOADED
    // data are not the contributor's change - a data sync replacing one must not bring the button
    // back (case 11) - so only the draft folder is read. Inside it, local bookkeeping never counts:
    // the draft marker, and the owner files Excel ("~$...") and LibreOffice (".~lock...#") keep
    // beside a workbook while it is open - otherwise merely opening the workbook in Excel would
    // count as a change. Hidden files, Thumbs.db and desktop.ini are the operating system's.
    //
    // *** WHICH WAY IT MAY BE WRONG. *** Something counted that is not really a change only brings
    // Submit back - today's behaviour. Something missed that IS a change would stop a contributor
    // sending real work, so anything this cannot read gives an EMPTY fingerprint, which never
    // greys anything out. "v1:" in front lets a later way of computing it never match an old one.
    //
    // Each file's hash is remembered by its length and write time, so a draft nobody touched costs
    // a directory listing - this runs for every draft on every refresh of the Drafts tab.
    // ###########################################################################################
    public static class DraftFingerprint
    {
        private const string Version = "v1";

        private static readonly ConcurrentDictionary<string, CachedHash> FileHashes = new(StringComparer.OrdinalIgnoreCase);

        private static readonly ConcurrentDictionary<string, CachedHash> WorkbookHashes = new(StringComparer.OrdinalIgnoreCase);

        public static string Compute(string workbookPath, string draftSystemFolder)
        {
            try
            {
                if (string.IsNullOrWhiteSpace(workbookPath) || !File.Exists(workbookPath))
                    return string.Empty;

                string? values = DraftFingerprint.WorkbookValuesHash(workbookPath);
                if (values is null)
                    return string.Empty;

                string sidecar = BoardComponentHighlightStorage.GetJsonPath(workbookPath);

                using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);

                DraftFingerprint.Add(hash, "workbook");
                DraftFingerprint.Add(hash, values);
                DraftFingerprint.Add(hash, "sidecar");
                DraftFingerprint.Add(hash, File.Exists(sidecar) ? DraftFingerprint.FileHash(sidecar) : string.Empty);

                foreach ((string relative, string full) in DraftFingerprint.OwnFiles(draftSystemFolder, workbookPath, sidecar))
                {
                    DraftFingerprint.Add(hash, relative);
                    DraftFingerprint.Add(hash, DraftFingerprint.FileHash(full));
                }

                return $"{DraftFingerprint.Version}:{Convert.ToHexString(hash.GetHashAndReset())}";
            }
            catch (Exception ex)
            {
                // Unknown never greys Submit out - see the header.
                CrtLog.Warning($"Could not work out what the draft [{workbookPath}] holds: [{ex.Message}]");
                return string.Empty;
            }
        }

        // The draft folder's files the contributor put there, in a fixed order.
        private static IEnumerable<(string Relative, string Full)> OwnFiles(string folder, string workbookPath, string sidecar)
        {
            if (string.IsNullOrWhiteSpace(folder) || !Directory.Exists(folder))
                return [];

            string workbook = Path.GetFullPath(workbookPath);
            string json = Path.GetFullPath(sidecar);

            return Directory.EnumerateFiles(folder, "*", SearchOption.AllDirectories)
                .Select(Path.GetFullPath)
                .Where(full => !string.Equals(full, workbook, StringComparison.OrdinalIgnoreCase))
                .Where(full => !string.Equals(full, json, StringComparison.OrdinalIgnoreCase))
                .Where(full => !DraftFingerprint.IsBookkeeping(Path.GetFileName(full)))
                .Select(full => (Relative: Path.GetRelativePath(folder, full).Replace('\\', '/'), Full: full))
                .OrderBy(file => file.Relative, StringComparer.OrdinalIgnoreCase)
                .ThenBy(file => file.Relative, StringComparer.Ordinal)
                .ToList();
        }

        internal static bool IsBookkeeping(string name) =>
            DraftFolderLayout.IsDraftOnlyFile(name)
            || name.StartsWith("~$", StringComparison.Ordinal)
            || name.StartsWith('.')
            || string.Equals(name, "Thumbs.db", StringComparison.OrdinalIgnoreCase)
            || string.Equals(name, "desktop.ini", StringComparison.OrdinalIgnoreCase);

        // The values of every sheet, as the schema writes them - null when the workbook cannot be read.
        private static string? WorkbookValuesHash(string workbookPath)
        {
            var info = new FileInfo(workbookPath);

            if (DraftFingerprint.WorkbookHashes.TryGetValue(info.FullName, out CachedHash cached) && cached.Matches(info))
                return cached.Hash;

            BoardData? board = BoardDataReader.ReadWorkbookUncached(workbookPath);
            if (board is null)
                return null;

            using var hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);

            DraftFingerprint.Add(hash, board.HardwareName);
            DraftFingerprint.Add(hash, board.BoardName);
            DraftFingerprint.Add(hash, board.RevisionDate);

            foreach (BoardWorkbookSchema.SheetDefinition sheet in BoardWorkbookSchema.AllSheets)
            {
                DraftFingerprint.Add(hash, sheet.SheetName);

                foreach (IReadOnlyDictionary<string, string> row in BoardWorkbookSchema.BuildRows(sheet, board))
                {
                    DraftFingerprint.Add(hash, "row");

                    foreach (string column in sheet.ColumnOrder)
                    {
                        DraftFingerprint.Add(hash, row.TryGetValue(column, out string? value) ? value : string.Empty);
                    }
                }
            }

            string result = Convert.ToHexString(hash.GetHashAndReset());
            DraftFingerprint.WorkbookHashes[info.FullName] = new CachedHash(info.Length, info.LastWriteTimeUtc, result);

            return result;
        }

        private static string FileHash(string path)
        {
            var info = new FileInfo(path);

            if (DraftFingerprint.FileHashes.TryGetValue(info.FullName, out CachedHash cached) && cached.Matches(info))
                return cached.Hash;

            string result;
            using (var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete))
            {
                result = Convert.ToHexString(SHA256.HashData(stream));
            }

            DraftFingerprint.FileHashes[info.FullName] = new CachedHash(info.Length, info.LastWriteTimeUtc, result);

            return result;
        }

        // Each part with its length in front, so "ab" + "c" and "a" + "bc" never hash alike.
        private static void Add(IncrementalHash hash, string? text)
        {
            byte[] bytes = Encoding.UTF8.GetBytes(text ?? string.Empty);
            hash.AppendData(BitConverter.GetBytes(bytes.Length));
            hash.AppendData(bytes);
        }

        private readonly record struct CachedHash(long Length, DateTime WriteUtc, string Hash)
        {
            public bool Matches(FileInfo info) => this.Length == info.Length && this.WriteUtc == info.LastWriteTimeUtc;
        }
    }
}
