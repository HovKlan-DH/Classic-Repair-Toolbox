using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // READS AND WRITES ONE CRT 4.0 BOARD FILE ("board.json") - the one place the app touches it.
    //
    // The shared readers hand a board.json here instead of opening a workbook (BoardDataReader,
    // BoardComponentHighlightStorage - see DataTreeLayout.IsBoardFile), and every write goes through
    // Write: the bytes DataTreeJson makes, replaced whole through a temporary file beside it
    // (FileReplacer), so a reader never sees half a board and a failed write leaves the old one.
    //
    // A board file that is missing reads as null, quietly; one that is there and cannot be read
    // reads as null with a warning naming it - the readers' "cannot load this board", as for a
    // broken workbook. Nothing here caches: BoardDataReader does, by the file's stamp.
    // ###########################################################################################
    public static class BoardFiles
    {
        public static BoardDocument? ReadDocument(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return null;

            try
            {
                return DataTreeJson.ReadBoard(BoardFiles.ReadAllTextShared(path));
            }
            catch (Exception ex) when (ex is InvalidDataException or IOException or UnauthorizedAccessException)
            {
                CrtLog.Warning($"Could not read the board file [{path}] - [{ex.Message}]");
                return null;
            }
        }

        public static BoardContent? ReadContent(string? path) =>
            BoardFiles.ReadDocument(path) is BoardDocument document ? BoardDocumentMapping.ToContent(document) : null;

        // ###########################################################################################
        // Writes the document as `path`, creating its folder. Throws IOException or
        // UnauthorizedAccessException when it cannot - with the file as it was.
        // ###########################################################################################
        public static void Write(string path, BoardDocument document)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(path);
            ArgumentNullException.ThrowIfNull(document);

            byte[] bytes = new UTF8Encoding(encoderShouldEmitUTF8Identifier: false).GetBytes(DataTreeJson.Write(document));

            Directory.CreateDirectory(Path.GetDirectoryName(Path.GetFullPath(path))!);
            FileReplacer.Replace(path, temporary => File.WriteAllBytes(temporary, bytes));
        }

        // ###########################################################################################
        // Reads the board file, changes it, and writes it back. Throws InvalidOperationException,
        // writing nothing, when the file is there and cannot be read IN FULL - writing a board
        // rebuilt from what could not be read would drop the rest (the rule
        // BoardComponentHighlightStorage.EnsureReadableInFull keeps for a sidecar). A missing file
        // is changed from an empty board.
        // ###########################################################################################
        public static void Change(string path, Action<BoardDocument> change)
        {
            ArgumentNullException.ThrowIfNull(change);

            BoardDocument document = BoardFiles.ReadInFull(path) ?? new BoardDocument();
            change(document);
            BoardFiles.Write(path, document);
        }

        // The document, or null when there is no file; InvalidOperationException when there is one
        // that does not read in full.
        public static BoardDocument? ReadInFull(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path))
                return null;

            try
            {
                return DataTreeJson.ReadBoard(BoardFiles.ReadAllTextShared(path));
            }
            catch (InvalidDataException ex)
            {
                throw new InvalidOperationException(
                    $"The board file [{Path.GetFileName(path)}] could not be read in full, so nothing was saved over it - [{ex.Message}].",
                    ex);
            }
        }

        // ###########################################################################################
        // The highlights, KiCad calibrations and schematic renames of the board file replaced by
        // these - the three parts the Excel generations kept in the sidecar, written here by the
        // sidecar writer's callers (BoardSidecarWriter.Write) when they are handed a board.json.
        // ###########################################################################################
        public static void WriteSidecarParts(
            string path,
            IReadOnlyList<ComponentHighlightEntry> highlights,
            IReadOnlyList<KiCadCalibrationEntry> calibrations,
            IReadOnlyList<SchematicRenameRecord> renames)
        {
            ArgumentNullException.ThrowIfNull(highlights);
            ArgumentNullException.ThrowIfNull(calibrations);
            ArgumentNullException.ThrowIfNull(renames);

            BoardFiles.Change(path, document =>
            {
                // A carrier for the three parts only: the board file's own caption is left as it is.
                BoardDocument parts = BoardDocumentMapping.ToDocument(
                    new BoardContent(
                        new BoardData { HardwareName = string.Empty, BoardName = string.Empty, ComponentHighlights = highlights.ToList() },
                        calibrations,
                        renames),
                    name => name,
                    []);

                document.Highlights = parts.Highlights;
                document.KiCadCalibrations = parts.KiCadCalibrations;
                document.SchematicRenames = parts.SchematicRenames;
            });
        }

        // Readable while another program has it open for reading (and, on Windows, for writing).
        private static string ReadAllTextShared(string path)
        {
            using var stream = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite | FileShare.Delete);
            using var reader = new StreamReader(stream, Encoding.UTF8);

            return reader.ReadToEnd();
        }
    }
}
