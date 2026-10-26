using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// Drafts as CRT 3.x made them - the workbook under its published file name, its JSON sidecar,
// copies of the board's own files under their paths, the base and the marker - for the tests of
// what reads them now: the conversion into a 4.0 draft (DraftConversion) and the import of a
// folder put into Drafts/ by hand (DraftFolderImport).
//
// CRT 4.0's seeder and draft writes make 4.0 drafts, so these are made here the way 3.1.0's
// DraftSeeder.SeedFromPublished and DraftWorkbookStore.Edit made them (git history,
// src/CRT.Data/DraftSeeder.cs before 2026-10-24).
// ###########################################################################################
internal static class LegacyDrafts
{
    // The workbook of an Excel-era draft: in the board's draft folder, under its published name.
    public static string WorkbookPath(string draftsRoot, string excelDataFile) =>
        Path.Combine(DraftFolderLayout.GetBoardFolder(draftsRoot, excelDataFile), Path.GetFileName(excelDataFile));

    // Copies the published board at `excelDataFile` under `dataRoot` (an Excel tree) into a draft.
    public static void Seed(string draftsRoot, string dataRoot, string excelDataFile)
    {
        string publishedPath = Path.Combine(dataRoot, excelDataFile.Replace('/', Path.DirectorySeparatorChar));
        BoardData published = BoardDataReader.ReadWorkbookUncached(publishedPath)
            ?? throw new InvalidOperationException($"[{publishedPath}] could not be read");

        string folder = DraftFolderLayout.GetBoardFolder(draftsRoot, excelDataFile);
        string workbookPath = LegacyDrafts.WorkbookPath(draftsRoot, excelDataFile);
        Directory.CreateDirectory(folder);

        BoardWorkbookWriter.Write(workbookPath, published);

        string publishedSidecar = BoardComponentHighlightStorage.GetJsonPath(publishedPath);
        if (File.Exists(publishedSidecar))
            File.Copy(publishedSidecar, BoardComponentHighlightStorage.GetJsonPath(workbookPath), overwrite: true);

        foreach (string relative in SubmissionManifestBuilder.CollectReferencedFiles(published))
        {
            string destination = DraftFolderLayout.GetReferencedFilePath(draftsRoot, excelDataFile, relative);
            string source = Path.Combine(dataRoot, relative.Replace('/', Path.DirectorySeparatorChar));

            if (destination.Length == 0 || !File.Exists(source))
                continue;

            Directory.CreateDirectory(Path.GetDirectoryName(destination)!);
            File.Copy(source, destination, overwrite: true);
        }

        DraftBaseStore.Save(
            draftsRoot,
            excelDataFile,
            DraftBaseStore.Capture(dataRoot, excelDataFile, published, draftsRoot, draftCopiesAreFresh: true));

        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile), new DraftMarker
        {
            BoardKey = excelDataFile,
            BaseRevision = published.RevisionDate ?? string.Empty,
            BaseContentHash = DraftBaseStore.PublishedContentHash(publishedPath),
            CreatedUtc = DateTimeOffset.UtcNow.ToString("o"),
        });
    }

    // A board that exists only as an Excel-era draft ("Add a new board").
    public static void CreateNewBoard(string draftsRoot, NewBoardRegistration registration)
    {
        string excelDataFile = registration.ExcelDataFile ?? string.Empty;
        Directory.CreateDirectory(DraftFolderLayout.GetBoardFolder(draftsRoot, excelDataFile));

        BoardWorkbookWriter.Write(LegacyDrafts.WorkbookPath(draftsRoot, excelDataFile), new BoardData
        {
            HardwareName = registration.HardwareName ?? string.Empty,
            BoardName = registration.BoardName ?? string.Empty,
        });

        DraftMarkerStore.Save(DraftFolderLayout.GetMarkerPath(draftsRoot, excelDataFile), new DraftMarker
        {
            BoardKey = excelDataFile,
            BaseRevision = string.Empty,
            NewBoard = registration,
            CreatedUtc = DateTimeOffset.UtcNow.ToString("o"),
        });
    }

    // Changes an Excel-era draft's workbook - the rows only; the sidecar stays as it is.
    public static void Edit(string draftsRoot, string excelDataFile, Func<BoardData, BoardData> change)
    {
        string workbookPath = LegacyDrafts.WorkbookPath(draftsRoot, excelDataFile);
        BoardData board = BoardDataReader.ReadWorkbookUncached(workbookPath)
            ?? throw new InvalidOperationException($"[{workbookPath}] could not be read");

        BoardWorkbookWriter.Write(workbookPath, change(board));
    }
}
