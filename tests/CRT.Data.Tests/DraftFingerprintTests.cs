using System;
using System.IO;
using Handlers.DataHandling;
using OfficeOpenXml;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// WHAT A DRAFT HOLDS, AS ONE VALUE (owner request, 2026-10-03: "When a contributor has just
// submitted, then the "Submit" button should be disabled, as the submitted is identical to what
// is in draft now"; cases agreed with the project owner) - DraftFingerprint.
//
// The receipt keeps the fingerprint the draft had when it was sent, and the Drafts tab greys
// Submit out while the draft still gives the same one. So the two directions matter differently:
//
//   - a change the submission would carry MUST change it (cases 2 and 5) - otherwise a
//     contributor could not send real work;
//   - something that changes nothing they send must NOT change it (cases 3, 4 and 11) - otherwise
//     the button comes back for nothing, which is merely the old behaviour.
//
// Each test that says "does not change" also shows one change that DOES, so a fingerprint that
// never changes at all cannot pass it.
// ###########################################################################################
public sealed class DraftFingerprintTests : IDisposable
{
    private const string BoardKey = "Commodore/C64/250407/Data C64 250407.xlsx";

    private readonly TempWorkspace thisWorkspace = new();

    private string DraftsRoot => Path.Combine(this.thisWorkspace.Root, "Drafts");

    private string DataRoot => Path.Combine(this.thisWorkspace.Root, "Data");

    private string Workbook => DraftFolderLayout.GetWorkbookPath(this.DraftsRoot, BoardKey);

    private string Folder => DraftFolderLayout.GetBoardFolder(this.DraftsRoot, BoardKey);

    public void Dispose() => this.thisWorkspace.Dispose();

    private string Fingerprint() => DraftFingerprint.Compute(this.Workbook, this.Folder);

    private static BoardData Board(string friendlyName = "CPU") => new()
    {
        RevisionDate = "2026-09-01",
        Schematics = [new BoardSchematicEntry { SchematicName = "Main", SchematicImageFile = "Commodore/C64/250407/Schematics/main.png" }],
        Components = [new ComponentEntry { BoardLabel = "U1", FriendlyName = friendlyName, TechnicalNameOrValue = "6510" }],
    };

    // Writes the draft workbook, and moves its write time on so nothing can mistake the new file
    // for the old one by its date alone.
    private void WriteDraft(BoardData board)
    {
        Directory.CreateDirectory(Path.GetDirectoryName(this.Workbook)!);
        BoardWorkbookWriter.Write(this.Workbook, board);
        File.SetLastWriteTimeUtc(this.Workbook, DateTime.UtcNow.AddMinutes(this.thisWrites++));
    }

    private int thisWrites = 1;

    private string DraftFile(string relativePath, string content)
    {
        string path = Path.Combine(this.Folder, relativePath);
        Directory.CreateDirectory(Path.GetDirectoryName(path)!);
        File.WriteAllText(path, content);
        File.SetLastWriteTimeUtc(path, DateTime.UtcNow.AddMinutes(this.thisWrites++));
        return path;
    }

    // A draft whose fingerprint is the one "sent": a workbook, a highlight and a calibration in
    // its JSON, and one picture of its own.
    private string SentDraft()
    {
        this.WriteDraft(Board());

        BoardComponentHighlightStorage.SaveComponentHighlights(
            this.Workbook,
            "Main",
            [new LabelEditorSaveRow { SchematicName = "Main", BoardLabel = "U1", X = 10, Y = 10, Width = 20, Height = 20 }]);

        BoardComponentHighlightStorage.SaveKiCadCalibration(this.Workbook, "Main", "main.kicad_pcb", 0, 0, 1, 1, false, false);

        this.DraftFile("Schematics/main.png", "picture one");

        return this.Fingerprint();
    }

    // Case 2.
    [Fact]
    public void Changing_one_cell_and_saving_changes_the_fingerprint()
    {
        string sent = this.SentDraft();

        this.WriteDraft(Board("CPU 6510"));

        Assert.NotEqual(sent, this.Fingerprint());
    }

    // Case 3.
    [Fact]
    public void Changing_the_cell_back_to_what_was_sent_gives_the_sent_fingerprint_again()
    {
        string sent = this.SentDraft();

        this.WriteDraft(Board("CPU 6510"));
        Assert.NotEqual(sent, this.Fingerprint());

        // The file is a new file - another date, and bytes EPPlus wrote afresh - holding the
        // values that were sent. The content is what counts.
        this.WriteDraft(Board());
        Assert.Equal(sent, this.Fingerprint());
    }

    // Case 4: opened in Excel and saved with nothing changed. Excel rewrites the file (another
    // date, other bytes) and keeps its "~$" owner file beside it while the workbook is open.
    [Fact]
    public void Saving_the_workbook_in_Excel_without_changing_a_value_keeps_the_fingerprint()
    {
        string sent = this.SentDraft();
        byte[] before = File.ReadAllBytes(this.Workbook);

        EpplusLicense.Ensure();
        using (var package = EpplusLicense.OpenPackage(new FileInfo(this.Workbook)))
        {
            package.Workbook.Properties.Comments = "Saved again in Excel";
            package.Save();
        }

        File.SetLastWriteTimeUtc(this.Workbook, DateTime.UtcNow.AddHours(1));
        this.DraftFile($"~${Path.GetFileName(this.Workbook)}", "Excel's owner file");

        Assert.NotEqual(before, File.ReadAllBytes(this.Workbook));
        Assert.Equal(sent, this.Fingerprint());

        // ...while a changed value still changes it.
        this.WriteDraft(Board("CPU 6510"));
        Assert.NotEqual(sent, this.Fingerprint());
    }

    // Case 5: every other kind of change a submission carries. (An edit in Excel and "Save to
    // draft" in the Contribute tab both change the workbook's values - case 2.)
    [Theory]
    [InlineData("a highlight moved in the label editor")]
    [InlineData("a KiCad calibration changed")]
    [InlineData("a file added to Schematic images")]
    [InlineData("a file removed")]
    [InlineData("a file added to KiCad data")]
    [InlineData("a picture replaced under the same name")]
    public void Every_kind_of_change_a_submission_carries_changes_the_fingerprint(string change)
    {
        string sent = this.SentDraft();

        switch (change)
        {
            case "a highlight moved in the label editor":
                BoardComponentHighlightStorage.SaveComponentHighlights(
                    this.Workbook,
                    "Main",
                    [new LabelEditorSaveRow { SchematicName = "Main", BoardLabel = "U1", X = 30, Y = 10, Width = 20, Height = 20 }]);
                break;

            case "a KiCad calibration changed":
                BoardComponentHighlightStorage.SaveKiCadCalibration(this.Workbook, "Main", "main.kicad_pcb", 5, 0, 1, 1, false, false);
                break;

            case "a file added to Schematic images":
                this.DraftFile("Schematics/main back.png", "picture two");
                break;

            case "a file removed":
                File.Delete(Path.Combine(this.Folder, "Schematics", "main.png"));
                break;

            case "a file added to KiCad data":
                this.DraftFile("KiCad data/board.kicad_pcb", "(kicad_pcb)");
                break;

            case "a picture replaced under the same name":
                // The same length, so only the bytes say it is another picture.
                this.DraftFile("Schematics/main.png", "picture two");
                break;
        }

        Assert.NotEqual(sent, this.Fingerprint());
    }

    // Case 11: what the draft uses from CRT's downloaded data is not the contributor's change, so
    // a data sync that replaces such a file leaves the fingerprint alone.
    [Fact]
    public void A_file_changing_in_the_downloaded_data_does_not_change_the_fingerprint()
    {
        string published = Path.Combine(this.DataRoot, "Commodore", "C64", "250407", "Schematics", "main.png");
        Directory.CreateDirectory(Path.GetDirectoryName(published)!);
        File.WriteAllText(published, "published picture");

        string sent = this.SentDraft();

        File.WriteAllText(published, "a newer published picture");
        File.SetLastWriteTimeUtc(published, DateTime.UtcNow.AddHours(1));

        Assert.Equal(sent, this.Fingerprint());

        // ...while the draft's own copy changing does.
        this.DraftFile("Schematics/main.png", "picture two");
        Assert.NotEqual(sent, this.Fingerprint());
    }

    // An empty fingerprint means "could not tell", which never greys Submit out - so a real
    // draft must always give one, or the whole rule would quietly do nothing.
    [Fact]
    public void A_drafts_fingerprint_is_never_empty()
    {
        Assert.NotEqual(string.Empty, this.SentDraft());
    }
}
