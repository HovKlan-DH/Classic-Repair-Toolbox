using System.Collections.Generic;
using System.IO;
using System.Linq;
using Avalonia.Controls;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Submit dialog's own side of "the board's KiCad data travels" (owner decision, 2026-09-26):
// the window collects the draft's "KiCad data" folder (SubmitDraftWindow.CollectKiCadFiles - the
// same helper the submission itself sends through) and says so in the "what will be sent" panel.
// Without this wire-up the pure rules all pass while the client still sends nothing, which is the
// exact one-side gap the KiCad data sat in until today.
//
// In "HeadlessUi" because it constructs a Window. The network path is not driven here (rule 6);
// what is pinned is that the collection the submission uses sees the draft's folder.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowKiCadTests : System.IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private static SubmissionIdentity Identity() => new()
    {
        BoardId = "Manu1/Hardware1/Board1",
        Manufacturer = "Manu1",
        Hardware = "Hardware1",
        Board = "Board1"
    };

    private SubmitDraftWindow BuildWindow(string draftFolder)
    {
        var window = new SubmitDraftWindow();

        window.Initialize(
            "Manu1 Hardware1 Board1",
            new BoardData(),
            SubmitDraftWindowKiCadTests.Identity(),
            draftFolder,
            this.thisWorkspace.Path_("Data"));

        return window;
    }

    private static IReadOnlyList<string> SummaryTexts(SubmitDraftWindow window) =>
        window.GetControl<StackPanel>("SummaryPanel")
            .GetVisualDescendants()
            .OfType<TextBlock>()
            .Select(text => text.Text ?? string.Empty)
            .ToList();

    [Fact]
    public void A_drafts_KiCad_data_is_counted_in_what_will_be_sent()
    {
        string draft = this.thisWorkspace.Path_("draft");
        string kiCad = Path.Combine(draft, "KiCad data", "Pages");
        Directory.CreateDirectory(kiCad);
        File.WriteAllText(Path.Combine(kiCad, "vic.kicad_sch"), "(kicad_sch)");
        File.WriteAllText(Path.Combine(draft, "KiCad data", "board.kicad_pcb"), "(kicad_pcb)");

        UiTest.Run(() =>
        {
            IReadOnlyList<string> texts = SubmitDraftWindowKiCadTests.SummaryTexts(this.BuildWindow(draft));

            int label = texts.ToList().IndexOf("KiCad files:");
            Assert.True(label >= 0, string.Join(" | ", texts));
            Assert.Equal("2", texts[label + 1]);
        });
    }

    // Most boards have no KiCad data; a permanent "KiCad files: 0" would read as something missing.
    [Fact]
    public void A_board_without_KiCad_data_gets_no_KiCad_line()
    {
        string draft = this.thisWorkspace.Path_("plain-draft");
        Directory.CreateDirectory(draft);

        UiTest.Run(() =>
        {
            Assert.DoesNotContain(
                "KiCad files:",
                SubmitDraftWindowKiCadTests.SummaryTexts(this.BuildWindow(draft)));
        });
    }
}
