using Avalonia;
using Avalonia.Controls;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Submit dialog's "what will be sent" panel (BuildSummary). Each line is a label and its value;
// a value longer than the room beside its label WRAPS inside the panel (owner report, 2026-10-26:
// "To upload: worked out when you submit - files the server already h" ran off the dialog's right
// edge). A row laid out as a horizontal StackPanel hands its text unlimited width, so it never
// wraps - the trap the Avalonia notes call "text cut off instead of wrapping".
//
// COLLECTION NOTE: "HeadlessUi" because it shows a Window; UserSettings points at a temp file, as
// the dialog reads the contact address there.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowSummaryTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public SubmitDraftWindowSummaryTests()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings.json", "{}"));
    }

    public void Dispose()
    {
        UserSettings.LoadFrom(this.thisWorkspace.WriteFile("settings-after.json", "{}"));
        this.thisWorkspace.Dispose();
    }

    [Fact]
    public void A_long_value_wraps_inside_the_summary_panel()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();

            window.Initialize(
                "Commodore 64 - 250407",
                new BoardData { Components = [new ComponentEntry { BoardLabel = "U18", Category = "IC" }] },
                new SubmissionIdentity { BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407" },
                this.thisWorkspace.Path_("draft"),
                this.thisWorkspace.Path_("Data"));

            window.Show();
            Dispatcher.UIThread.RunJobs();

            StackPanel panel = window.GetControl<StackPanel>("SummaryPanel");
            TextBlock value = panel.GetVisualDescendants()
                .OfType<TextBlock>()
                .Single(text => (text.Text ?? string.Empty).StartsWith("worked out when you submit", StringComparison.Ordinal));
            // One line's height: the first line, "Schematics: 0" - a label beside the value would be
            // stretched to the value's own height.
            Control oneLine = (Control)panel.Children[0];

            double right = value.TranslatePoint(new Point(value.Bounds.Width, 0), panel)!.Value.X;

            Assert.True(right <= panel.Bounds.Width + 0.5, $"The value ends at {right}, past the panel's {panel.Bounds.Width}.");
            Assert.True(value.Bounds.Height > oneLine.Bounds.Height * 1.5, "The value should wrap onto a second line.");

            window.Close();
        });
    }
}
