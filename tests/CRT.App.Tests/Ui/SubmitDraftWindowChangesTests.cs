using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Submit dialog's "Your changes" (CRT 4.0 step 1's last slice, 2026-10-25 - the plan: "Review
// changes shows a readable summary, and the Submit dialog shows the same"): the summary Review
// changes shows, group title then its lines, in the same words - and nothing when there is none.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowChangesTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose() => this.thisWorkspace.Dispose();

    private SubmitDraftWindow BuildWindow(DraftChangeSummary? changes)
    {
        var window = new SubmitDraftWindow();

        window.Initialize(
            "Commodore 64 - 250407",
            new BoardData(),
            new SubmissionIdentity { BoardId = "Commodore/C64/250407", Manufacturer = "Commodore", Hardware = "C64", Board = "250407" },
            this.thisWorkspace.Path_("draft"),
            this.thisWorkspace.Path_("Data"),
            changes: changes);

        return window;
    }

    private static BoardContent Content(params ComponentEntry[] components) =>
        new(new BoardData { Components = [.. components] }, [], []);

    [Fact]
    public void The_dialog_shows_the_same_summary_as_Review_changes()
    {
        UiTest.Run(() =>
        {
            DraftChangeSummary changes = DraftChangeSummary.Build(
                SubmitDraftWindowChangesTests.Content(new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU", Description = "old" }),
                SubmitDraftWindowChangesTests.Content(
                    new ComponentEntry { BoardLabel = "U8", Category = "IC", FriendlyName = "CPU", Description = "new" },
                    new ComponentEntry { BoardLabel = "U9", Category = "IC", FriendlyName = "VIC" }));

            SubmitDraftWindow window = this.BuildWindow(changes);

            Assert.Equal(
                changes.Groups.SelectMany(group => new[] { group.Title }.Concat(group.Lines.Select(line => "- " + line))),
                window.ChangesForTests);
            Assert.Equal(["U8 - CPU", "- Description changed", "U9 - VIC", "- New component"], window.ChangesForTests);
            Assert.True(window.GetControl<TextBlock>("ChangesHeading").IsVisible);
        });
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void With_no_summary_or_an_empty_one_the_section_is_hidden(bool empty)
    {
        UiTest.Run(() =>
        {
            SubmitDraftWindow window = this.BuildWindow(empty ? DraftChangeSummary.Empty : null);

            Assert.Empty(window.ChangesForTests);
            Assert.False(window.GetControl<TextBlock>("ChangesHeading").IsVisible);
        });
    }
}
