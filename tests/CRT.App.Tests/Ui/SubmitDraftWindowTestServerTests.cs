using Avalonia.Controls;
using CRT;
using Handlers.Online;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// The Submit dialog of a CRT talking to the contribution service's TEST instance (CRT 4.0's
// pre-releases, 2026-10-25) says so: what it publishes there is test data, and a tester must never
// take it for a contribution everybody now has.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowTestServerTests
{
    [Fact]
    public void A_CRT_talking_to_the_test_instance_says_its_changes_go_there()
    {
        UiTest.Run(() =>
        {
            var window = new SubmitDraftWindow();
            TextBlock note = window.GetControl<TextBlock>("TestServerSubmitText");

            // As this CRT is built.
            Assert.Equal(AppConfig.UsesTestServer, note.IsVisible);

            window.ShowTestServerNote(usesTestServer: true);
            Assert.True(note.IsVisible);
            Assert.Equal(ContributionServer.SubmitNote, note.Text);

            window.ShowTestServerNote(usesTestServer: false);
            Assert.False(note.IsVisible);
        });
    }
}
