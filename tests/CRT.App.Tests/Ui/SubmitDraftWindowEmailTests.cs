using System;
using Avalonia.Controls;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// ###########################################################################################
// ONE EMAIL ADDRESS FOR THE WHOLE APPLICATION (owner request, 2026-09-22).
//
// *** THE SETTING ALREADY EXISTED; THIS SCREEN JUST DID NOT READ IT. *** UserSettings.ContactEmail
// has been written by the Feedback tab since long before submissions existed, so a contributor who
// had already given their address there was asked for it again when submitting - with no sign the
// app already knew it. Reported as exactly that.
//
// *** IT IS TESTED THROUGH THE REAL UserSettings STATIC, via LoadFrom pointed at a temp file. ***
// That is the seam CLAUDE.md requires (never UserSettings.Load(), which reads the user's own
// AppData), and going through the real static is what makes these tests prove the shipped path
// rather than a parallel one.
//
// In "HeadlessUi" because it constructs a Window, which needs the shared dispatcher thread. Note
// that this puts it OUTSIDE the "UserSettings" collection - a class can only join one - so each
// test points the settings store at its OWN file rather than relying on collection isolation.
// ###########################################################################################
[Collection("HeadlessUi")]
public sealed class SubmitDraftWindowEmailTests : IDisposable
{
    private readonly TempWorkspace thisWorkspace = new();

    public void Dispose()
    {
        // This class WRITES to the shared UserSettings static, so it cleans up after itself for the
        // same reason FreshSettings has to clear: the next class to read ContactEmail would
        // otherwise see an address this one saved. Symmetry with the setup, not superstition.
        UserSettings.ContactEmail = string.Empty;

        this.thisWorkspace.Dispose();
    }

    // ###########################################################################################
    // A fresh settings file per test - AND the address explicitly cleared.
    //
    // *** POINTING AT A NEW PATH IS NOT ENOUGH, and assuming it was cost a real failure. ***
    // UserSettings.LoadFrom returns EARLY when the file does not exist ("using defaults"), leaving
    // the static _data object exactly as the previous test left it. So a test expecting an empty
    // address passed alone and failed in the full run, as soon as another test had saved one - the
    // classic shared-static symptom CLAUDE.md describes (the tell is that it never fails in
    // isolation).
    //
    // Clearing the value is what actually isolates these tests; the fresh path only stops them
    // writing over each other's files.
    // ###########################################################################################
    private void FreshSettings()
    {
        UserSettings.LoadFrom(this.thisWorkspace.Path_(Guid.NewGuid().ToString("N") + ".json"));
        UserSettings.ContactEmail = string.Empty;
    }

    private static BoardData EmptyBoard() => new();

    private static SubmissionIdentity Identity() => new()
    {
        SystemId = "Commodore/C64/250407/Data.xlsx",
        Manufacturer = "Commodore",
        Hardware = "C64",
        Board = "250407"
    };

    private SubmitDraftWindow BuildWindow()
    {
        var window = new SubmitDraftWindow();

        window.Initialize(
            "Commodore C64 250407",
            SubmitDraftWindowEmailTests.EmptyBoard(),
            SubmitDraftWindowEmailTests.Identity(),
            this.thisWorkspace.Root,
            this.thisWorkspace.Root);

        return window;
    }

    [Fact]
    public void The_email_box_is_PREFILLED_from_the_address_given_anywhere_else()
    {
        UiTest.Run(() =>
        {
            this.FreshSettings();

            // As the Feedback tab would have saved it.
            UserSettings.ContactEmail = "dennis@example.com";

            SubmitDraftWindow window = this.BuildWindow();

            Assert.Equal(
                "dennis@example.com",
                window.GetControl<TextBox>("EmailTextBox").Text);
        });
    }

    [Fact]
    public void With_NO_address_stored_the_box_is_simply_empty()
    {
        UiTest.Run(() =>
        {
            this.FreshSettings();

            SubmitDraftWindow window = this.BuildWindow();

            Assert.True(string.IsNullOrEmpty(
                window.GetControl<TextBox>("EmailTextBox").Text));
        });
    }

    [Fact]
    public void A_value_ALREADY_IN_THE_BOX_is_not_replaced_by_the_stored_one()
    {
        UiTest.Run(() =>
        {
            this.FreshSettings();
            UserSettings.ContactEmail = "stored@example.com";

            var window = new SubmitDraftWindow();

            // A caller that pre-populated the dialog before initialising it. Silently overwriting
            // that would be a surprise that is very hard to see in a UI.
            window.GetControl<TextBox>("EmailTextBox").Text = "typed@example.com";

            window.Initialize(
                "Commodore C64 250407",
                SubmitDraftWindowEmailTests.EmptyBoard(),
                SubmitDraftWindowEmailTests.Identity(),
                this.thisWorkspace.Root,
                this.thisWorkspace.Root);

            Assert.Equal(
                "typed@example.com",
                window.GetControl<TextBox>("EmailTextBox").Text);
        });
    }

    [Fact]
    public void The_box_STAYS_EDITABLE_so_a_different_address_can_be_used()
    {
        UiTest.Run(() =>
        {
            // Prefilled, not locked: a contributor may genuinely want a different address on a
            // contribution than on a bug report, and this is the one that a person acts on.
            this.FreshSettings();
            UserSettings.ContactEmail = "dennis@example.com";

            SubmitDraftWindow window = this.BuildWindow();

            var box = window.GetControl<TextBox>("EmailTextBox");

            Assert.False(box.IsReadOnly);
            Assert.True(box.IsEnabled);
        });
    }

    [Fact]
    public void An_address_typed_HERE_is_remembered_for_the_rest_of_the_app()
    {
        UiTest.Run(() =>
        {
            // The other half of "one address everywhere": entering it while submitting must fill
            // the Feedback tab too, or the sharing only works in one direction.
            this.FreshSettings();

            SubmitDraftWindow window = this.BuildWindow();

            window.GetControl<TextBox>("EmailTextBox").Text = "  new@example.com  ";

            window.CaptureEntriesForTests();

            // Trimmed, because that is what is actually sent with the submission.
            Assert.Equal("new@example.com", UserSettings.ContactEmail);
        });
    }

    [Fact]
    public void An_IMPLAUSIBLE_address_is_NOT_saved_over_a_good_one()
    {
        UiTest.Run(() =>
        {
            // Saving "dennis@" would quietly poison the prefill on every other screen. The submit
            // button is disabled for an implausible address anyway, so this is belt and braces
            // against a future caller reaching this path another way.
            this.FreshSettings();
            UserSettings.ContactEmail = "good@example.com";

            SubmitDraftWindow window = this.BuildWindow();

            window.GetControl<TextBox>("EmailTextBox").Text = "not-an-address";

            window.CaptureEntriesForTests();

            Assert.Equal("good@example.com", UserSettings.ContactEmail);
        });
    }
}
