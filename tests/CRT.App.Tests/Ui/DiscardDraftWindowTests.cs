using Avalonia.Controls;
using Avalonia.Headless;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.VisualTree;
using CRT;
using Handlers.DataHandling;

namespace ClassicRepairToolbox.Tests.Ui;

// The discard-draft confirmation modal's keyboard behaviour - the twin of
// DeleteWorkbookWindowTests, and for the same reason: its "submit" is a permanent discard of
// local, unpublished work, so Enter and Escape both CANCEL, including with the Discard button
// focused (the Tunnel-vs-bubbling regression - see DeleteWorkbookWindowTests' own header for the
// full explanation of why a real keypress through the headless input stack is what makes this
// test able to fail against a bubbling KeyDown subscription).
[Collection("HeadlessUi")]
public sealed class DiscardDraftWindowTests
{
    private static DiscardDraftWindow BuildWindow()
    {
        var window = new DiscardDraftWindow();
        window.Initialize("Commodore 64 - 250469 (short board)");

        // Show() so the visual tree is built and focus can actually land on a button.
        window.Show();
        return window;
    }

    private static Button ButtonWithContent(DiscardDraftWindow window, string content) =>
        window.GetVisualDescendants()
            .OfType<Button>()
            .First(b => (b.Content as string) == content);

    private static void PressKey(DiscardDraftWindow window, Key key, PhysicalKey physicalKey) =>
        window.KeyPress(key, RawInputModifiers.None, physicalKey, keySymbol: null);

    [Fact]
    public void Enter_cancels_even_when_the_discard_button_has_focus()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();
            var discardButton = ButtonWithContent(window, "Discard draft");

            bool discardConfirmed = false;
            discardButton.Click += (_, _) => discardConfirmed = true;

            discardButton.Focus();
            PressKey(window, Key.Enter, PhysicalKey.Enter);

            Assert.False(discardConfirmed);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void Enter_cancels_when_the_cancel_button_has_focus()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            bool discardConfirmed = false;
            ButtonWithContent(window, "Discard draft").Click += (_, _) => discardConfirmed = true;

            ButtonWithContent(window, "Cancel").Focus();
            PressKey(window, Key.Enter, PhysicalKey.Enter);

            Assert.False(discardConfirmed);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void Escape_cancels()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            bool discardConfirmed = false;
            ButtonWithContent(window, "Discard draft").Click += (_, _) => discardConfirmed = true;

            PressKey(window, Key.Escape, PhysicalKey.Escape);

            Assert.False(discardConfirmed);
            Assert.False(window.IsVisible);
        });
    }

    [Fact]
    public void The_confirmation_names_the_board_being_discarded_on_its_own_bold_line()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            var nameBlock = window.GetControl<TextBlock>("BoardNameText");

            Assert.Equal("Commodore 64 - 250469 (short board)", nameBlock.Text);
            Assert.Equal(FontWeight.Bold, nameBlock.FontWeight);
        });
    }

    [Fact]
    public void The_surrounding_text_explains_and_warns_before_and_after_the_name()
    {
        UiTest.Run(() =>
        {
            var window = BuildWindow();

            var textBlocks = window.GetVisualDescendants().OfType<TextBlock>().ToList();

            Assert.Contains(textBlocks, t => t.Text == "This permanently discards your local, unpublished edits to:");
            Assert.Contains(textBlocks, t => t.Text != null && t.Text.Contains("officially published data is not affected", StringComparison.Ordinal));
            Assert.Contains(textBlocks, t => t.Text == "This cannot be undone!");
        });
    }

    // ###########################################################################################
    // *** THE CONTRIBUTOR IS TOLD THE MAINTAINERS WILL KNOW (owner request, 2026-09-28). *** Only
    // when something sent from this draft is still being reviewed; in the words "My submissions"
    // gives its state; and never suggesting the submission is withdrawn (it is not).
    // ###########################################################################################
    [Fact]
    public void With_a_submission_still_in_review_the_dialog_says_the_maintainers_are_told()
    {
        UiTest.Run(() =>
        {
            var window = new DiscardDraftWindow();
            window.Initialize("Commodore 128 - 310378", [new SubmissionReceipt { SubmissionId = 9, UploadToken = "t", LastKnownState = "merged" }]);

            TextBlock notice = window.FindControl<TextBlock>("SubmissionNoticeText")!;

            Assert.True(notice.IsVisible);
            Assert.Contains(SubmissionReceiptPresenter.DescribeState("merged"), notice.Text, StringComparison.Ordinal);
            Assert.Contains("the maintainers are told", notice.Text, StringComparison.Ordinal);
            Assert.Contains("does not withdraw", notice.Text, StringComparison.Ordinal);
        });
    }

    [Fact]
    public void With_nothing_in_review_the_dialog_says_nothing_about_maintainers()
    {
        UiTest.Run(() =>
        {
            var window = new DiscardDraftWindow();
            window.Initialize("Commodore 128 - 310378", []);

            Assert.False(window.FindControl<TextBlock>("SubmissionNoticeText")!.IsVisible);
        });
    }
}
