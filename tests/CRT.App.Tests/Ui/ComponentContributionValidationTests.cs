using System.Collections.ObjectModel;
using System.Reflection;
using Avalonia.Controls;
using Avalonia.Media;
using Avalonia.Threading;
using Avalonia.VisualTree;
using CRT;

namespace ClassicRepairToolbox.Tests.Ui;

// Pre-submit validation of the "Component images" section, driven through the real window.
//
// ContributionPackagingTests already pins the rule itself (no file and no note -> NoFileOrNote,
// wrong type -> NotDisplayable, a note and no file -> fine). What is checked here is the half that rule alone cannot prove: that a marked
// row actually LOOKS marked. The mark travels model -> binding -> style, and every step of that is
// silently survivable if it breaks - a mistyped Classes binding or a style that loses to a local
// value simply shows nothing, and the user is told "you cannot save" with nothing to point at.
//
// The private members are reached by reflection, the same approach and the same reasoning as
// ExternalTargetLauncherTests: the logic is welded to a Window.
//
// Session 2c note: this window used to zip its edited rows and POST them to the contribution
// server (email required, a mandatory change comment, a locked "already sent" button). It now
// saves into the current board's local draft instead - see ComponentDraftWriter's own header.
// The email field and the mandatory-comment requirement are gone with it (a draft is not
// submitted to anyone, so neither has anything to validate yet); what a real SAVE actually writes
// is not exercised here, for the same reason TabSchematics.LabelEditor.cs's own
// ApplyLabelEditorChanges has no dedicated headless test - it needs a real board Excel file on
// disk, which this suite does not construct. Row/field validation, delete mode and the status
// area are unaffected by the save-target change and are still tested in full.
[Collection("HeadlessUi")]
public class ComponentContributionValidationTests
{
    // Case 2 (2026-10-02). The message says what would make the row acceptable - a file OR a note -
    // since "no file selected" alone sent the contributor looking for a file a note-only row never
    // needed.
    [Fact]
    public void A_component_image_row_with_neither_a_file_nor_a_note_is_marked_and_blocks_saving()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            var rows = GetComponentImageRows(window);

            // Exactly what "Add new component image" produces: a row with nothing filled in yet.
            var emptyRow = new ContributionComponentImageRow();
            rows.Add(emptyRow);

            var problem = Validate(window);

            Assert.NotNull(problem);
            Assert.True(emptyRow.HasFileError);
            Assert.Equal("No file or note", emptyRow.FileErrorText);

            // The status message has to name the row, or "something is wrong" is all the user gets.
            Assert.Equal("Component image #1 has neither a file nor a note - add one of them", ProblemMessage(problem));
        });
    }

    // Case 6. A row added here with only a note typed is the same kind of row as the C128's
    // "Pinout" rows that give a compatible part number - valid in the table and on the server, so
    // valid here too.
    [Fact]
    public void A_new_image_row_with_only_a_note_typed_is_accepted()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            var rows = GetComponentImageRows(window);

            var noteRow = new ContributionComponentImageRow { Name = "Pinout", Note = "Compatible part-number: 1N4148" };
            rows.Add(noteRow);

            Assert.Null(Validate(window));
            Assert.False(noteRow.HasFileError);
            Assert.Equal(string.Empty, noteRow.FileErrorText);
        });
    }

    // A row holding a file of the wrong type is a different problem and must say so differently -
    // this is the case that started the work: an .xlsx chosen as a component image.
    [Fact]
    public void A_component_image_row_holding_a_non_image_is_marked_as_the_wrong_type()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            var rows = GetComponentImageRows(window);

            var badTypeRow = new ContributionComponentImageRow
            {
                File = "baselines.xlsx",
                OriginalFilePath = "/pictures/baselines.xlsx"
            };
            rows.Add(badTypeRow);

            var problem = Validate(window);

            Assert.NotNull(problem);
            Assert.True(badTypeRow.HasFileError);
            Assert.Equal("Not an image the application can display", badTypeRow.FileErrorText);
            Assert.Contains(".png", ProblemMessage(problem));
        });
    }

    // Every bad row is marked in one pass, not just the one named in the status line - otherwise a
    // user with three broken rows fixes them one submission at a time. The first bad row is still
    // the one reported and scrolled to.
    [Fact]
    public void All_bad_rows_are_marked_while_the_good_ones_are_left_clean()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            var rows = GetComponentImageRows(window);

            var goodRow = new ContributionComponentImageRow { File = "pin1.png", OriginalFilePath = "/pictures/pin1.png" };
            var firstBadRow = new ContributionComponentImageRow();
            var secondBadRow = new ContributionComponentImageRow { File = "notes.txt", OriginalFilePath = "/pictures/notes.txt" };

            rows.Add(goodRow);
            rows.Add(firstBadRow);
            rows.Add(secondBadRow);

            var problem = Validate(window);

            Assert.NotNull(problem);
            Assert.False(goodRow.HasFileError);
            Assert.True(firstBadRow.HasFileError);
            Assert.True(secondBadRow.HasFileError);

            // #2 - the first BAD row, counted over all rows so the number matches what is on screen.
            Assert.Contains("Component image #2", ProblemMessage(problem));
        });
    }

    // A row that gets fixed must lose its mark, or the red border outlives the problem and the
    // editor lies about a row that is now perfectly fine.
    [Fact]
    public void Re_validating_a_row_that_has_been_given_an_image_clears_its_mark()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            var rows = GetComponentImageRows(window);

            var row = new ContributionComponentImageRow();
            rows.Add(row);

            Assert.NotNull(Validate(window));
            Assert.True(row.HasFileError);

            row.File = "pin1.png";
            row.OriginalFilePath = "/pictures/pin1.png";

            Assert.Null(Validate(window));
            Assert.False(row.HasFileError);
            Assert.Equal(string.Empty, row.FileErrorText);
        });
    }

    // With no component image rows at all there is nothing to complain about - contributing only a
    // description or a link must stay possible.
    [Fact]
    public void A_contribution_with_no_component_image_rows_has_nothing_to_report()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();

            Assert.Null(Validate(window));
        });
    }

    // The visual half of the feature: the model flag has to reach the row's Border as a style class
    // and repaint it. Anything less and validation blocks submission without showing where.
    [Fact]
    public void The_mark_turns_the_rows_border_red_on_screen()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            var rows = GetComponentImageRows(window);
            rows.Add(new ContributionComponentImageRow());

            // The section is collapsed by default, so its rows are not built until it is opened -
            // which is exactly why RevealComponentImageRow expands it before scrolling.
            window.Show();
            window.FindControl<Expander>("ComponentImagesExpander")!.IsExpanded = true;
            PumpLayout(window);

            var rowBorder = FindComponentImageRowBorder(window);
            Assert.NotNull(rowBorder);

            var normalBrush = rowBorder!.BorderBrush;
            Assert.DoesNotContain("HasFileError", rowBorder.Classes);

            Validate(window);
            PumpLayout(window);

            Assert.Contains("HasFileError", rowBorder.Classes);
            Assert.NotEqual(normalBrush, rowBorder.BorderBrush);
            Assert.Equal(3, rowBorder.BorderThickness.Left);

            window.Close();
        });
    }

    // ---------------------------------------------------------------- sending once

    // Saving a draft is never one-shot, unlike the old network submission: the window stays open so
    // more rows can be added and saved again, so the button always comes back regardless of outcome.
    [Fact]
    public void The_save_button_re_enables_after_either_outcome()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            var submitButton = window.FindControl<Button>("SubmitButton")!;

            Assert.True(submitButton.IsEnabled);

            ApplySubmissionOutcome(window, saveSucceeded: true);
            Assert.True(submitButton.IsEnabled);

            submitButton.IsEnabled = false;
            ApplySubmissionOutcome(window, saveSucceeded: false);
            Assert.True(submitButton.IsEnabled);
        });
    }

    // A submission refused before it is ever saved - here by an unusable component image row - is
    // not a successful save, so the button must still be pressable: the user is meant to fix the
    // form and press it again.
    [Fact]
    public void A_submission_refused_by_validation_leaves_the_send_button_alone()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            window.Show();

            var rows = GetComponentImageRows(window);
            rows.Add(new ContributionComponentImageRow());

            Submit(window);
            PumpLayout(window);

            Assert.True(window.FindControl<Button>("SubmitButton")!.IsEnabled);

            window.Close();
        });
    }

    // ---------------------------------------------------------------- the status area

    // The messages this window shows name the field at fault and say what happens next, so they run
    // long - and they used to sit in whatever width was left between the email box and the buttons,
    // where the end of the sentence was simply cut off. The area is now a row of its own that grows
    // with the message.
    [Fact]
    public void A_long_message_is_laid_out_over_as_many_lines_as_it_needs()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            window.Show();
            PumpLayout(window);

            var panel = window.FindControl<Border>("StatusPanel")!;
            var text = window.FindControl<TextBlock>("StatusTextBlock")!;

            Assert.False(panel.IsVisible);
            Assert.Equal(TextWrapping.Wrap, text.TextWrapping);

            ShowStatus(window, "Saved", false);
            PumpLayout(window);

            Assert.True(panel.IsVisible);
            double oneLineHeight = text.Bounds.Height;
            Assert.True(oneLineHeight > 0, "the status text was not laid out at all");

            ShowStatus(
                window,
                "Saved to your local draft. The new component [Dennis-5] now appears on this board, " +
                "and every other edit you have made to this board stays saved alongside it, " +
                "visible only to you until you submit your draft and it is reviewed and accepted.",
                false);
            PumpLayout(window);

            // It wrapped instead of running off the end...
            Assert.True(text.Bounds.Height > oneLineHeight, "the long message did not wrap onto further lines");

            // ...and the box grew with it, so every line of it is inside the box.
            Assert.True(panel.Bounds.Height >= text.Bounds.Height);
        });
    }

    // Which kind of message this is has to be readable from the box as well as the text, and the two
    // states must swap rather than pile up - a panel wearing both classes at once takes whichever
    // style happens to win.
    [Fact]
    public void The_area_carries_the_state_of_the_message_and_swaps_it()
    {
        UiTest.Run(() =>
        {
            var window = new ComponentContributionWindow();
            window.Show();

            var panel = window.FindControl<Border>("StatusPanel")!;
            var text = window.FindControl<TextBlock>("StatusTextBlock")!;

            Assert.True(window.TryFindResource("Text_Fail_Fg", window.ActualThemeVariant, out object? failBrush));
            Assert.True(window.TryFindResource("Text_Success_Fg", window.ActualThemeVariant, out object? successBrush));

            ShowStatus(window, "Something is wrong", true);
            PumpLayout(window);

            Assert.Contains("error", panel.Classes);
            Assert.DoesNotContain("success", panel.Classes);
            Assert.Equal(failBrush, panel.BorderBrush);
            Assert.Equal(failBrush, text.Foreground);

            ShowStatus(window, "All good", false);
            PumpLayout(window);

            Assert.Contains("success", panel.Classes);
            Assert.DoesNotContain("error", panel.Classes);
            Assert.Equal(successBrush, panel.BorderBrush);
            Assert.Equal(successBrush, text.Foreground);

            window.Close();
        });
    }

    // ---------------------------------------------------------------- helpers

    private static void ApplySubmissionOutcome(ComponentContributionWindow window, bool saveSucceeded)
    {
        var method = typeof(ComponentContributionWindow).GetMethod(
            "ApplySubmissionOutcome",
            BindingFlags.Instance | BindingFlags.NonPublic);

        method!.Invoke(window, new object?[] { saveSucceeded });
    }

    // ShowStatus posts its update, so every caller here drains the dispatcher afterwards.
    private static void ShowStatus(ComponentContributionWindow window, string message, bool isError)
    {
        var method = typeof(ComponentContributionWindow).GetMethod(
            "ShowStatus",
            BindingFlags.Instance | BindingFlags.NonPublic);

        method!.Invoke(window, new object?[] { message, isError });
    }


    private static ObservableCollection<ContributionComponentImageRow> GetComponentImageRows(ComponentContributionWindow window)
    {
        var field = typeof(ComponentContributionWindow).GetField(
            "thisComponentImageRows",
            BindingFlags.Instance | BindingFlags.NonPublic);

        return (ObservableCollection<ContributionComponentImageRow>)field!.GetValue(window)!;
    }

    // Returns the (Row, Message) tuple ValidateComponentImageRows produced, or null when it found
    // nothing wrong. Boxed as object because the tuple type itself is private to the window.
    private static object? Validate(ComponentContributionWindow window)
    {
        var method = typeof(ComponentContributionWindow).GetMethod(
            "ValidateComponentImageRows",
            BindingFlags.Instance | BindingFlags.NonPublic);

        return method!.Invoke(window, null);
    }

    private static string ProblemMessage(object? problem)
    {
        // The nullable tuple comes back boxed as its underlying (Row, Message) value.
        return (string)problem!.GetType().GetField("Item2")!.GetValue(problem)!;
    }

    // Presses "Save to draft" through the real handler. Every test using this must be sure
    // validation rejects the form, or the handler would go on to try a real draft save - which
    // needs a board Excel file on disk this suite does not construct (see the class header).
    private static void Submit(ComponentContributionWindow window)
    {
        var method = typeof(ComponentContributionWindow).GetMethod(
            "OnSubmitClick",
            BindingFlags.Instance | BindingFlags.NonPublic);

        method!.Invoke(window, new object?[] { null, new Avalonia.Interactivity.RoutedEventArgs() });
    }

    private static Border? FindComponentImageRowBorder(ComponentContributionWindow window)
    {
        var itemsControl = window.FindControl<ItemsControl>("ComponentImageRowsItemsControl")!;

        return itemsControl.GetVisualDescendants()
            .OfType<Border>()
            .FirstOrDefault(border => border.Classes.Contains("ComponentImageRow"));
    }

    // Headless windows do not lay out on their own, and an ItemsControl builds no containers until
    // it has been measured - so the dispatcher is drained and a layout pass forced by hand.
    private static void PumpLayout(Window window)
    {
        Dispatcher.UIThread.RunJobs();
        window.Measure(window.ClientSize);
        window.Arrange(new Avalonia.Rect(window.ClientSize));
        Dispatcher.UIThread.RunJobs();
    }
}
