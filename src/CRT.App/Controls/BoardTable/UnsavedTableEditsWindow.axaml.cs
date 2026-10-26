using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;

namespace CRT
{
    // What the contributor chose in UnsavedTableEditsWindow. ShowDialog returns null when the
    // window is closed from its title bar, which callers treat as Cancel.
    public enum UnsavedTableEditsChoice
    {
        Cancel,
        Discard,
        Save
    }

    // Which question the prompt asks - see UnsavedTableEditsWindow's header.
    public enum UnsavedTableEditsPrompt
    {
        // Closing the MAINTAINER TAB's table on a submission (2026-09-25): Save, Discard or
        // Cancel - the edits go into the submission.
        LeavingSubmission,

        // Leaving the table of a BOARD on the Maintainer tab's Boards screen (2026-10-03): Save,
        // Discard or Cancel - Save asks for a reason and publishes the change straight to BETA.
        LeavingBoard,

        // Leaving the Contribute tab's table, or acting on its draft's file, after the table's save was
        // REFUSED - the draft's file changed outside the table: Discard or Cancel, and the reason.
        // A draft saves itself (CRT 4.0 step 1), so there is no Save to offer - it was just tried.
        DraftChangedOnDisk,

        // The same after the table's save FAILED - the draft's file could not be written (another
        // program holding it, say): Discard or Cancel, and the reason.
        DraftNotWritten,

        // Leaving a MAINTAINER TAB table while the tab says "CRT has to be updated" (code review,
        // 2026-10-09): Discard or Cancel. Its Save sends to the server, which turns this CRT away -
        // offered, it was refused out of sight under the cover and the quit silently cancelled.
        LeavingUpdateRequired
    }

    // ###########################################################################################
    // Asked before a table's changes that are not saved would be lost - leaving a Maintainer tab
    // table (Save, Discard or Cancel: its saves publish or amend, so they stay explicit), and
    // leaving the Contribute tab's table, or acting on its draft's file, when the table could not save
    // (2026-09-24; since CRT 4.0 step 1, 2026-10-25, a draft saves itself, so this is asked only
    // after its save was refused or failed - Discard or Cancel, never a Save that was just tried and
    // would loop: Save, refused, the table stays open, Close asks again).
    //
    // The other editors are no longer held back by the Contribute tab's table (CRT 4.0 step 1,
    // 2026-10-25): every editor changes the one draft held in memory, so their saves land on top of
    // the table's changes, and the "SavingElsewhere" notices that sent the contributor to the Drafts
    // tab first are gone.
    //
    // Enter and Escape both CANCEL, on the Tunnel route - the same rule and the same reason as
    // DiscardDraftWindow: "Discard edits" is permanent, and a focused Button consumes Enter on the
    // bubbling route, so a reflexive keypress must never be the thing that throws work away.
    // ###########################################################################################
    public partial class UnsavedTableEditsWindow : Window
    {
        public UnsavedTableEditsWindow()
        {
            this.InitializeComponent();

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        public void Initialize(UnsavedTableEditsPrompt prompt)
        {
            this.SaveButton.IsVisible = prompt is UnsavedTableEditsPrompt.LeavingSubmission or UnsavedTableEditsPrompt.LeavingBoard;

            // A draft's - the table's or the component pane's (CRT 4.0 step 2): said of the draft.
            if (prompt is UnsavedTableEditsPrompt.DraftChangedOnDisk or UnsavedTableEditsPrompt.DraftNotWritten)
            {
                this.Title = "Changes not saved";
                this.HeadingText.Text = "Changes not saved";
            }

            this.MessageText.Text = prompt switch
            {
                UnsavedTableEditsPrompt.LeavingSubmission =>
                    "The table has changes that are not saved yet. Save them into the submission, or discard them?",
                UnsavedTableEditsPrompt.LeavingBoard =>
                    "The table has changes that are not published yet. Save them - you are asked for a reason, and they go straight to BETA - or discard them?",
                UnsavedTableEditsPrompt.DraftNotWritten =>
                    "Your latest changes to the draft could not be saved: its file could not be written - " +
                    "another program may have it open. Cancel to stay (close the file there, then " +
                    "press Try again) - or discard the changes that could not be saved.",
                UnsavedTableEditsPrompt.LeavingUpdateRequired =>
                    "The table has changes that are not sent yet, and they cannot be sent until CRT is updated - " +
                    "the server turns this version of CRT away. Discard them, or cancel to keep CRT open and " +
                    "copy what you need first.",
                _ =>
                    "Your latest changes to the draft were not saved: its file was changed somewhere else after it " +
                    "was read, and CRT does not save over that unasked. Cancel to stay - the table view " +
                    "offers to use the saved version or to keep yours - or discard the changes that could not be saved.",
            };
        }

        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape || e.Key == Key.Enter)
            {
                this.Close(UnsavedTableEditsChoice.Cancel);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(UnsavedTableEditsChoice.Cancel);

        private void OnDiscardClick(object? sender, RoutedEventArgs e) => this.Close(UnsavedTableEditsChoice.Discard);

        private void OnSaveClick(object? sender, RoutedEventArgs e) => this.Close(UnsavedTableEditsChoice.Save);
    }
}
