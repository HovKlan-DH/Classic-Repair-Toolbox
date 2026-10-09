using Avalonia.Controls;
using System.Threading.Tasks;
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
        // Closing the table, another action on the draft, quitting: Save, Discard or Cancel.
        Leaving,

        // "Reload": Discard or Cancel.
        Reloading,

        // Closing the MAINTAINER TAB's table on a submission (2026-09-25): Save, Discard or
        // Cancel - the same choice as Leaving, but the edits go into the submission, not a draft.
        LeavingSubmission,

        // Leaving the table of a BOARD on the Maintainer tab's Boards screen (2026-10-03): Save,
        // Discard or Cancel - Save asks for a reason and publishes the change straight to BETA.
        LeavingBoard,

        // Leaving, but the draft file changed after the table was opened, so a save is REFUSED:
        // Discard or Cancel, and the reason said.
        DraftChangedOnDisk,

        // Leaving, but the draft is held open in Excel, so a save would FAIL: Discard or Cancel,
        // and the reason said - offering Save there would loop as DraftChangedOnDisk once did.
        DraftOpenElsewhere,

        // ANOTHER editor - the Contribute tab's component editor, the label editor - wants to save
        // into a draft whose table has unsaved edits. Nothing can be decided from here, so it is a
        // notice with Cancel alone: see ShowSavingBlockedAsync.
        SavingElsewhere,

        // Leaving a MAINTAINER TAB table while the tab says "CRT has to be updated" (code review,
        // 2026-10-09): Discard or Cancel. Its Save sends to the server, which turns this CRT away -
        // offered, it was refused out of sight under the cover and the quit silently cancelled.
        LeavingUpdateRequired,

        // SavingElsewhere while the Drafts tab says "CRT has to be updated" (code review,
        // 2026-10-09): Cancel alone, as SavingElsewhere - but the Drafts tab cannot be reached to
        // settle the table there, so it says what does: updating CRT, or quitting it, asks first.
        SavingElsewhereUpdateRequired
    }

    // ###########################################################################################
    // Asked before the Drafts tab's table editor would lose edits that are not saved yet - closing
    // the table, another action on the same draft, reloading it, or quitting the application
    // (2026-09-24).
    //
    // Three wordings from one window, chosen by Initialize: leaving the table offers Save, Discard
    // and Cancel; reloading offers only Discard and Cancel, because the usual reason to reload is
    // that saving has just been REFUSED (the file changed underneath the table), so a Save button
    // there would promise something that cannot happen.
    //
    // *** AND LEAVING A TABLE WHOSE DRAFT CHANGED ON DISK OFFERS NO SAVE EITHER (reported,
    // 2026-09-24). *** "Save to draft" in the Contribute tab writes the same file, and then the
    // table's own save is refused. Offering Save there made "Close table" a loop: Save, refused,
    // the table stays open, Close asks again. The prompt now says why and offers Discard or Cancel.
    //
    // *** AND THE OTHER EDITORS DO NOT WRITE UNDER A TABLE WITH UNSAVED EDITS AT ALL (owner
    // request, 2026-09-24). *** Their "Save to draft" is held back with the SavingElsewhere notice
    // instead, sending the contributor to the Drafts tab to save or discard the table first. The
    // project owner chose that over offering to save the table from here: from another tab you may not
    // remember what you did in the table, and saving it blind is not a decision to make there.
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
            bool notice = prompt is UnsavedTableEditsPrompt.SavingElsewhere or UnsavedTableEditsPrompt.SavingElsewhereUpdateRequired;

            this.SaveButton.IsVisible = prompt is UnsavedTableEditsPrompt.Leaving or UnsavedTableEditsPrompt.LeavingSubmission or UnsavedTableEditsPrompt.LeavingBoard;
            this.DiscardButton.IsVisible = !notice;

            if (notice)
            {
                this.Title = "Unsaved table edits in the Drafts tab";
                this.HeadingText.Text = "Unsaved table edits in the Drafts tab";
            }

            this.MessageText.Text = prompt switch
            {
                UnsavedTableEditsPrompt.Leaving =>
                    "The table has edits that are not saved yet. Save them into the draft first, or discard them?",
                UnsavedTableEditsPrompt.LeavingSubmission =>
                    "The table has changes that are not saved yet. Save them into the submission, or discard them?",
                UnsavedTableEditsPrompt.LeavingBoard =>
                    "The table has changes that are not published yet. Save them - you are asked for a reason, and they go straight to BETA - or discard them?",
                UnsavedTableEditsPrompt.Reloading =>
                    "Reloading reads the draft again from its file. The edits in the table that are not saved yet will be lost.",
                UnsavedTableEditsPrompt.DraftOpenElsewhere =>
                    "The draft is open in Excel (or another spreadsheet program), so the table's edits cannot " +
                    "be saved into it until it is closed there. Cancel, close it in Excel and try again - or " +
                    "discard the table's edits.",
                UnsavedTableEditsPrompt.SavingElsewhere =>
                    "The table in the Drafts tab has unsaved edits for this same board. They need to be dealt " +
                    "with before this can be saved. Go to the Drafts tab, save or discard the table's edits, " +
                    "then come back here and save this again if it is still needed.",
                UnsavedTableEditsPrompt.SavingElsewhereUpdateRequired =>
                    "The table in the Drafts tab has unsaved edits for this same board, and the Drafts tab " +
                    "cannot be used until CRT is updated. Updating CRT - or quitting it - asks first whether " +
                    "to save the table's edits. After that, save this again if it is still needed.",
                UnsavedTableEditsPrompt.LeavingUpdateRequired =>
                    "The table has changes that are not sent yet, and they cannot be sent until CRT is updated - " +
                    "the server turns this version of CRT away. Discard them, or cancel to keep CRT open and " +
                    "copy what you need first.",
                _ =>
                    "The draft was changed outside this table after it was opened - by \"Save to draft\" in the " +
                    "Contribute tab, for example, or in Excel - so the edits in the table can no longer be saved " +
                    "into it. Discard them, or cancel to keep the table open and copy what you need first.",
            };
        }

        // ###########################################################################################
        // The notice another editor shows instead of saving while the Drafts tab's table holds
        // unsaved edits for the same board. Nothing is saved; the editor stays open as it was.
        // ###########################################################################################
        public static async Task ShowSavingBlockedAsync(Window owner, UnsavedTableEditsPrompt prompt = UnsavedTableEditsPrompt.SavingElsewhere)
        {
            var window = new UnsavedTableEditsWindow();
            window.Initialize(prompt);

            await window.ShowDialog<UnsavedTableEditsChoice?>(owner);
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
