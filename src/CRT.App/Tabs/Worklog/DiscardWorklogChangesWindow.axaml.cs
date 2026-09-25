using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // "You are about to discard a worklog you have not saved yet" - shown when Cancel, Escape or
    // the title-bar close would throw away work typed into a NEW worklog.
    //
    // WHY THIS EXISTS. Sub-list changes behave differently on a draft than on a saved worklog, and
    // the UI gives no sign of it. On a SAVED worklog, adding a comment, a photo, a link or a
    // work-done row writes to disk immediately (PersistEntrySilently), so Escape can only ever cost
    // the pending Title/Description edits. On a NEW one, nothing is written until Save: the same
    // actions only update an in-memory record, and Escape discards the lot - the typed fields, every
    // comment, and the attachment bytes with them (DiscardDraftAttachments).
    //
    // Reported as real data loss, and the reason given is exactly that mismatch: after adding a
    // comment the way you would on a saved worklog, it FEELS saved, so Escape feels like closing a
    // finished thing rather than abandoning an unfinished one.
    //
    // Deliberately NOT shown for a saved worklog. There, Escape is a cheap, well-understood way to
    // back out of some half-typed edits, and a prompt on every one of those would train the user to
    // dismiss it - which would then get dismissed on the one that mattered.
    //
    // Returns true via ShowDialog when the user confirms the discard, and null/false when they
    // choose to keep editing.
    // ###########################################################################################
    public partial class DiscardWorklogChangesWindow : Window
    {
        public DiscardWorklogChangesWindow()
        {
            this.InitializeComponent();

            // Tunnel, NOT a plain KeyDown subscription - the same reasoning as DeleteWorklogWindow
            // and DeleteWorkbookWindow. A focused Button handles Enter itself and marks the event
            // handled, so a bubbling handler never runs, and "Discard worklog" having focus would
            // make Enter destroy the very work this dialog exists to protect.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Escape AND Enter both KEEP EDITING - the same inversion DeleteWorklogWindow makes, and for
        // the same reason: this dialog's destructive option is the one that loses data, so a
        // reflexive keypress must land on the safe side.
        //
        // Escape carries an extra risk here that the delete dialogs do not have: the user got to
        // this dialog BY pressing Escape, so a second Escape - a double-tap, or an impatient repeat
        // when the first seemed not to work - is genuinely likely. Making it discard would turn the
        // guard into a two-keystroke version of the same accident.
        //
        // Runs on the Tunnel route (see the constructor) so it beats a focused Button's own Enter.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape || e.Key == Key.Enter)
            {
                this.Close(false);
                e.Handled = true;
            }
        }

        private void OnKeepEditingClick(object? sender, RoutedEventArgs e) => this.Close(false);

        private void OnDiscardClick(object? sender, RoutedEventArgs e) => this.Close(true);
    }
}
