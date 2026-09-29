using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // Confirmation modal for "Discard draft" on the Drafts tab - discarding a draft removes every
    // local, unpublished edit to a system permanently (see DraftManager.DiscardDraft), the same
    // shape and the same reasoning as DeleteWorkbookWindow. Returns true via ShowDialog when the
    // user confirms, or null when cancelled.
    // ###########################################################################################
    public partial class DiscardDraftWindow : Window
    {
        public DiscardDraftWindow()
        {
            this.InitializeComponent();

            // Tunnel, NOT a plain KeyDown subscription - see DeleteWorkbookWindow's own comment on
            // this exact trap: a focused Button handles Enter itself and marks the event handled,
            // so a bubbling handler never runs, which would let a reflexive Enter after clicking
            // (or tabbing to) "Discard draft" actually discard it.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // Names the system being discarded, on its own bold line, so a user with several boards'
        // worth of drafts cannot mistake which one they are about to lose.
        //
        // `unfinished` are this board's submissions still with the maintainers
        // (DraftDiscardContract.WhichToReport). When there are any, the window says what discarding
        // means for them - see NoticeFor.
        // ###########################################################################################
        public void Initialize(string systemDisplayName, IReadOnlyList<SubmissionReceipt>? unfinished = null)
        {
            this.SystemNameText.Text = systemDisplayName;

            string? notice = DiscardDraftWindow.NoticeFor(unfinished);
            this.SubmissionNoticeText.Text = notice ?? string.Empty;
            this.SubmissionNoticeText.IsVisible = notice is not null;
        }

        // ###########################################################################################
        // *** THE CONTRIBUTOR IS TOLD THE MAINTAINERS WILL KNOW (owner request, 2026-09-28). *** A
        // discard is reported to the server so a maintainer does not publish work its author has
        // thrown away without asking them first - and the person pressing the button should know
        // that before they do, in words that do not suggest their submission is withdrawn (it is
        // not). Null when nothing sent from this draft is still being reviewed.
        // ###########################################################################################
        internal static string? NoticeFor(IReadOnlyList<SubmissionReceipt>? unfinished)
        {
            if (unfinished is null || unfinished.Count == 0)
                return null;

            // The newest, in the words the Drafts tab's badge and "My submissions" use for it.
            SubmissionReceipt latest = unfinished.OrderByDescending(receipt => receipt.SentUtc).First();
            string state = SubmissionReceiptPresenter.DescribeState(latest.LastKnownState);

            return $"You have sent this board for review, and that is not finished yet ({state}). " +
                "Discarding your draft does not withdraw what you sent - but the maintainers are told that " +
                "you discarded it, so they can check with you before it is published.";
        }

        // ###########################################################################################
        // Escape AND Enter both cancel - deliberately unlike every other modal in the app, because
        // this one's "submit" is a permanent discard. See DeleteWorkbookWindow's own comment: this
        // must run on the Tunnel route to beat a focused Button's own Enter handling.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape || e.Key == Key.Enter)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            this.Close(null);
        }

        private void OnDiscardClick(object? sender, RoutedEventArgs e)
        {
            this.Close(true);
        }
    }
}
