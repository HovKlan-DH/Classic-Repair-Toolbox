using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE SUBMISSION'S THREE VIEWS (owner request, 2026-09-30) - Board data, Files, Contributor,
    // chosen by the buttons above the submission. What each is and why Board data comes first is
    // SubmissionViews' header; this part owns which one is shown.
    //
    //   - A NEWLY CHOSEN submission opens on Board data, with nothing of the previous one in the
    //     other two. The table stays "the default first view" (owner request, 2026-09-26).
    //   - The SAME submission keeps its view through the queue's own minute check and a re-read
    //     detail - the check "should not conflict if the maintainer is in the middle of something".
    //   - SWITCHING HIDES, IT NEVER CLOSES: the table - unsaved changes and all - is exactly where it
    //     was on coming back to Board data, so nothing is asked. The views are siblings whose
    //     visibility changes; none is taken out of the tree.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private SubmissionView thisSubmissionView = SubmissionViews.Opening;

        // The view on show, for tests.
        internal SubmissionView ShownSubmissionView => this.thisSubmissionView;

        private static readonly (SubmissionView View, string Button, string Panel)[] SubmissionViewParts =
        [
            (SubmissionView.BoardData, "BoardDataViewButton", "BoardDataView"),
            (SubmissionView.Files, "FilesViewButton", "FilesView"),
            (SubmissionView.Contributor, "ContributorViewButton", "ContributorView")
        ];

        private async void OnSubmissionViewClick(object? sender, RoutedEventArgs e)
        {
            if (sender is not Button button)
                return;

            foreach ((SubmissionView view, string name, _) in TabMaintainer.SubmissionViewParts)
            {
                if (string.Equals(button.Name, name, StringComparison.Ordinal))
                {
                    await this.ShowSubmissionViewAsync(view);
                    return;
                }
            }
        }

        // Shows one view; the Files view reads its tree the first time it is shown for a submission.
        internal async Task ShowSubmissionViewAsync(SubmissionView view)
        {
            this.thisSubmissionView = view;
            this.ApplySubmissionView();

            if (view == SubmissionView.Files)
                await this.LoadFilesViewAsync();
        }

        private void ApplySubmissionView()
        {
            foreach ((SubmissionView view, string button, string panel) in TabMaintainer.SubmissionViewParts)
            {
                bool shown = view == this.thisSubmissionView;

                // No tooltips, as on the four screen buttons (owner request, 2026-09-30).
                if (this.FindControl<Button>(button) is Button viewButton)
                    viewButton.Classes.Set("Selected", shown);

                this.SetShown(panel, shown);
            }
        }

        // ###########################################################################################
        // Another submission chosen - or none: back to Board data, and the other two views emptied,
        // so nothing of the previous submission's files or contributor can be read as this one's.
        // ###########################################################################################
        private void ResetSubmissionViews()
        {
            this.thisSubmissionView = SubmissionViews.Opening;
            this.ClearFilesView();
            this.ShowContributorHistory(null);
            this.ApplySubmissionView();
        }
    }
}
