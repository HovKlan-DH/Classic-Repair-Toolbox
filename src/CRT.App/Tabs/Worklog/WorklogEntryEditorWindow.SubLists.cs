using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Handlers.DataHandling;
using System;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // The Links, Comments and Work done sub-lists: refresh, add/edit/delete, and (for Comments and
    // Work done) the newest-first/oldest-first sort toggle. See WorklogEntryEditorWindow.axaml.cs
    // for the file map of the whole partial class.
    // ###########################################################################################
    public partial class WorklogEntryEditorWindow
    {
        // Newest-first is the default sort for both lists - persisted globally via UserSettings so it
        // carries over between entries and app restarts, rather than resetting every time this window
        // is opened.
        private bool thisCommentsSortNewestFirst = UserSettings.WorklogCommentsSortNewestFirst;
        private bool thisWorkDoneSortNewestFirst = UserSettings.WorklogWorkDoneSortNewestFirst;

        // ###########################################################################################
        // Links of interest
        // ###########################################################################################
        private void RefreshLinkRows()
        {
            this.thisLinkRows.Clear();
            foreach (var link in this.thisEntry.Links)
            {
                this.thisLinkRows.Add(new WorklogLinkRow { Id = link.Id, Headline = link.Headline, Url = link.Url });
            }
            this.EditorNoLinksText.IsVisible = this.thisLinkRows.Count == 0 && this.IsListSectionExpanded("EditorLinksHeader");
            this.EditorLinksCountText.Text = FormatItemCount(this.thisLinkRows.Count, "link", "links");
        }

        private async void OnAddLinkClick(object? sender, RoutedEventArgs e)
        {
            // Guarded as a whole (code review, 2026-10-13): an "async void" handler has no caller,
            // so an exception after an await reached App's handler and ended CRT. Logged, and CRT goes on.
            try
            {
                var dialog = new WorklogAddLinkWindow();
                var result = await dialog.ShowDialog<(string Headline, string Url)?>(this);
                if (result == null)
                    return;

                int nextId = RecordIds.Next(this.thisEntry.Links.Select(l => l.Id));
                this.thisEntry.Links.Add(new WorklogLinkRecord { Id = nextId, Headline = result.Value.Headline, Url = result.Value.Url });

                // A new row landing in a folded list would look like nothing happened, so adding always
                // opens the section it went into.
                this.EnsureListSectionExpanded("EditorLinksHeader");
                this.RefreshLinkRows();
                this.PersistEntrySilently();
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnAddLinkClick failed - [{ex}]");
            }
        }

        private void OnDeleteLinkClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: int id })
            {
                this.thisEntry.Links.RemoveAll(l => l.Id == id);
                this.RefreshLinkRows();
                this.PersistEntrySilently();
            }
        }

        // ###########################################################################################
        // Clicking anywhere on a link row (other than its Edit/Delete icons, which handle their own
        // Click and so never reach here) opens the link in the system browser, via the same sanctioned
        // launcher the rest of the app uses for external URLs.
        // ###########################################################################################
        private void OnLinkRowPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            if (sender is Border { Tag: string url } && !string.IsNullOrWhiteSpace(url))
            {
                this.OpenLink(url);
            }
        }

        // A link as typed, opened by LinkTarget's rule - "www.c64-wiki.com" too, as the exported PDF opens it (code review,
        // 2026-11-08).
        internal void OpenLink(string url) => ExternalTargetLauncher.OpenLinkOrSay(url, this.OpenLinkOverrideForTests);

        // Handed the address that would be opened, a test stands in for opening it. Null in the app.
        internal Action<string>? OpenLinkOverrideForTests { get; set; }

        private async void OnEditLinkClick(object? sender, RoutedEventArgs e)
        {
            // Guarded as a whole (code review, 2026-10-13): an "async void" handler has no caller,
            // so an exception after an await reached App's handler and ended CRT. Logged, and CRT goes on.
            try
            {
                if (sender is not Button { Tag: int id })
                    return;

                var link = this.thisEntry.Links.FirstOrDefault(l => l.Id == id);
                if (link == null)
                    return;

                var dialog = new WorklogAddLinkWindow();
                dialog.InitializeForEdit(link.Headline, link.Url);
                var result = await dialog.ShowDialog<(string Headline, string Url)?>(this);
                if (result == null)
                    return;

                link.Headline = result.Value.Headline;
                link.Url = result.Value.Url;
                this.RefreshLinkRows();
                this.PersistEntrySilently();
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnEditLinkClick failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // Comments
        // ###########################################################################################
        private void RefreshCommentRows()
        {
            this.thisCommentRows.Clear();
            var orderedComments = this.thisCommentsSortNewestFirst
                ? this.thisEntry.Comments.OrderByDescending(c => c.Date)
                : this.thisEntry.Comments.OrderBy(c => c.Date);
            foreach (var comment in orderedComments)
            {
                this.thisCommentRows.Add(new WorklogCommentRow
                {
                    Id = comment.Id,
                    Text = comment.Text,
                    DateText = comment.Date.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture)
                });
            }

            this.EditorNoCommentsText.IsVisible = this.thisCommentRows.Count == 0 && this.IsListSectionExpanded("EditorCommentsHeader");
            this.EditorCommentsCountText.Text = FormatItemCount(this.thisCommentRows.Count, "comment", "comments");

            this.UpdateCommentsSortIconVisuals();
        }

        private void OnCommentsSortNewestFirstClick(object? sender, RoutedEventArgs e)
        {
            this.thisCommentsSortNewestFirst = true;
            UserSettings.WorklogCommentsSortNewestFirst = true;
            this.RefreshCommentRows();
        }

        private void OnCommentsSortOldestFirstClick(object? sender, RoutedEventArgs e)
        {
            this.thisCommentsSortNewestFirst = false;
            UserSettings.WorklogCommentsSortNewestFirst = false;
            this.RefreshCommentRows();
        }

        private void UpdateCommentsSortIconVisuals()
        {
            var activeBrush = this.ResolveThemeBrush("Main_TabUnderline_Selected", new SolidColorBrush(Colors.IndianRed));
            var inactiveBrush = this.ResolveThemeBrush("Schematics_Panels_Fg", Brushes.Black);

            this.CommentsSortNewestFirstIcon.Foreground = this.thisCommentsSortNewestFirst ? activeBrush : inactiveBrush;
            this.CommentsSortOldestFirstIcon.Foreground = !this.thisCommentsSortNewestFirst ? activeBrush : inactiveBrush;
        }

        private async void OnAddCommentClick(object? sender, RoutedEventArgs e)
        {
            // Guarded as a whole (code review, 2026-10-13): an "async void" handler has no caller,
            // so an exception after an await reached App's handler and ended CRT. Logged, and CRT goes on.
            try
            {
                await this.AddCommentAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnAddCommentClick failed - [{ex}]");
            }
        }

        internal async Task AddCommentAsync()
        {
            if (await this.AskCommentAsync(null) is not { } text)
                return;

            int nextId = RecordIds.Next(this.thisEntry.Comments.Select(c => c.Id));
            this.thisEntry.Comments.Add(new WorklogCommentRecord { Id = nextId, Text = text, Date = DateTime.Now });

            this.EnsureListSectionExpanded("EditorCommentsHeader");
            this.RefreshCommentRows();
            this.PersistEntrySilently();
        }

        // ###########################################################################################
        // The "Add comment" window - on `text` to change it - and what it closes with: the comment's markup, trimmed by
        // RichNote.TrimMarkup, never Trim() (code review, 2026-11-08: a second Trim() here took a first line's bullet
        // indent off again, after the window had kept it) - or null, cancelled or nothing typed.
        // ###########################################################################################
        private async Task<string?> AskCommentAsync(string? text)
        {
            string? answered;

            if (this.AskCommentOverrideForTests is { } ask)
            {
                answered = await ask(text);
            }
            else
            {
                var dialog = new WorklogAddCommentWindow();

                if (text is not null)
                    dialog.InitializeForEdit(text);

                answered = await dialog.ShowDialog<string?>(this);
            }

            return RichNote.TrimMarkup(answered) is { Length: > 0 } trimmed ? trimmed : null;
        }

        // Stands in for the "Add comment" window in a test, which cannot show a dialog - handed the comment being
        // changed (null: a new one), it answers as the window would. Null in the app.
        internal Func<string?, Task<string?>>? AskCommentOverrideForTests { get; set; }

        private void OnDeleteCommentClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: int id })
            {
                this.thisEntry.Comments.RemoveAll(c => c.Id == id);
                this.RefreshCommentRows();
                this.PersistEntrySilently();
            }
        }

        // ###########################################################################################
        // Clicking anywhere on a comment row (other than its Delete icon, which handles its own Click
        // and so never reaches here) reopens the same modal Add-comment uses, pre-filled for editing.
        // ###########################################################################################
        private async void OnCommentRowPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            // Guarded as a whole (code review, 2026-10-13): an "async void" handler has no caller,
            // so an exception after an await reached App's handler and ended CRT. Logged, and CRT goes on.
            try
            {
                if (sender is Border { Tag: int id })
                    await this.EditCommentAsync(id);
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnCommentRowPointerPressed failed - [{ex}]");
            }
        }

        internal async Task EditCommentAsync(int commentId)
        {
            var comment = this.thisEntry.Comments.FirstOrDefault(c => c.Id == commentId);
            if (comment == null)
                return;

            if (await this.AskCommentAsync(comment.Text) is not { } text)
                return;

            comment.Text = text;
            this.RefreshCommentRows();
            this.PersistEntrySilently();
        }

        // ###########################################################################################
        // Work done
        // ###########################################################################################
        private void RefreshWorkDoneRows()
        {
            this.thisWorkDoneRows.Clear();
            var orderedWork = this.thisWorkDoneSortNewestFirst
                ? this.thisEntry.WorkDoneItems.OrderByDescending(w => w.Date)
                : this.thisEntry.WorkDoneItems.OrderBy(w => w.Date);
            foreach (var work in orderedWork)
            {
                this.thisWorkDoneRows.Add(new WorklogWorkDoneRow
                {
                    Id = work.Id,
                    Text = work.Text,
                    DateText = work.Date.ToString("yyyy-MM-dd HH:mm", CultureInfo.InvariantCulture),
                    // The cost through WorklogCurrency.FormatCost rather than a bare "{Cost:0.##}",
                    // and the time as WORDS through WorklogDurationFormatter, so a row reads
                    // "1 hour and 30 minutes · 430 DKK" - the same shape the Workbooks tab's entry
                    // card, the summary strip and the exported PDF all print. The decimal hours
                    // this is stored as appear in exactly one place, the NumericUpDown that types
                    // them. A row with no time logged shows the cost alone.
                    SummaryText = JoinSummaryParts(
                        WorklogDurationFormatter.Format(work.HoursSpent),
                        WorklogCurrency.FormatCostOrEmpty(
                            work.Cost,
                            WorklogCurrency.ResolveRecordedCode(work.CurrencyCode, UserSettings.WorklogCurrencyCode)))
                });
            }

            // The SAME sums the Workbooks tab's entry-detail card shows - see
            // WorklogEntryScope.GetWorkDoneTotals. Both were written out separately, with a comment
            // on the other one saying its formatting had to match this line by hand.
            var (totalHours, totalCost) = WorklogEntryScope.GetWorkDoneTotals(this.thisEntry);
            this.EditorWorkDoneCountText.Text = this.thisWorkDoneRows.Count == 0
                ? "none"
                : JoinSummaryParts(
                    FormatItemCount(this.thisWorkDoneRows.Count, "entry", "entries"),
                    WorklogDurationFormatter.Format(totalHours),
                    WorklogCurrency.FormatCostOrEmpty(totalCost, UserSettings.WorklogCurrencyCode));

            this.EditorNoWorkDoneText.IsVisible = this.thisWorkDoneRows.Count == 0 && this.IsListSectionExpanded("EditorWorkDoneHeader");

            this.UpdateWorkDoneSortIconVisuals();
        }

        // Joins the parts of a work-done summary with the app's " · ", DROPPING any that came back
        // empty. That is what lets a zero duration and a zero cost each contribute nothing at all
        // (WorklogDurationFormatter.Format and WorklogCurrency.FormatCostOrEmpty both return "" for
        // one) without leaving a stray separator behind, which a plain string interpolation of the
        // three would. A line with neither reads as just its count.
        private static string JoinSummaryParts(params string[] parts) =>
            string.Join(" · ", parts.Where(p => !string.IsNullOrWhiteSpace(p)));

        private void OnWorkDoneSortNewestFirstClick(object? sender, RoutedEventArgs e)
        {
            this.thisWorkDoneSortNewestFirst = true;
            UserSettings.WorklogWorkDoneSortNewestFirst = true;
            this.RefreshWorkDoneRows();
        }

        private void OnWorkDoneSortOldestFirstClick(object? sender, RoutedEventArgs e)
        {
            this.thisWorkDoneSortNewestFirst = false;
            UserSettings.WorklogWorkDoneSortNewestFirst = false;
            this.RefreshWorkDoneRows();
        }

        private void UpdateWorkDoneSortIconVisuals()
        {
            var activeBrush = this.ResolveThemeBrush("Main_TabUnderline_Selected", new SolidColorBrush(Colors.IndianRed));
            var inactiveBrush = this.ResolveThemeBrush("Schematics_Panels_Fg", Brushes.Black);

            this.WorkDoneSortNewestFirstIcon.Foreground = this.thisWorkDoneSortNewestFirst ? activeBrush : inactiveBrush;
            this.WorkDoneSortOldestFirstIcon.Foreground = !this.thisWorkDoneSortNewestFirst ? activeBrush : inactiveBrush;
        }

        private async void OnAddWorkDoneClick(object? sender, RoutedEventArgs e)
        {
            // Guarded as a whole (code review, 2026-10-13): an "async void" handler has no caller,
            // so an exception after an await reached App's handler and ended CRT. Logged, and CRT goes on.
            try
            {
                var dialog = new WorklogAddWorkDoneWindow();
                var result = await dialog.ShowDialog<(string Text, double HoursSpent, double Cost)?>(this);
                if (result == null)
                    return;

                int nextId = RecordIds.Next(this.thisEntry.WorkDoneItems.Select(w => w.Id));
                this.thisEntry.WorkDoneItems.Add(new WorklogWorkDoneRecord
                {
                    Id = nextId,
                    Text = result.Value.Text,
                    Date = DateTime.Now,
                    HoursSpent = result.Value.HoursSpent,
                    Cost = result.Value.Cost,

                    // The currency is recorded ON the row, at the moment the figure is typed - the
                    // dialog labels its Cost field with this same setting, so this is the code the user
                    // was looking at. Reading the setting later instead would relabel every historical
                    // cost the next time the preference changed. See WorklogWorkDoneRecord.CurrencyCode.
                    CurrencyCode = UserSettings.WorklogCurrencyCode
                });

                this.EnsureListSectionExpanded("EditorWorkDoneHeader");
                this.RefreshWorkDoneRows();
                this.PersistEntrySilently();
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnAddWorkDoneClick failed - [{ex}]");
            }
        }

        private void OnDeleteWorkDoneClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: int id })
            {
                this.thisEntry.WorkDoneItems.RemoveAll(w => w.Id == id);
                this.RefreshWorkDoneRows();
                this.PersistEntrySilently();
            }
        }

        // ###########################################################################################
        // Clicking anywhere on a work-done row (other than its Delete icon, which handles its own
        // Click and so never reaches here) reopens the same modal "Add work" uses, pre-filled for
        // editing - same click-to-edit behavior as the Links and Comments rows.
        // ###########################################################################################
        private async void OnWorkDoneRowPointerPressed(object? sender, PointerPressedEventArgs e)
        {
            // Guarded as a whole (code review, 2026-10-13): an "async void" handler has no caller,
            // so an exception after an await reached App's handler and ended CRT. Logged, and CRT goes on.
            try
            {
                if (sender is not Border { Tag: int id })
                    return;

                var work = this.thisEntry.WorkDoneItems.FirstOrDefault(w => w.Id == id);
                if (work == null)
                    return;

                var dialog = new WorklogAddWorkDoneWindow();
                dialog.InitializeForEdit(work.Text, work.HoursSpent, work.Cost);
                var result = await dialog.ShowDialog<(string Text, double HoursSpent, double Cost)?>(this);
                if (result == null)
                    return;

                work.Text = result.Value.Text;
                work.HoursSpent = result.Value.HoursSpent;
                work.Cost = result.Value.Cost;
                this.RefreshWorkDoneRows();
                this.PersistEntrySilently();
            }
            catch (Exception ex)
            {
                Logger.Critical($"OnWorkDoneRowPointerPressed failed - [{ex}]");
            }
        }
    }
}
