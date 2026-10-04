using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // THE "BETA" SCREEN'S LIST (owner request, 2026-09-27): the systems whose BETA data is ahead of
    // production that this account may publish - the SERVER decides the list. Choosing one shows it
    // in BetaView on the right. It was the left half of the "Publish to production" window.
    //
    // *** A CHECK NOBODY ASKED FOR NEVER THROWS AWAY A TICK. *** The list is read with the queue
    // every minute. The system on screen keeps its plan - and the maintainer's "I have checked this
    // in BETA" tick - unless its row CHANGED (BETA moved, or the other approver acted), when the plan
    // is worked out again: a tick given to other content must not carry over. A system that LEFT the
    // list meanwhile (published or pushed back by somebody else) empties the panel and says so.
    //
    // A system that waits only for the OTHER approver is dimmed and says so, the queue's way, and is
    // not counted on the BETA button's badge.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private readonly List<ProductionSystemRow> thisBeta = [];

        // Whether the list has been read since signing in - the badge says nothing until it has.
        private bool thisBetaKnown;

        // True while code - not the maintainer - moves the list's selection.
        private bool thisSuppressBetaSelection;

        private async Task RefreshBetaAsync(bool background)
        {
            if (!this.CanAskServer)
                return;

            ReviewApiResult<ProductionListResponse> result = await this.thisClient!.GetProductionListAsync(this.thisSession!);

            if (!result.IsOk)
            {
                // The list is left as it was: "could not ask" must not read as "nothing waiting".
                this.ShowBetaListMessage(result.Message, isError: true);
                return;
            }

            await this.ApplyBetaListAsync(result.Value!, background);
        }

        // ###########################################################################################
        // Puts the server's list on screen, keeping the chosen system chosen - see the header for
        // what happens to the panel beside it. Internal so a test can hand it an answer.
        // ###########################################################################################
        internal async Task ApplyBetaListAsync(ProductionListResponse response, bool background)
        {
            ArgumentNullException.ThrowIfNull(response);

            ProductionSystemRow? before = this.SelectedBetaRow;

            this.thisBeta.Clear();
            this.thisBeta.AddRange(response.Systems);
            this.thisBetaKnown = true;

            ProductionSystemRow? kept = before is null
                ? null
                : this.thisBeta.FirstOrDefault(row => string.Equals(row.SystemId, before.SystemId, StringComparison.Ordinal));

            if (this.FindControl<ListBox>("BetaList") is ListBox list)
            {
                this.thisSuppressBetaSelection = true;

                try
                {
                    List<ListBoxItem> items = this.thisBeta.Select(TabMaintainer.BetaItem).ToList();
                    list.ItemsSource = items;
                    list.SelectedItem = kept is null ? null : items.FirstOrDefault(item => ReferenceEquals(item.Tag, kept));
                }
                finally
                {
                    this.thisSuppressBetaSelection = false;
                }
            }

            if (!response.Configured)
                this.ShowBetaListMessage("Publishing to the stable source is not switched on for this server.", isError: true);
            else if (this.thisBeta.Count == 0)
                this.ShowBetaListMessage("The stable source is up to date with BETA for every system you review.", isError: false);
            else
                this.ShowBetaListMessage(null, isError: false);

            this.UpdateModeBadges();

            // A system arriving in or leaving BETA changes whether the open submission may be approved.
            this.ReapplyApprovalGate();

            if (before is null)
                return;

            if (kept is null)
            {
                if (background)
                    this.BetaDetail.ShowGoneElsewhere();
                else
                    await this.BetaDetail.ShowSystemAsync(null);

                return;
            }

            if (!background || !Equals(before, kept))
                await this.BetaDetail.ShowSystemAsync(kept);
        }

        private async void OnBetaSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisSuppressBetaSelection)
                return;

            ProductionSystemRow? row = this.SelectedBetaRow;

            if (row is null)
            {
                await this.BetaDetail.ShowSystemAsync(null);
                return;
            }

            // What the screen opens on next time (TabMaintainer.OpenOnEntry.cs).
            this.RememberBetaSystem(row.SystemId);

            // Chosen by the maintainer, so waited for under the overlay (2026-09-28) - the queue's
            // own minute check re-reads a changed row without it.
            if (!await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingProductionPlan(row.SystemId), () => this.BetaDetail.ShowSystemAsync(row)))
                this.BetaDetail.ShowMessage(WaitWording.NoAnswer, isError: true);
        }

        // ###########################################################################################
        // After a publish or a push-back: the BETA list (the system usually leaves it), the queue (a
        // push-back returns its submissions there) and the systems. Awaited by BetaView before it
        // says what happened, so the re-read cannot clear that sentence.
        // ###########################################################################################
        private async Task AfterBetaChangeAsync()
        {
            await this.RefreshBetaAsync(background: false);
            await this.RefreshQueueAsync(background: true);
            await this.RefreshSystemsAsync(background: true);
        }

        private ProductionSystemRow? SelectedBetaRow =>
            (this.FindControl<ListBox>("BetaList")?.SelectedItem as ListBoxItem)?.Tag as ProductionSystemRow;

        internal ProductionSystemRow? SelectedBetaRowForTests => this.SelectedBetaRow;

        // Chooses a system in the list as the maintainer would - for tests.
        internal void SelectBetaRowForTests(string systemId)
        {
            if (this.FindControl<ListBox>("BetaList") is ListBox list && list.ItemsSource is IEnumerable<ListBoxItem> items)
                list.SelectedItem = items.FirstOrDefault(item => (item.Tag as ProductionSystemRow)?.SystemId == systemId);
        }

        // Every text the list shows, top to bottom - for tests.
        internal IReadOnlyList<string> BetaTextsForTests() =>
            this.FindControl<ListBox>("BetaList")?.ItemsSource is IEnumerable<ListBoxItem> items
                ? items.SelectMany(item => item.Content is Control content ? TabMaintainer.VisibleTexts(content) : []).ToList()
                : [];

        // Signed out: nothing of the previous account's list stays.
        private void ClearBeta()
        {
            this.thisBeta.Clear();
            this.thisBetaKnown = false;

            if (this.FindControl<ListBox>("BetaList") is ListBox list)
            {
                this.thisSuppressBetaSelection = true;
                list.ItemsSource = null;
                this.thisSuppressBetaSelection = false;
            }

            this.ShowBetaListMessage(null, isError: false);
            this.BetaDetail.Clear();
        }

        private void ShowBetaListMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("BetaListMessageText"), message, isError);

        // One system: its name, and one grey line - where BETA and production stand, and when it
        // waits for the other approver rather than this account (then dimmed, the queue's way).
        private static ListBoxItem BetaItem(ProductionSystemRow row)
        {
            var content = new StackPanel { Spacing = 1, Opacity = row.AwaitsYou == false ? 0.55 : 1 };

            content.Children.Add(new TextBlock
            {
                Text = ProductionDisplay.SystemName(row),
                FontSize = 13,
                TextWrapping = TextWrapping.Wrap
            });

            content.Children.Add(new TextBlock
            {
                Text = ProductionDisplay.ListFooter(row),
                FontSize = 11,
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap
            });

            // A contributor whose work this carries discarded their own draft (owner request,
            // 2026-09-28) - said on the list, before the system is even opened.
            if (row.CarriesDiscardedDraft == true)
            {
                content.Children.Add(new TextBlock
                {
                    Text = DraftDiscardWording.ListMark,
                    FontSize = 11,
                    FontWeight = FontWeight.SemiBold,
                    Foreground = Brushes.IndianRed,
                    TextWrapping = TextWrapping.Wrap
                });
            }

            return new ListBoxItem { Tag = row, Classes = { "ListEntry" }, Content = content };
        }
    }
}
