using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Media;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // THE "SYSTEMS" SCREEN'S LIST (owner request, 2026-09-27): every system - the `systems` rows and
    // the boards in the BETA tree - each as its name and one grey line. Choosing one shows its
    // maintainers, contributors and submissions in SystemView on the right.
    //
    // For EVERY maintainer, every system (owner decision: "Everything for everyone") - the server's
    // SystemOverviewFlow says why. The Systems button's discreet badge is the count.
    //
    // A check nobody asked for re-reads the chosen system's detail only when its row changed, and
    // then without blanking the panel first - the same rule the queue and the BETA list keep.
    //
    // *** THE DROP-DOWN LISTS ARE READ WITH IT (owner request, 2026-09-27). *** A new system must be
    // placed in CRT's drop-down lists before it can be published to BETA, and it is placed here. So
    // the list is in CRT's drop-down order with the systems waiting for a place FIRST and marked,
    // the Systems button's badge turns into an attention badge counting the ones this account can
    // place, and the chosen system's panel carries SystemPlacementView. A listing that could not be
    // read leaves the previous one in place - never "not in the lists" on a failed request.
    // ###########################################################################################
    public partial class TabMaintainer
    {
        private readonly List<SystemOverviewEntry> thisSystems = [];

        // BETA's drop-down lists and the systems not in them, as last read.
        private SystemListingAnswer? thisListing;

        // Whether the list has been read since signing in - the badge says nothing until it has.
        private bool thisSystemsKnown;

        private bool thisSuppressSystemsSelection;

        private async Task RefreshSystemsAsync(bool background)
        {
            if (!this.CanAskServer)
                return;

            ReviewApiResult<SystemOverviewAnswer> result = await this.thisClient!.GetSystemOverviewAsync(this.thisSession!);

            if (!result.IsOk)
            {
                // Left as it was: "could not ask" must not read as "there are none".
                this.ShowSystemsListMessage(result.Message, isError: true);
                return;
            }

            ReviewApiResult<SystemListingAnswer> listing = await this.thisClient!.GetSystemListingAsync(this.thisSession!);

            await this.ApplySystemsListAsync(result.Value!, background, listing.IsOk ? listing.Value : null);
        }

        // After a placement was saved: the lists again, the chosen system kept.
        private Task AfterPlacementSavedAsync() => this.RefreshSystemsAsync(background: true);

        // ###########################################################################################
        // The minute check away from the Systems screen (code review, 2026-09-27): the drop-down
        // listing only, which the Approve gate and the Systems badge read - not the overview, whose
        // answer walks both data trees on the server. The list on the Systems screen is brought up
        // to date when that screen is shown again.
        // ###########################################################################################
        private async Task RefreshListingAsync()
        {
            if (!this.CanAskServer)
                return;

            ReviewApiResult<SystemListingAnswer> listing = await this.thisClient!.GetSystemListingAsync(this.thisSession!);

            // Left as it was: "could not ask" must not read as "nothing needs a place".
            if (!listing.IsOk || listing.Value is null)
                return;

            this.thisListing = listing.Value;
            this.UpdateModeBadges();
            this.ReapplyApprovalGate();
        }

        // `listing` null keeps the one already known (a failed read, or a test that has none).
        internal async Task ApplySystemsListAsync(SystemOverviewAnswer answer, bool background, SystemListingAnswer? listing = null)
        {
            ArgumentNullException.ThrowIfNull(answer);

            SystemOverviewEntry? before = this.SelectedSystem;

            if (listing is not null)
                this.thisListing = listing;

            this.thisSystems.Clear();
            this.thisSystems.AddRange(SystemPlacementDisplay.InListOrder(answer.Systems, this.thisListing));
            this.thisSystemsKnown = true;

            SystemOverviewEntry? kept = before is null
                ? null
                : this.thisSystems.FirstOrDefault(system => string.Equals(system.SystemId, before.SystemId, StringComparison.Ordinal));

            if (this.FindControl<ListBox>("SystemsList") is ListBox list)
            {
                this.thisSuppressSystemsSelection = true;

                try
                {
                    List<ListBoxItem> items = this.thisSystems
                        .Select(system => TabMaintainer.SystemItem(system, SystemPlacementDisplay.ListMark(this.thisListing, system.SystemId)))
                        .ToList();
                    list.ItemsSource = items;
                    list.SelectedItem = kept is null ? null : items.FirstOrDefault(item => ReferenceEquals(item.Tag, kept));
                }
                finally
                {
                    this.thisSuppressSystemsSelection = false;
                }
            }

            this.ShowSystemsListMessage(this.thisSystems.Count == 0 ? "There are no systems yet." : null, isError: false);
            this.UpdateModeBadges();

            // Before the detail: a system just placed into BETA's lists shows as listed at once.
            this.SystemDetail.UseListing(this.thisListing);

            // A new system placed since changes whether the open submission may be approved.
            this.ReapplyApprovalGate();

            if (before is null)
                return;

            if (kept is null)
            {
                await this.SystemDetail.ShowSystemAsync(null);
                return;
            }

            if (!background || QueueRefreshRules.SystemChanged(before, kept))
                await this.SystemDetail.ShowSystemAsync(kept, keepShown: background);
        }

        private async void OnSystemsSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisSuppressSystemsSelection)
                return;

            SystemOverviewEntry? system = this.SelectedSystem;

            if (system is null)
            {
                await this.SystemDetail.ShowSystemAsync(null);
                return;
            }

            // Chosen by the maintainer, so waited for under the overlay (2026-09-28) - the minute
            // check re-reads a changed system without it.
            if (!await ServerWait.RunAsync(this, MaintainerWaitWording.ReadingSystem(system.SystemId), () => this.SystemDetail.ShowSystemAsync(system)))
                this.SystemDetail.ShowMessage(WaitWording.NoAnswer, isError: true);
        }

        private SystemOverviewEntry? SelectedSystem =>
            (this.FindControl<ListBox>("SystemsList")?.SelectedItem as ListBoxItem)?.Tag as SystemOverviewEntry;

        internal SystemOverviewEntry? SelectedSystemForTests => this.SelectedSystem;

        internal IReadOnlyList<string> SystemsTextsForTests() =>
            this.FindControl<ListBox>("SystemsList")?.ItemsSource is IEnumerable<ListBoxItem> items
                ? items.SelectMany(item => item.Content is Control content ? TabMaintainer.VisibleTexts(content) : []).ToList()
                : [];

        private void ClearSystems()
        {
            this.thisSystems.Clear();
            this.thisListing = null;
            this.thisSystemsKnown = false;

            if (this.FindControl<ListBox>("SystemsList") is ListBox list)
            {
                this.thisSuppressSystemsSelection = true;
                list.ItemsSource = null;
                this.thisSuppressSystemsSelection = false;
            }

            this.ShowSystemsListMessage(null, isError: false);
            this.SystemDetail.Clear();
        }

        private void ShowSystemsListMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("SystemsListMessageText"), message, isError);

        // One system: its name, and where its data is, how many maintain it, and - when it is so -
        // that it is closed to contributions. A system waiting for a place in the drop-down lists
        // carries `mark` under it, in the queue's new-system colours.
        private static ListBoxItem SystemItem(SystemOverviewEntry system, string? mark)
        {
            var content = new StackPanel { Spacing = 1 };

            content.Children.Add(new TextBlock
            {
                Text = SystemsDisplay.Name(system),
                FontSize = 13,
                TextWrapping = TextWrapping.Wrap
            });

            // Where its data is and who maintains it, what is still to be done in bold.
            var line = new TextBlock
            {
                FontSize = 11,
                Opacity = 0.7,
                TextWrapping = TextWrapping.Wrap
            };

            TabMaintainer.ShowParts(line, SystemsDisplay.ListLineParts(system));
            content.Children.Add(line);

            if (mark is not null)
            {
                var text = new TextBlock { Text = mark, FontSize = 11, TextWrapping = TextWrapping.Wrap };

                var badge = new Border
                {
                    Classes = { "PlacementMark" },
                    BorderThickness = new Avalonia.Thickness(1),
                    CornerRadius = new Avalonia.CornerRadius(3),
                    Padding = new Avalonia.Thickness(6, 1),
                    Margin = new Avalonia.Thickness(0, 3, 0, 0),
                    HorizontalAlignment = Avalonia.Layout.HorizontalAlignment.Left,
                    Child = text
                };

                badge.Bind(Border.BackgroundProperty, badge.GetResourceObservable("Badge_NewSystem_Bg"));
                badge.Bind(Border.BorderBrushProperty, badge.GetResourceObservable("Badge_NewSystem_Border"));
                text.Bind(TextBlock.ForegroundProperty, text.GetResourceObservable("Badge_NewSystem_Fg"));

                content.Children.Add(badge);
            }

            return new ListBoxItem { Tag = system, Classes = { "ListEntry" }, Content = content };
        }
    }
}
