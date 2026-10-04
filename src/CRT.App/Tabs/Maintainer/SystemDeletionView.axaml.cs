using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Controls;
using Avalonia.Layout;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // "Delete a system" on the Account screen (owner request, 2026-10-03). See the markup for what it
    // is; the server's SystemDeletionFlow for what a delete does and refuses.
    //
    // *** PLAN, CONFIRM, DELETE - in that order, every time. *** Delete first asks the server what
    // would go (SystemDeletePlanAnswer) and shows it in DeleteSystemWindow; the delete then sends
    // the plan's fingerprint back, so the server deletes nothing the administrator was not shown. A
    // second press after a delete that stopped part-way plans afresh, which is what finishes it.
    //
    // *** NOTHING IS READ UNTIL IT IS CHOSEN. *** The systems are read when this item is chosen on
    // the Account screen and after each delete - not on the minute check, which this screen does not
    // need.
    // ###########################################################################################
    public partial class SystemDeletionView : UserControl
    {
        private ReviewApiClient? thisClient;
        private ReviewSession? thisSession;

        public SystemDeletionView()
        {
            this.InitializeComponent();

            this.SetText("HeadingText", SystemDeletionWording.ListHeading);
            this.SetText("ExplanationText", SystemDeletionWording.Explanation);
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // After a delete went through: the tab reads its queue, BETA list and systems again.
        public Func<Task>? AfterDelete { get; set; }

        // ###########################################################################################
        // The confirmation, for tests: ShowDialog blocks headlessly, so a test answers in its place.
        // The shipped path builds DeleteSystemWindow through the same Initialize a test of the window
        // drives.
        // ###########################################################################################
        internal Func<SystemDeletePlanAnswer, Task<(bool Confirmed, string Reason)>>? ConfirmOverrideForTests { get; set; }

        public void Initialize(ReviewApiClient? client, ReviewSession? session)
        {
            this.thisClient = client;
            this.thisSession = session;
        }

        // Signed out: nothing of the previous account's list or message left on screen.
        public void Clear()
        {
            this.FindControl<StackPanel>("SystemsPanel")?.Children.Clear();
            this.ShowMessage(null, isError: false);
        }

        // ###########################################################################################
        // Reads every system, under CRT's overlay. A list that cannot be read leaves the previous one
        // in place, with the reason - "could not ask" must never read as "there are none".
        // ###########################################################################################
        public async Task LoadAsync()
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;

            ReviewApiResult<SystemOverviewAnswer> result = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ReadingSystems,
                token => client.GetSystemOverviewAsync(session, token));

            if (!result.IsOk)
            {
                this.ShowMessage(result.Message, isError: true);
                return;
            }

            this.ShowSystems(result.Value!.Systems);
        }

        // ###########################################################################################
        // One row per system, by name: its name, its grey line - the Systems screen's own words - and
        // a red Delete button. Internal so a test can lay a list out without a server.
        // ###########################################################################################
        internal void ShowSystems(IReadOnlyList<SystemOverviewEntry> systems)
        {
            ArgumentNullException.ThrowIfNull(systems);

            if (this.FindControl<StackPanel>("SystemsPanel") is not StackPanel panel)
                return;

            panel.Children.Clear();

            foreach (SystemOverviewEntry system in systems.OrderBy(SystemsDisplay.Name, StringComparer.OrdinalIgnoreCase))
                panel.Children.Add(this.SystemRow(system));

            if (systems.Count == 0)
                panel.Children.Add(new TextBlock { Text = SystemDeletionWording.NoSystems, Opacity = 0.7 });
        }

        private Control SystemRow(SystemOverviewEntry system)
        {
            var texts = new StackPanel { Spacing = 1, VerticalAlignment = VerticalAlignment.Center };

            texts.Children.Add(new TextBlock { Text = SystemsDisplay.Name(system), FontSize = 13, TextWrapping = TextWrapping.Wrap });

            string line = SystemsDisplay.ListLine(system);

            if (line.Length > 0)
                texts.Children.Add(new TextBlock { Text = line, FontSize = 11, Opacity = 0.7, TextWrapping = TextWrapping.Wrap });

            var delete = new Button
            {
                Content = SystemDeletionWording.DeleteButton,
                Tag = system,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Thickness(12, 0, 0, 0)
            };

            delete.Classes.Add("DeleteSystemButton");
            delete.Bind(Button.ForegroundProperty, delete.GetResourceObservable("Button_Cancel_Fg"));
            delete.Bind(Button.BackgroundProperty, delete.GetResourceObservable("Button_Cancel_Bg"));
            delete.Bind(Button.BorderBrushProperty, delete.GetResourceObservable("Button_Cancel_Border"));
            delete.Click += async (_, _) => await this.DeleteAsync(system);

            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto") };
            grid.Children.Add(texts);
            grid.Children.Add(delete);
            Grid.SetColumn(delete, 1);

            var row = new Border
            {
                Padding = new Thickness(0, 6),
                BorderThickness = new Thickness(0, 0, 0, 1),
                Child = grid
            };

            row.Bind(Border.BorderBrushProperty, row.GetResourceObservable("Queue_Divider"));

            return row;
        }

        // ###########################################################################################
        // Plan, confirm, delete - see the header. Every Delete button is off for the duration, so no
        // second delete starts while one is being confirmed or carried out. Internal so a test can
        // await what a click starts.
        // ###########################################################################################
        internal async Task DeleteAsync(SystemOverviewEntry system)
        {
            if (this.thisClient is null || this.thisSession is null)
                return;

            ReviewApiClient client = this.thisClient;
            ReviewSession session = this.thisSession;

            this.SetDeleteButtonsEnabled(false);
            this.ShowMessage(null, isError: false);

            try
            {
                ReviewApiResult<SystemDeletePlanAnswer> planned = await ServerWait.CallAsync(
                    this,
                    MaintainerWaitWording.ReadingDeletionPlan(system.SystemId),
                    token => client.GetSystemDeletePlanAsync(session, system.SystemId, token));

                if (!planned.IsOk)
                {
                    this.ShowMessage(planned.Message, isError: true);
                    return;
                }

                SystemDeletePlanAnswer plan = planned.Value!;

                // Cannot go at all: said here, in the server's words, and nothing to confirm.
                if (!string.IsNullOrWhiteSpace(plan.BlockedBecause))
                {
                    this.ShowMessage(plan.BlockedBecause, isError: true);
                    return;
                }

                (bool confirmed, string reason) = await this.ConfirmAsync(plan);

                if (!confirmed)
                    return;

                await this.CarryOutAsync(client, session, plan, reason);
            }
            finally
            {
                this.SetDeleteButtonsEnabled(true);
            }
        }

        private async Task<(bool Confirmed, string Reason)> ConfirmAsync(SystemDeletePlanAnswer plan)
        {
            if (this.ConfirmOverrideForTests is not null)
                return await this.ConfirmOverrideForTests(plan);

            if (TopLevel.GetTopLevel(this) is not Window owner)
                return (false, string.Empty);

            var confirm = new DeleteSystemWindow();
            confirm.Initialize(plan);

            await confirm.ShowDialog(owner);

            return (confirm.WasConfirmed, confirm.Reason);
        }

        // ###########################################################################################
        // The delete, held under the overlay until the lists are read again too (BusyOverlay.HoldAsync
        // - one wait, no flicker between the steps). No answer in two minutes is never "it failed":
        // whether the system is still listed says whether it went (MaintainerWaitWording).
        // ###########################################################################################
        private async Task CarryOutAsync(ReviewApiClient client, ReviewSession session, SystemDeletePlanAnswer plan, string reason)
        {
            string waiting = MaintainerWaitWording.DeletingSystem(plan.SystemId);
            string? outcome = null;
            bool failed = true;

            await BusyOverlay.HoldAsync(this, waiting, async () =>
            {
                ReviewApiResult<SystemDeleteAnswer> result = await ServerWait.CallAsync(
                    this,
                    waiting,
                    token => client.DeleteSystemAsync(session, plan.SystemId, plan.Fingerprint, reason, token));

                if (result.Failure == ReviewApiFailure.TimedOut)
                {
                    bool? stillListed = await this.IsStillListedAsync(client, session, plan.SystemId);

                    outcome = MaintainerWaitWording.DeleteAfterTimeout(plan.SystemId, stillListed);
                    failed = stillListed != false;
                    await this.AfterChangeAsync();
                    return;
                }

                if (!result.IsOk)
                {
                    outcome = result.Message;

                    // A delete refused part-way has still changed something - the lists show what is left.
                    await this.LoadAsync();
                    return;
                }

                outcome = SystemDeletionWording.Done(result.Value!);
                failed = false;
                await this.AfterChangeAsync();
            });

            this.ShowMessage(outcome ?? "The delete did not complete.", isError: failed);
        }

        // The list again, and the tab's own lists. Each reads quietly; a failed read leaves its list.
        private async Task AfterChangeAsync()
        {
            await this.LoadAsync();

            if (this.AfterDelete is not null)
                await this.AfterDelete();
        }

        // Null when that look failed too.
        private async Task<bool?> IsStillListedAsync(ReviewApiClient client, ReviewSession session, string systemId)
        {
            ReviewApiResult<SystemOverviewAnswer> list = await ServerWait.CallAsync(
                this,
                MaintainerWaitWording.ReadingSystems,
                token => client.GetSystemOverviewAsync(session, token));

            return list.IsOk
                ? list.Value!.Systems.Any(entry => string.Equals(entry.SystemId, systemId, StringComparison.Ordinal))
                : null;
        }

        private void SetDeleteButtonsEnabled(bool enabled)
        {
            foreach (Button button in this.DeleteButtons())
                button.IsEnabled = enabled;
        }

        // Every row's Delete button, in the order shown.
        internal IReadOnlyList<Button> DeleteButtons() =>
            this.FindControl<StackPanel>("SystemsPanel")?.Children
                .OfType<Border>()
                .Select(row => row.Child)
                .OfType<Grid>()
                .SelectMany(grid => grid.Children.OfType<Button>())
                .ToList() ?? [];

        // Every system's name, in the order shown - for tests.
        internal IReadOnlyList<string> NamesForTests() =>
            this.FindControl<StackPanel>("SystemsPanel")?.Children
                .OfType<Border>()
                .Select(row => row.Child)
                .OfType<Grid>()
                .SelectMany(grid => grid.Children.OfType<StackPanel>())
                .Select(texts => texts.Children.OfType<TextBlock>().First().Text ?? string.Empty)
                .ToList() ?? [];

        internal string? MessageForTests => this.FindControl<TextBlock>("MessageText") is { IsVisible: true } message ? message.Text : null;

        private void SetText(string name, string text)
        {
            if (this.FindControl<TextBlock>(name) is TextBlock block)
                block.Text = text;
        }

        private void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);
    }
}
