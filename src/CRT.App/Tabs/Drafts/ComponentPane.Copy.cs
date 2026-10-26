using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // COPY FROM ANOTHER BOARD (CRT 4.0 step 2, slice 5, 2026-10-26 - the plan: "Copy from another
    // board pulls data across for a part with the same part number").
    //
    // The same chip on another board of this machine - found by its PART NUMBER (ComponentCopy, the
    // host's FindOnOtherBoards reading the published boards) - lends this component what was worked out
    // there: its package, short description and repair notes, its pins, replacement parts and links.
    // A panel above the sections picks the board and what to copy (each ticked when it has something);
    // "Copy" puts it into the draft as ONE step, so Undo takes it back. Files and anything the board
    // decides itself (label, category, highlights, scope baselines) are never copied.
    //
    // AND THE OTHER WAY: "Also on other boards" puts what this draft changed of the part onto the same
    // part of other boards, each in its own draft (below).
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // The components of the other boards with a part number, in a region - the host's (TabDrafts).
        internal Func<string, string, Task<IReadOnlyList<ComponentCopySource>>>? FindOnOtherBoards { get; set; }

        internal bool IsCopyPanelOpen => this.CopyPanel.IsVisible;

        private async void OnCopyFromOtherBoardClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.OpenCopyPanelAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Looking for the component on the other boards failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The panel opened on the components of the other boards with this one's part number - or a
        // line saying why there are none.
        // ###########################################################################################
        internal async Task OpenCopyPanelAsync()
        {
            this.CommitHeld();
            this.ShowMessage(null);

            string partNumber = this.thisShownValues.TryGetValue(ComponentField.PartNumber, out string? shown) ? shown : string.Empty;

            if (partNumber.Length == 0)
            {
                this.ShowMessage(ComponentPaneWording.CopyNeedsPartNumber);
                return;
            }

            IReadOnlyList<ComponentCopySource> sources = this.FindOnOtherBoards is { } find
                ? await find(partNumber, this.thisRegion)
                : [];

            if (sources.Count == 0)
            {
                this.ShowMessage(ComponentPaneWording.CopyFoundNothing(partNumber));
                return;
            }

            this.CopyIntroText.Text = ComponentPaneWording.CopyIntro(partNumber);
            this.CopySourceBox.ItemsSource = sources;
            this.CopySourceBox.SelectedIndex = 0;
            this.ShowCopyParts();
            this.CopyPanel.IsVisible = true;
        }

        private void OnCopySourceChanged(object? sender, SelectionChangedEventArgs e) => this.ShowCopyParts();

        // A tick box for each part the board picked has - ticked.
        private void ShowCopyParts()
        {
            this.CopyPartsPanel.Children.Clear();

            if (this.CopySourceBox.SelectedItem is not ComponentCopySource source)
            {
                return;
            }

            foreach (ComponentCopyPart part in ComponentCopy.Available(source))
            {
                this.CopyPartsPanel.Children.Add(new CheckBox { Content = ComponentCopy.Describe(source, part), IsChecked = true, Tag = part });
            }

            this.CopyNothingText.IsVisible = this.CopyPartsPanel.Children.Count == 0;
            this.CopyButton.IsEnabled = this.CopyPartsPanel.Children.Count > 0;
        }

        private void OnCopyClick(object? sender, RoutedEventArgs e) => this.CopyPicked();

        // ###########################################################################################
        // What is ticked into the component, as one step named after the board it came from - then the
        // panel closed and the component read again. False when nothing changed.
        // ###########################################################################################
        internal bool CopyPicked()
        {
            if (this.thisSession is not { } session || this.CopySourceBox.SelectedItem is not ComponentCopySource source)
            {
                return false;
            }

            List<ComponentCopyPart> parts = this.CopyPartsPanel.Children
                .OfType<CheckBox>()
                .Where(box => box.IsChecked == true)
                .Select(box => box.Tag)
                .OfType<ComponentCopyPart>()
                .ToList();

            string label = this.thisLabel;
            string region = this.thisRegion;

            bool copied = session.Apply(
                DraftStepWording.ComponentCopied(label, source.BoardName),
                content => ComponentCopy.Apply(content.Board, label, region, source, parts) is BoardData board && !ReferenceEquals(board, content.Board)
                    ? content with { Board = board }
                    : null,
                this.thisOrigin);

            this.CloseCopyPanel();

            if (copied)
            {
                this.ShowComponent();
                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }

            return copied;
        }

        private void OnCancelCopyClick(object? sender, RoutedEventArgs e) => this.CloseCopyPanel();

        private void CloseCopyPanel()
        {
            this.CopyPanel.IsVisible = false;
            this.CopySourceBox.ItemsSource = null;
            this.CopyPartsPanel.Children.Clear();
            this.CloseAlsoPanel();
        }

        // ---- also on other boards -----------------------------------------------------------------

        // ###########################################################################################
        // "ALSO ON OTHER BOARDS" (the plan: "This datasheet is also used by 11 other boards - change it
        // there too? Each board gets the change in its own draft, reviewed by its own maintainers ... It
        // works for files, text, pins and replacements"; default accepted 2026-10-24: every board
        // ticked). Offered while the draft changes this component and it has a part number: what the
        // draft changed of it (ComponentCopy.Changed) put onto the same part of the ticked boards - by
        // the host (ApplyToOtherBoards), each board's own draft made when it has none, one step there.
        // Files are not carried yet (their bytes are this draft's).
        // ###########################################################################################
        internal Func<ComponentCopySource, IReadOnlyList<ComponentCopySource>, IReadOnlyList<ComponentCopyPart>, Task<string>>? ApplyToOtherBoards { get; set; }

        internal bool IsAlsoPanelOpen => this.AlsoPanel.IsVisible;

        // The button, while the draft changes this component and it has a part number to match by.
        private void ShowAlsoOffer()
        {
            this.AlsoOnOtherBoardsButton.IsVisible =
                !this.thisIsNew && this.thisBoardPart is null && this.IsComponentChanged &&
                this.thisShownValues.TryGetValue(ComponentField.PartNumber, out string? partNumber) && partNumber.Length > 0;
        }

        private async void OnAlsoOnOtherBoardsClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.OpenAlsoPanelAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Looking for the component on the other boards failed - [{ex}]");
            }
        }

        // The panel opened on the same part of the other boards, every board ticked, and what changed.
        internal async Task OpenAlsoPanelAsync()
        {
            this.CommitHeld();
            this.ShowMessage(null);
            this.CloseCopyPanel();

            if (this.thisSession is not { } session ||
                !this.thisShownValues.TryGetValue(ComponentField.PartNumber, out string? partNumber) || partNumber.Length == 0)
            {
                this.ShowMessage(ComponentPaneWording.CopyNeedsPartNumber);
                return;
            }

            List<ComponentCopyPart> changed = ComponentCopy.Changed(session.Content.Board, this.thisPublished, this.thisLabel, this.thisRegion);

            if (changed.Count == 0)
            {
                this.ShowMessage(ComponentPaneWording.AlsoNothingChanged);
                return;
            }

            IReadOnlyList<ComponentCopySource> targets = this.FindOnOtherBoards is { } find
                ? await find(partNumber, this.thisRegion)
                : [];

            if (targets.Count == 0)
            {
                this.ShowMessage(ComponentPaneWording.CopyFoundNothing(partNumber));
                return;
            }

            ComponentCopySource here = new(string.Empty, session.Content.Board, ComponentPaneEdits.RowShown(session.Content.Board, this.thisLabel, this.thisRegion)!);

            this.AlsoIntroText.Text = ComponentPaneWording.AlsoIntro(partNumber);
            this.AlsoBoardsPanel.Children.Clear();
            this.AlsoPartsPanel.Children.Clear();

            foreach (ComponentCopySource target in targets)
            {
                this.AlsoBoardsPanel.Children.Add(new CheckBox { Content = target.DisplayText, IsChecked = true, Tag = target });
            }

            foreach (ComponentCopyPart part in changed)
            {
                this.AlsoPartsPanel.Children.Add(new CheckBox { Content = ComponentCopy.Describe(here, part), IsChecked = true, Tag = part });
            }

            this.AlsoPanel.IsVisible = true;
        }

        private async void OnAlsoChangeClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.ChangeOtherBoardsAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Changing the component on the other boards failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // What is ticked onto the ticked boards, through the host - the panel closed, what came of it
        // said under the title. The component here is not changed by it.
        // ###########################################################################################
        internal async Task ChangeOtherBoardsAsync()
        {
            if (this.thisSession is not { } session || this.ApplyToOtherBoards is not { } apply ||
                ComponentPaneEdits.RowShown(session.Content.Board, this.thisLabel, this.thisRegion) is not ComponentEntry row)
            {
                return;
            }

            List<ComponentCopySource> targets = this.AlsoBoardsPanel.Children.OfType<CheckBox>()
                .Where(box => box.IsChecked == true).Select(box => box.Tag).OfType<ComponentCopySource>().ToList();
            List<ComponentCopyPart> parts = this.AlsoPartsPanel.Children.OfType<CheckBox>()
                .Where(box => box.IsChecked == true).Select(box => box.Tag).OfType<ComponentCopyPart>().ToList();

            this.CloseAlsoPanel();

            if (targets.Count == 0 || parts.Count == 0)
            {
                return;
            }

            string said = await apply(new ComponentCopySource(string.Empty, session.Content.Board, row), targets, parts);
            this.ShowMessage(said);
        }

        private void OnCancelAlsoClick(object? sender, RoutedEventArgs e) => this.CloseAlsoPanel();

        private void CloseAlsoPanel()
        {
            this.AlsoPanel.IsVisible = false;
            this.AlsoBoardsPanel.Children.Clear();
            this.AlsoPartsPanel.Children.Clear();
        }
    }
}
