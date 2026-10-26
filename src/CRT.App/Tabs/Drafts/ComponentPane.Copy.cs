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
    // part of other boards, each in its own draft - and a file it replaced (a datasheet, a picture, a
    // board file under Files and links) onto the boards that have the same file (below).
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
        // ticked). Offered while the draft changes this component: what the draft changed of it
        // (ComponentCopy.Changed) put onto the same part - the same part number - of the ticked boards,
        // and each file or picture it replaced (ChangedFiles) put in place of the same file on the
        // boards that have it - by the host (ApplyToOtherBoards), each board's own draft made when it
        // has none.
        // ###########################################################################################
        internal Func<ComponentCopySource?, IReadOnlyList<ComponentCopySource>, IReadOnlyList<ComponentCopyPart>, IReadOnlyList<ChangedFileElsewhere>, Task<string>>? ApplyToOtherBoards { get; set; }

        // ###########################################################################################
        // What "Also on other boards" offers, found by the host (TabDrafts) in ONE read of the other
        // boards (code review, 2026-10-26: two searches, two "please wait"s, every board read twice):
        // the same part's components by the part number - "" for no such search - in a region, and
        // the boards listing the bytes each changed file replaced.
        // ###########################################################################################
        internal Func<string, string, IReadOnlyList<ChangedFile>, Task<AlsoTargets>>? FindAlsoTargets { get; set; }

        internal bool IsAlsoPanelOpen => this.AlsoPanel.IsVisible;

        internal bool IsBoardAlsoPanelOpen => this.BoardAlsoPanel.IsVisible;

        // The published board's list of files - what a file of the draft replaced is told by.
        private IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? PublishedFiles => DraftFiles.ListedIn(this.thisPublishedBoardFile);

        // The component's files and pictures the draft replaced, as the session last saw its list.
        private List<ChangedFile> ChangedFilesHere() =>
            this.thisSession is not { } session || this.thisLabel.Length == 0
                ? []
                : ChangedFiles.OfComponent(session.Content.Board, session.ListedFiles, this.thisPublished, this.PublishedFiles, this.thisLabel, this.thisRegion);

        // The board's own files the draft replaced.
        private List<ChangedFile> ChangedBoardFilesHere() =>
            this.thisSession is not { } session
                ? []
                : ChangedFiles.OfBoardFiles(session.Content.Board, session.ListedFiles, this.thisPublished, this.PublishedFiles);

        // Both offers as the draft now reads - after a change and after a save (a file's new bytes are
        // listed only once the draft is written).
        private void ShowAlsoOffers()
        {
            this.ShowAlsoOffer();
            this.ShowBoardFilesAlsoOffer();
        }

        // The button, while the draft changes this component - and it has a part number to match by,
        // or a file replaced to look for.
        private void ShowAlsoOffer()
        {
            this.AlsoOnOtherBoardsButton.IsVisible =
                !this.thisIsNew && this.thisBoardPart is null && this.IsComponentChanged &&
                ((this.thisShownValues.TryGetValue(ComponentField.PartNumber, out string? partNumber) && partNumber.Length > 0) ||
                 this.ChangedFilesHere().Count > 0);
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

        // ###########################################################################################
        // The panel opened on the same part of the other boards, every board ticked, and what changed -
        // and on the files replaced, each with the boards that have it, every one ticked. A line saying
        // why, when there is nothing to offer.
        // ###########################################################################################
        internal async Task OpenAlsoPanelAsync()
        {
            this.CommitHeld();
            this.ShowMessage(null);
            this.CloseCopyPanel();

            if (this.thisSession is not { } session)
            {
                return;
            }

            string partNumber = this.thisShownValues.TryGetValue(ComponentField.PartNumber, out string? shown) ? shown : string.Empty;

            List<ComponentCopyPart> changed = partNumber.Length > 0
                ? ComponentCopy.Changed(session.Content.Board, this.thisPublished, this.thisLabel, this.thisRegion)
                : [];
            List<ChangedFile> files = this.ChangedFilesHere();

            if (changed.Count == 0 && files.Count == 0)
            {
                this.ShowMessage(partNumber.Length == 0 ? ComponentPaneWording.CopyNeedsPartNumber : ComponentPaneWording.AlsoNothingChanged);
                return;
            }

            // The part is looked for only when something of it changed - and only then said not found
            // (code review, 2026-10-26).
            string partSearched = changed.Count > 0 ? partNumber : string.Empty;

            AlsoTargets found = this.FindAlsoTargets is { } find
                ? await find(partSearched, this.thisRegion, files)
                : AlsoTargets.None;

            IReadOnlyList<ComponentCopySource> targets = found.Parts;
            IReadOnlyList<ChangedFileElsewhere> fileTargets = found.Files;

            if (targets.Count == 0 && fileTargets.Count == 0)
            {
                this.ShowMessage(files.Count == 0 ? ComponentPaneWording.CopyFoundNothing(partNumber) : ComponentPaneWording.AlsoFoundNothing(partSearched));
                return;
            }

            this.AlsoBoardsPanel.Children.Clear();
            this.AlsoPartsPanel.Children.Clear();
            this.AlsoPartsSection.IsVisible = targets.Count > 0;

            if (targets.Count > 0)
            {
                ComponentCopySource here = new(string.Empty, session.Content.Board, ComponentPaneEdits.RowShown(session.Content.Board, this.thisLabel, this.thisRegion)!);

                this.AlsoIntroText.Text = ComponentPaneWording.AlsoIntro(partNumber);

                foreach (ComponentCopySource target in targets)
                {
                    this.AlsoBoardsPanel.Children.Add(new CheckBox { Content = target.DisplayText, IsChecked = true, Tag = target });
                }

                foreach (ComponentCopyPart part in changed)
                {
                    this.AlsoPartsPanel.Children.Add(new CheckBox { Content = ComponentCopy.Describe(here, part), IsChecked = true, Tag = part });
                }
            }

            this.AlsoFilesIntroText.Text = ComponentPaneWording.AlsoFilesIntro;
            ComponentPane.FillAlsoFiles(this.AlsoFilesPanel, fileTargets);
            this.AlsoFilesSection.IsVisible = fileTargets.Count > 0;

            this.AlsoPanel.IsVisible = true;

            // The part changed, but no other board has it - said, while its files are offered (code
            // review, 2026-10-26: the part's section was left out without a word).
            if (changed.Count > 0 && targets.Count == 0)
            {
                this.ShowMessage(ComponentPaneWording.CopyFoundNothing(partNumber));
            }
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
            List<ChangedFileElsewhere> files = ComponentPane.TickedAlsoFiles(this.AlsoFilesPanel);

            this.CloseAlsoPanel();

            if (targets.Count == 0 || parts.Count == 0)
            {
                targets = [];
                parts = [];
            }

            if (targets.Count == 0 && files.Count == 0)
            {
                return;
            }

            string said = await apply(new ComponentCopySource(string.Empty, session.Content.Board, row), targets, parts, files);
            this.ShowMessage(said);
        }

        private void OnCancelAlsoClick(object? sender, RoutedEventArgs e) => this.CloseAlsoPanel();

        private void CloseAlsoPanel()
        {
            this.AlsoPanel.IsVisible = false;
            this.AlsoBoardsPanel.Children.Clear();
            this.AlsoPartsPanel.Children.Clear();
            this.AlsoFilesPanel.Children.Clear();
            this.CloseBoardAlsoPanel();
        }

        // ---- the board's own files on other boards ------------------------------------------------

        // Under Files and links: the button, while the draft replaces one of the board's files.
        private void ShowBoardFilesAlsoOffer()
        {
            this.BoardFilesAlsoButton.IsVisible =
                this.thisBoardPart == PaneBoardPart.FilesAndLinks && this.ChangedBoardFilesHere().Count > 0;
        }

        private async void OnBoardFilesAlsoClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.OpenBoardAlsoPanelAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Looking for the board's files on the other boards failed - [{ex}]");
            }
        }

        // The board files replaced, each with the other boards that have it, every one ticked.
        internal async Task OpenBoardAlsoPanelAsync()
        {
            this.CommitHeld();
            this.ShowMessage(null);
            this.CloseBoardAlsoPanel();

            List<ChangedFile> files = this.ChangedBoardFilesHere();

            if (files.Count == 0)
            {
                this.ShowMessage(ComponentPaneWording.AlsoBoardFilesNothingChanged);
                return;
            }

            IReadOnlyList<ChangedFileElsewhere> targets = this.FindAlsoTargets is { } find
                ? (await find(string.Empty, string.Empty, files)).Files
                : [];

            if (targets.Count == 0)
            {
                this.ShowMessage(ComponentPaneWording.AlsoBoardFilesFoundNothing);
                return;
            }

            this.BoardAlsoIntroText.Text = ComponentPaneWording.AlsoBoardFilesIntro;
            ComponentPane.FillAlsoFiles(this.BoardAlsoFilesPanel, targets);
            this.BoardAlsoPanel.IsVisible = true;
        }

        private async void OnBoardAlsoChangeClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.ChangeBoardFilesOnOtherBoardsAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Changing the board's files on the other boards failed - [{ex}]");
            }
        }

        // The ticked files onto their boards, through the host - what came of it said under the title.
        internal async Task ChangeBoardFilesOnOtherBoardsAsync()
        {
            if (this.ApplyToOtherBoards is not { } apply)
            {
                return;
            }

            List<ChangedFileElsewhere> files = ComponentPane.TickedAlsoFiles(this.BoardAlsoFilesPanel);
            this.CloseBoardAlsoPanel();

            if (files.Count > 0)
            {
                this.ShowMessage(await apply(null, [], [], files));
            }
        }

        private void OnCancelBoardAlsoClick(object? sender, RoutedEventArgs e) => this.CloseBoardAlsoPanel();

        private void CloseBoardAlsoPanel()
        {
            this.BoardAlsoPanel.IsVisible = false;
            this.BoardAlsoFilesPanel.Children.Clear();
        }

        // ###########################################################################################
        // Each changed file named, with a tick box under it for each board that has it - ticked.
        // ###########################################################################################
        private static void FillAlsoFiles(StackPanel panel, IReadOnlyList<ChangedFileElsewhere> targets)
        {
            panel.Children.Clear();

            foreach (IGrouping<ChangedFile, ChangedFileElsewhere> file in targets.GroupBy(target => target.File))
            {
                panel.Children.Add(new TextBlock { Classes = { "FieldLabel" }, Text = file.Key.DisplayText });

                foreach (ChangedFileElsewhere target in file)
                {
                    panel.Children.Add(new CheckBox { Content = target.DisplayText, IsChecked = true, Tag = target });
                }
            }
        }

        private static List<ChangedFileElsewhere> TickedAlsoFiles(StackPanel panel) =>
            panel.Children.OfType<CheckBox>()
                .Where(box => box.IsChecked == true)
                .Select(box => box.Tag)
                .OfType<ChangedFileElsewhere>()
                .ToList();
    }

    // What "Also on other boards" can change: the same part's components, and the same files.
    public sealed record AlsoTargets(IReadOnlyList<ComponentCopySource> Parts, IReadOnlyList<ChangedFileElsewhere> Files)
    {
        public static AlsoTargets None { get; } = new([], []);
    }
}
