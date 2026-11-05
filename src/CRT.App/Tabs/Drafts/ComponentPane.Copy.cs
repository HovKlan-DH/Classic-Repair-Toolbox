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
    // CLONE A COMPONENT FROM ANOTHER BOARD (owner request, 2026-11-04: "I would like to get a selector, so I
    // can select a hardware, a board and then a component. When I have selected the component it will show
    // some of its most important data and I can then select to clone all its data to this component. It will
    // not change the component name, but all other data should be identical"). "Copy from another board" until
    // then, which found only the same part number and let parts be ticked.
    //
    // A panel above the sections picks the hardware, the board - read when picked, by the host - and the
    // component, and says what that one is (ComponentPaneWording.CloneFacts). "Clone" puts all of it into the
    // draft as ONE step, the files it cites copied in first; every rule is ComponentClone's. Only published
    // boards, never the one the pane is on and never another draft. It opens next time on the hardware and
    // board just cloned from, the component to be picked (thisCloneRemembered).
    //
    // AND THE OTHER WAY: "Also on other boards" puts what this draft changed of the part onto the same
    // part of other boards, each in its own draft - and a file it replaced (a datasheet, a picture, a
    // board file under Files and links) onto the boards that have the same file (below).
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // The published boards a component can be cloned from - never the one the pane is on. The host's.
        internal Func<IReadOnlyList<HardwareBoardEntry>>? CloneBoards { get; set; }

        // The hardware of the board shown, as the drop-down lists name it - what the panel opens on. The host's.
        internal Func<string?>? CloneHardwareShown { get; set; }

        // One of them read, off the UI thread, under CRT's "please wait". Null when it cannot be. The host's.
        internal Func<HardwareBoardEntry, Task<ComponentCloneSource?>>? ReadCloneSource { get; set; }

        private IReadOnlyList<HardwareBoardEntry> thisCloneBoards = [];

        // The board picked, as read - and which read of it is the latest: one started later wins.
        private ComponentCloneSource? thisCloneSource;

        private int thisCloneAsk;

        private Task thisCloneRead = Task.CompletedTask;

        // ###########################################################################################
        // The hardware and board last cloned from, for the run (owner request, 2026-11-04: "When I have used
        // the 'Clone component from another board' once, then do remember its hardware and board, but do
        // empty the component"): one board's chips usually go across one after another. Never the
        // component - each is its own pick.
        // ###########################################################################################
        private (string? Hardware, string? Board)? thisCloneRemembered;

        // True while the panel's lists are filled from code - no pick then.
        private bool thisFillingClone;

        internal bool IsClonePanelOpen => this.ClonePanel.IsVisible;

        // What the board picked last started - a test waits for it.
        internal Task CloneReadForTests => this.thisCloneRead;

        private void OnCloneFromOtherBoardClick(object? sender, RoutedEventArgs e) => this.OpenClonePanel();

        // ###########################################################################################
        // The panel opened on the hardware and board last cloned from, as far as they can still be picked -
        // that board read again, no component picked - else the hardware of the board shown when another
        // board of it can be picked; or a line saying there is nothing to clone from. At the top of the
        // page, where the panel is.
        // ###########################################################################################
        internal void OpenClonePanel()
        {
            this.CommitHeld();
            this.ShowMessage(null);
            this.CloseClonePanel();

            this.thisCloneBoards = this.CloneBoards?.Invoke() ?? [];

            if (this.thisCloneBoards.Count == 0)
            {
                this.ShowMessage(ComponentPaneWording.CloneNoOtherBoards);
                return;
            }

            List<string> hardware = this.thisCloneBoards
                .Select(board => board.HardwareName)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            string? shown = this.CloneHardwareShown?.Invoke();

            string? Offered(string? wanted) => hardware.Find(name => string.Equals(name, wanted, StringComparison.OrdinalIgnoreCase));

            this.CloneIntroText.Text = ComponentPaneWording.CloneIntro;

            this.thisFillingClone = true;

            try
            {
                this.CloneHardwareBox.ItemsSource = hardware;
                this.CloneHardwareBox.SelectedItem = Offered(this.thisCloneRemembered?.Hardware) ?? Offered(shown);
            }
            finally
            {
                this.thisFillingClone = false;
            }

            this.ShowCloneBoards();

            // Open before the board is picked: what a read finds is dropped with the panel shut.
            this.ClonePanel.IsVisible = true;
            this.SectionsScrollViewer.Offset = default;

            if (this.thisCloneRemembered?.Board is { } board &&
                this.CloneBoardBox.Items.OfType<string>().FirstOrDefault(name => string.Equals(name, board, StringComparison.OrdinalIgnoreCase)) is { } offered)
            {
                this.CloneBoardBox.SelectedItem = offered;
            }
        }

        private void OnCloneHardwareChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (!this.thisFillingClone)
            {
                this.ShowCloneBoards();
            }
        }

        // The boards of the hardware picked - nothing of the board picked before stays.
        private void ShowCloneBoards()
        {
            string? hardware = this.CloneHardwareBox.SelectedItem as string;

            this.thisFillingClone = true;

            try
            {
                this.CloneBoardBox.ItemsSource = this.thisCloneBoards
                    .Where(board => string.Equals(board.HardwareName, hardware, StringComparison.OrdinalIgnoreCase))
                    .Select(board => board.BoardName)
                    .ToList();
                this.CloneBoardBox.SelectedItem = null;
            }
            finally
            {
                this.thisFillingClone = false;
            }

            this.ForgetCloneSource();
        }

        private void ForgetCloneSource()
        {
            this.thisCloneAsk++;
            this.thisCloneSource = null;

            this.thisFillingClone = true;

            try
            {
                this.CloneComponentBox.ItemsSource = null;
                this.CloneComponentBox.SelectedItem = null;
            }
            finally
            {
                this.thisFillingClone = false;
            }

            this.CloneNoteText.IsVisible = false;
            this.ShowCloneFacts();
        }

        private void OnCloneBoardChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisFillingClone)
            {
                return;
            }

            this.ForgetCloneSource();

            string? hardware = this.CloneHardwareBox.SelectedItem as string;
            string? boardName = this.CloneBoardBox.SelectedItem as string;

            HardwareBoardEntry? board = this.thisCloneBoards.FirstOrDefault(entry =>
                string.Equals(entry.HardwareName, hardware, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(entry.BoardName, boardName, StringComparison.OrdinalIgnoreCase));

            if (board is not null)
            {
                this.thisCloneRead = this.ReadCloneBoardAsync(board, this.thisCloneAsk);
            }
        }

        // ###########################################################################################
        // The board picked, read - its components listed by the popup's title of each, as CRT's region shows
        // them. Nothing may leave here: it runs from a selection change, and App lets such an exception end CRT.
        // ###########################################################################################
        private async Task ReadCloneBoardAsync(HardwareBoardEntry board, int ask)
        {
            ComponentCloneSource? source = null;

            try
            {
                source = this.ReadCloneSource is { } read ? await read(board) : null;
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read [{board.ExcelDataFile}] to clone a component from it - [{ex.Message}]");
            }

            if (ask != this.thisCloneAsk || !this.ClonePanel.IsVisible)
            {
                return;
            }

            this.thisCloneSource = source;

            List<ComponentCloneChoice> choices = source is null ? [] : this.CloneChoicesOf(source);

            this.thisFillingClone = true;

            try
            {
                this.CloneComponentBox.ItemsSource = choices;
                this.CloneComponentBox.SelectedItem = null;
            }
            finally
            {
                this.thisFillingClone = false;
            }

            string? note = source is null ? ComponentPaneWording.CloneBoardNotRead
                : choices.Count == 0 ? ComponentPaneWording.CloneBoardHasNoComponents
                : null;

            this.CloneNoteText.Text = note ?? string.Empty;
            this.CloneNoteText.IsVisible = note is not null;
            this.ShowCloneFacts();
        }

        // Its components by the popup's title of each, as the region shown shows them.
        private List<ComponentCloneChoice> CloneChoicesOf(ComponentCloneSource source) =>
            ComponentClone.Labels(source.Board)
                .Select(label => new ComponentCloneChoice(label, ComponentPopupContent.TitleOf(label, ComponentPaneEdits.RowShown(source.Board, label, this.thisRegion))))
                .ToList();

        // ###########################################################################################
        // Another region shown, by the buttons under CRT's list, which stand beside the panel: its titles and what
        // it says of the component picked are the new region's, the pick staying (code review, 2026-11-04).
        // ###########################################################################################
        private void FollowRegionInClonePanel()
        {
            if (!this.ClonePanel.IsVisible || this.thisCloneSource is not { } source)
            {
                return;
            }

            string? picked = (this.CloneComponentBox.SelectedItem as ComponentCloneChoice)?.Label;
            List<ComponentCloneChoice> choices = this.CloneChoicesOf(source);

            this.thisFillingClone = true;

            try
            {
                this.CloneComponentBox.ItemsSource = choices;
                this.CloneComponentBox.SelectedItem = choices.Find(choice => string.Equals(choice.Label, picked, StringComparison.Ordinal));
            }
            finally
            {
                this.thisFillingClone = false;
            }

            this.ShowCloneFacts();
        }

        private void OnCloneComponentChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (!this.thisFillingClone)
            {
                this.ShowCloneFacts();
            }
        }

        // What the component picked is, and "Clone" only with one.
        private void ShowCloneFacts()
        {
            ComponentCloneFacts? facts =
                this.thisSession is { } session && this.thisCloneSource is { } source && this.CloneComponentBox.SelectedItem is ComponentCloneChoice picked
                    ? ComponentClone.FactsOf(session.Content.Board, this.thisLabel, source.Board, picked.Label, this.thisRegion)
                    : null;

            this.CloneFactsLines.Children.Clear();
            this.CloneFactsPanel.IsVisible = facts is not null;
            this.CloneButton.IsEnabled = facts is not null;

            foreach (string line in facts is null ? [] : ComponentPaneWording.CloneFacts(facts))
            {
                var block = new TextBlock { TextWrapping = Avalonia.Media.TextWrapping.Wrap };

                Handlers.Theming.BoldRuns.ShowMarked(block, line);
                this.CloneFactsLines.Children.Add(block);
            }
        }

        private async void OnCloneClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.ClonePickedAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Cloning a component from another board failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The component picked into this one as ONE step, named after where it came from, the files it cites
        // copied in first and the draft saved - then the panel closed and what came of it said. A file not
        // on this computer, or under a name that could not lie in the draft, does not come: what cites it
        // is left out, and said. False when nothing changed.
        // ###########################################################################################
        internal async Task<bool> ClonePickedAsync()
        {
            if (this.thisSession is not { } session ||
                this.thisCloneSource is not { } source ||
                this.CloneComponentBox.SelectedItem is not ComponentCloneChoice picked)
            {
                return false;
            }

            this.CommitHeld();

            this.thisCloneRemembered = (this.CloneHardwareBox.SelectedItem as string, this.CloneBoardBox.SelectedItem as string);

            string label = this.thisLabel;
            string region = this.thisRegion;
            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);
            BoardData board = session.Content.Board;

            List<ComponentCloneFile> files = ComponentClone.NamesHere(
                ComponentClone.FilesCited(board, label, source.Board, picked.Label, region),
                source.Files,
                session.ListedFiles,
                this.TakenNames());

            var names = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            var leftOut = new List<string>();
            var copies = new List<(string Source, string Name)>();
            var cameFrom = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);

            await BusyOverlay.RunLocalAsync(this, CrtWaitWording.CloningComponent, async () =>
            {
                // Found off the UI thread - a name is looked up in the data folder.
                List<(ComponentCloneFile File, string? Path)> found = await Task.Run(() => files
                    .Select(file => (file, file.CopyIn ? source.PathOf(file.NameThere) : null))
                    .ToList());

                foreach ((ComponentCloneFile file, string? path) in found)
                {
                    if (!file.CopyIn)
                    {
                        names[file.NameThere] = file.NameHere;
                    }
                    else if (path is null || !DraftFiles.TryArrivedPath(draftFolder, file.NameHere, out _))
                    {
                        leftOut.Add(file.NameThere);
                    }
                    else
                    {
                        copies.Add((path, file.NameHere));
                        cameFrom[file.NameHere] = file.NameThere;
                    }
                }

                return true;
            });

            bool cloned = false;

            void Change(IReadOnlyList<string> copied)
            {
                foreach (string name in copied)
                {
                    names[cameFrom[name]] = name;
                }

                leftOut.AddRange(cameFrom.Where(copy => !copied.Contains(copy.Key, StringComparer.OrdinalIgnoreCase)).Select(copy => copy.Value));

                // No mark of the pane's: the pane reads the component back, as after any other editor's change.
                cloned = session.Apply(
                    DraftStepWording.ComponentCloned(label, picked.Label, source.BoardName),
                    content => ComponentClone.Apply(content.Board, label, source.Board, picked.Label, region, names) is BoardData changed && !ReferenceEquals(changed, content.Board)
                        ? content with { Board = changed }
                        : null);
            }

            // None of its files could be copied: what cites them is left out and said, as when one of two
            // could not be - the rest of the component still comes (code review, 2026-11-04: it came to
            // nothing, with only "could not be copied" to say why).
            if (copies.Count == 0 || !await this.AddToDraftAsync(copies, Change))
            {
                Change([]);
            }

            this.CloseClonePanel();

            if (cloned)
            {
                this.ShowMessage(ComponentPaneWording.Cloned(picked.Label, source.BoardName, leftOut));
                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }

            return cloned;
        }

        private void OnCancelCloneClick(object? sender, RoutedEventArgs e) => this.CloseClonePanel();

        private void CloseClonePanel()
        {
            this.ClonePanel.IsVisible = false;
            this.thisCloneBoards = [];

            this.thisFillingClone = true;

            try
            {
                this.CloneHardwareBox.ItemsSource = null;
                this.CloneBoardBox.ItemsSource = null;
            }
            finally
            {
                this.thisFillingClone = false;
            }

            this.ForgetCloneSource();
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
                !this.thisIsNew && this.thisBoardPart is null && this.HoldsWorkOfItsOwn &&
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
            this.CloseClonePanel();

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
                ComponentCopySource here = new(string.Empty, session.Content.Board, ComponentPaneEdits.RowShown(session.Content.Board, this.thisLabel, this.thisRegion)!, RegionShown: this.thisRegion);

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

            string said = await apply(new ComponentCopySource(string.Empty, session.Content.Board, row, RegionShown: this.thisRegion), targets, parts, files);
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
                this.thisBoardPart == PaneBoardPart.Files && this.ChangedBoardFilesHere().Count > 0;
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

    // A board a component can be cloned from, as read: what it is called, its data, its list of files and
    // where the bytes of a file of it lie on this computer (null: not here).
    internal sealed record ComponentCloneSource(
        string BoardName,
        BoardData Board,
        IReadOnlyDictionary<string, BoardDocument.StoredFileRow>? Files,
        Func<string, string?> PathOf);

    // One component of that board in the panel's list - its title, by the popup's one rule.
    internal sealed record ComponentCloneChoice(string Label, string Title)
    {
        public override string ToString() => this.Title;
    }
}
