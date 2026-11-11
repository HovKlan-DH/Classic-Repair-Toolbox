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
        }
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
