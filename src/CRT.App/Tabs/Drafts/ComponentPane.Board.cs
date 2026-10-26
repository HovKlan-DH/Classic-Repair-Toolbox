using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE BOARD'S OWN LISTS IN THE PANE (CRT 4.0 step 2, slice 4a, 2026-10-26 - the plan: "Board >
    // Files and links, grouped by category ... Board > Credits ... Board > Important signals: pick nets
    // from a searchable KiCad list").
    //
    // The pane shows ONE of them in place of a component (ShowBoardPart): the board's files and links,
    // its credits, or its important signals - and, since slice 4b, its pictures with their colours
    // (ComponentPane.BoardPictures.cs). They were reached only through their sheet of the table.
    //
    //   - EVERY CHANGE GOES INTO THE DRAFT AS A COMPONENT'S DOES - a box left, Enter, a typing pause,
    //     a row removed - each list one step (BoardPartEdits, the list replaced as a whole). Files,
    //     links and credits are listed GROUPED BY CATEGORY, as they are written; the category box
    //     offers the board's own categories, so a new row joins a group rather than starting a twin.
    //   - A FILE ADDED is copied in under a name of the board's own ("Board local files/<file>",
    //     ComponentPaneFileNames.ForBoardFile) and the draft saved at once, as a component's file is.
    //   - "ADD ME" adds a credit for the person at the keyboard, by the name and contact CRT remembers
    //     (CreditIdentity, the host's - the settings), remembered again once the row has a name.
    //   - AN IMPORTANT SIGNAL picks its net from the board's KiCad data (UseNetChoices - the host reads
    //     them off the UI thread), and one the Schematics tab's panel will not find says so under it
    //     (KiCadNetChoices, the panel's own rule). Without KiCad data any net may be typed.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // The board part shown in place of a component - null while a component (or its form) is.
        private PaneBoardPart? thisBoardPart;

        private readonly ObservableCollection<ComponentPaneBoardFileRow> thisBoardFiles = new();
        private readonly ObservableCollection<ComponentPaneBoardLinkRow> thisBoardLinks = new();
        private readonly ObservableCollection<ComponentPaneCreditRow> thisCredits = new();
        private readonly ObservableCollection<ComponentPaneSignalRow> thisSignals = new();

        private KiCadNetChoices? thisNetChoices;

        internal PaneBoardPart? BoardPart => this.thisBoardPart;

        internal IReadOnlyList<ComponentPaneBoardFileRow> BoardFilesForTests => this.thisBoardFiles;

        internal IReadOnlyList<ComponentPaneBoardLinkRow> BoardLinksForTests => this.thisBoardLinks;

        internal IReadOnlyList<ComponentPaneCreditRow> CreditsForTests => this.thisCredits;

        internal IReadOnlyList<ComponentPaneSignalRow> SignalsForTests => this.thisSignals;

        private void InitializeBoardParts()
        {
            this.BoardFilesItemsControl.ItemsSource = this.thisBoardFiles;
            this.BoardLinksItemsControl.ItemsSource = this.thisBoardLinks;
            this.CreditsItemsControl.ItemsSource = this.thisCredits;
            this.SignalsItemsControl.ItemsSource = this.thisSignals;
            this.InitializeBoardPictures();
        }

        // ###########################################################################################
        // Shows `part` of the draft held in `session`. `published` and its file as for Show - what the
        // "Changed" mark compares with. What the pane held for a component goes in first.
        // ###########################################################################################
        // `region`: the one CRT shows - the Pictures part marks its components (code review, 2026-10-26:
        // the part used whatever region the last component opened had, empty at first).
        internal void ShowBoardPart(DraftSession session, PaneBoardPart part, BoardData? published, string publishedBoardFile = "", string region = "")
        {
            ArgumentNullException.ThrowIfNull(session);

            if (this.thisSession is not null && !ReferenceEquals(session, this.thisSession))
            {
                this.Close();
            }
            else
            {
                this.CommitHeld();
            }

            this.UseSession(session);

            this.thisRegion = region?.Trim() ?? string.Empty;
            this.thisLabel = string.Empty;
            this.thisIsNew = false;
            this.thisBoardPart = part;
            this.CloseCopyPanel();
            this.thisPublished = published;
            this.thisPublishedBoardFile = publishedBoardFile ?? string.Empty;

            // The KiCad nets are read again for the part each time it opens (UseNetChoices).
            if (part == PaneBoardPart.ImportantSignals)
            {
                this.thisNetChoices = null;
            }

            this.ForgetComponentRows();
            this.ShowBoardPartRows();
        }

        // ###########################################################################################
        // The nets an important signal picks from, once the host has read them - each row offered
        // them, and told when its net is none of them.
        // ###########################################################################################
        internal void UseNetChoices(KiCadNetChoices choices)
        {
            this.thisNetChoices = choices ?? KiCadNetChoices.None;

            foreach (ComponentPaneSignalRow row in this.thisSignals)
            {
                row.UseChoices(this.thisNetChoices);
            }

            this.ShowNetsLine();
        }

        // The component's rows let go - a board part shows none of them, and none may be committed.
        private void ForgetComponentRows()
        {
            this.thisPins.Clear();
            this.thisReplacements.Clear();
            this.thisLinks.Clear();
            this.ForgetImagesAndFiles();
            this.thisShownValues.Clear();
        }

        // ###########################################################################################
        // The part read again from the draft as it is now - after it opens, and after any change but
        // the pane's own edit (ShowComponent sends it here).
        // ###########################################################################################
        private void ShowBoardPartRows()
        {
            if (this.thisSession is not { } session || this.thisBoardPart is not { } part)
            {
                return;
            }

            BoardData board = session.Content.Board;

            this.thisShowing = true;

            try
            {
                IReadOnlyList<string> fileCategories = BoardPartEdits.CategoriesOf(board.BoardLocalFiles, row => row.Category);
                IReadOnlyList<string> linkCategories = BoardPartEdits.CategoriesOf(board.BoardLinks, row => row.Category);
                IReadOnlyList<string> creditCategories = BoardPartEdits.CategoriesOf(board.Credits, row => row.Category);

                ComponentPane.Fill(
                    this.thisBoardFiles,
                    BoardPartEdits.Grouped(board.BoardLocalFiles, row => row.Category).Select(row => ComponentPaneBoardFileRow.From(row, fileCategories)));
                ComponentPane.Fill(
                    this.thisBoardLinks,
                    BoardPartEdits.Grouped(board.BoardLinks, row => row.Category).Select(row => ComponentPaneBoardLinkRow.From(row, linkCategories)));
                ComponentPane.Fill(
                    this.thisCredits,
                    BoardPartEdits.Grouped(board.Credits, row => row.Category).Select(row => ComponentPaneCreditRow.From(row, creditCategories)));
                ComponentPane.Fill(
                    this.thisSignals,
                    board.KiCadImportantSignals.Select(row => ComponentPaneSignalRow.From(row, this.thisNetChoices)));

                if (part == PaneBoardPart.Pictures)
                {
                    this.ShowBoardPictures(board);
                }
            }
            finally
            {
                this.thisShowing = false;
            }

            this.TitleText.Text = ComponentPaneWording.Title(part);
            this.RegionText.IsVisible = false;
            this.ActionsPanel.IsVisible = false;
            this.GonePanel.IsVisible = false;
            this.NewComponentPanel.IsVisible = false;
            this.SectionsScrollViewer.IsVisible = false;
            this.BoardPartPanel.IsVisible = true;

            this.FilesAndLinksPart.IsVisible = part == PaneBoardPart.FilesAndLinks;
            this.CreditsPart.IsVisible = part == PaneBoardPart.Credits;
            this.ImportantSignalsPart.IsVisible = part == PaneBoardPart.ImportantSignals;
            this.PicturesPart.IsVisible = part == PaneBoardPart.Pictures;

            this.ShowBoardPartLines();
            this.ShowBoardPartChanged();
        }

        // The headings with their counts, and the line for a list with nothing in it.
        private void ShowBoardPartLines()
        {
            this.BoardFilesHeading.Text = ComponentPaneWording.Heading("Files", this.thisBoardFiles.Count);
            this.NoBoardFilesText.IsVisible = this.thisBoardFiles.Count == 0;
            this.BoardFilesHeadings.IsVisible = this.thisBoardFiles.Count > 0;

            this.BoardLinksHeading.Text = ComponentPaneWording.Heading("Links", this.thisBoardLinks.Count);
            this.NoBoardLinksText.IsVisible = this.thisBoardLinks.Count == 0;
            this.BoardLinksHeadings.IsVisible = this.thisBoardLinks.Count > 0;

            this.CreditsHeading.Text = ComponentPaneWording.Heading("Credits", this.thisCredits.Count);
            this.NoCreditsText.IsVisible = this.thisCredits.Count == 0;
            this.CreditsHeadings.IsVisible = this.thisCredits.Count > 0;

            this.SignalsHeading.Text = ComponentPaneWording.Heading("Important signals", this.thisSignals.Count);
            this.NoSignalsText.IsVisible = this.thisSignals.Count == 0;
            this.SignalsHeadings.IsVisible = this.thisSignals.Count > 0;

            this.ShowNetsLine();
        }

        // Above the signals: the nets are being read, or the board has none to pick from.
        private void ShowNetsLine()
        {
            string line = this.thisNetChoices switch
            {
                null => ComponentPaneWording.ReadingKiCadNets,
                { HasNets: false } => ComponentPaneWording.NoKiCadNets,
                _ => string.Empty,
            };

            this.KiCadNetsLineText.Text = line;
            this.KiCadNetsLineText.IsVisible = line.Length > 0;
        }

        // "Changed" while the draft's list differs from the published board's - by the edits' own rule,
        // so a list in another order is no change.
        internal bool IsBoardPartChanged =>
            this.thisSession is { } session && this.thisBoardPart is { } part &&
            ComponentPane.BoardPartDiffers(part, session.Content.Board, this.thisPublished);

        private static bool BoardPartDiffers(PaneBoardPart part, BoardData board, BoardData? published)
        {
            BoardData baseline = published ?? new BoardData();

            return part switch
            {
                PaneBoardPart.FilesAndLinks =>
                    !ReferenceEquals(BoardPartEdits.WithBoardFiles(baseline, board.BoardLocalFiles), baseline) ||
                    !ReferenceEquals(BoardPartEdits.WithBoardLinks(baseline, board.BoardLinks), baseline),
                PaneBoardPart.Credits => !ReferenceEquals(BoardPartEdits.WithCredits(baseline, board.Credits), baseline),
                PaneBoardPart.Pictures =>
                    !baseline.Schematics.Select(ComponentPane.PictureKey).SequenceEqual(board.Schematics.Select(ComponentPane.PictureKey), StringComparer.OrdinalIgnoreCase) ||
                    !ReferenceEquals(
                        BoardPartEdits.WithPictureColours(baseline, board.Schematics.Select(row => (row.SchematicName, row.SchematicImageFile, row.ReadColours())).ToList()),
                        baseline),
                _ => !ReferenceEquals(BoardPartEdits.WithImportantSignals(baseline, board.KiCadImportantSignals), baseline),
            };
        }

        // A picture as the board lists it - its name and its image file.
        private static string PictureKey(BoardSchematicEntry row) =>
            string.Join('\u001F', row.SchematicName?.Trim(), row.SchematicImageFile?.Trim());

        private void ShowBoardPartChanged()
        {
            this.ChangedMark.IsVisible = this.IsBoardPartChanged;
            ToolTip.SetTip(this.ChangedMark, ComponentPaneWording.BoardPartChangedTip);
        }

        // ---- the part's changes into the draft ----------------------------------------------------

        // Everything the part's boxes hold that the draft does not have yet - each list one step.
        private void CommitBoardPart()
        {
            switch (this.thisBoardPart)
            {
                case PaneBoardPart.FilesAndLinks:
                    this.CommitBoardFiles();
                    this.CommitBoardLinks();
                    break;
                case PaneBoardPart.Credits:
                    this.CommitCredits();
                    break;
                case PaneBoardPart.ImportantSignals:
                    this.CommitSignals();
                    break;
                case PaneBoardPart.Pictures:
                    this.CommitPictureColours();
                    break;
            }
        }

        private void CommitBoardFiles()
        {
            List<BoardLocalFileEntry> rows = this.thisBoardFiles.Select(row => row.ToEntry()).ToList();
            this.ApplyBoardPart(DraftStepWording.BoardFiles, board => BoardPartEdits.WithBoardFiles(board, rows));
        }

        private void CommitBoardLinks()
        {
            List<BoardLinkEntry> rows = this.thisBoardLinks.Select(row => row.ToEntry()).ToList();
            this.ApplyBoardPart(DraftStepWording.BoardLinks, board => BoardPartEdits.WithBoardLinks(board, rows));
        }

        private void CommitCredits()
        {
            List<CreditEntry> rows = this.thisCredits.Select(row => row.ToEntry()).ToList();
            this.ApplyBoardPart(DraftStepWording.Credits, board => BoardPartEdits.WithCredits(board, rows));
            this.RememberMine();
        }

        private void CommitSignals()
        {
            List<KiCadImportantSignalEntry> rows = this.thisSignals.Select(row => row.ToEntry()).ToList();
            this.ApplyBoardPart(DraftStepWording.ImportantSignals, board => BoardPartEdits.WithImportantSignals(board, rows));

            // Each row's net judged again as it now reads.
            foreach (ComponentPaneSignalRow row in this.thisSignals)
            {
                row.UseChoices(this.thisNetChoices);
            }
        }

        // One change of a board part, as the pane's own step - "Changed" follows.
        private void ApplyBoardPart(string description, Func<BoardData, BoardData> transform)
        {
            if (this.thisSession is not { } session)
            {
                return;
            }

            bool changed = session.Apply(
                description,
                content => transform(content.Board) is BoardData board && !ReferenceEquals(board, content.Board)
                    ? content with { Board = board }
                    : null,
                this.thisOrigin);

            if (changed)
            {
                this.ShowBoardPartChanged();
                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        // ---- files and links ----------------------------------------------------------------------

        private async void OnAddBoardFileClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AddBoardFileAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Adding a file of the board in the component pane failed - [{ex}]");
            }
        }

        // A picked file into the draft under a name of the board's own, its row in the last category
        // listed (so it joins a group), the draft saved at once.
        internal async Task AddBoardFileAsync()
        {
            this.CommitHeld();
            this.ShowMessage(null);

            if (await this.PickFileAsync(PaneFileKind.BoardFile) is not { } path)
            {
                return;
            }

            string name = ComponentPaneFileNames.ForBoardFile(path, this.TakenNames());
            IReadOnlyList<string> categories = BoardPartEdits.CategoriesOf(this.thisBoardFiles, row => row.Category);

            await this.AddToDraftAsync(path, name, () =>
            {
                this.thisBoardFiles.Add(new ComponentPaneBoardFileRow
                {
                    Category = this.thisBoardFiles.LastOrDefault()?.Category ?? string.Empty,
                    Name = Path.GetFileNameWithoutExtension(path),
                    File = name,
                    CategoryChoices = categories,
                });

                this.CommitBoardFiles();
                this.ShowBoardPartLines();
            });
        }

        private async void OnChangeBoardFileClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Control { Tag: ComponentPaneBoardFileRow row })
                {
                    await this.ChangeBoardFileAsync(row);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Changing a file of the board in the component pane failed - [{ex}]");
            }
        }

        internal async Task ChangeBoardFileAsync(ComponentPaneBoardFileRow row)
        {
            this.CommitHeld();
            this.ShowMessage(null);

            if (await this.PickFileAsync(PaneFileKind.BoardFile) is not { } path)
            {
                return;
            }

            string name = ComponentPaneFileNames.ForBoardFile(path, this.TakenNames());

            await this.AddToDraftAsync(path, name, () =>
            {
                row.File = name;
                this.CommitBoardFiles();
            });
        }

        private void OnRemoveBoardFileClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneBoardFileRow row } && this.thisBoardFiles.Remove(row))
            {
                this.CommitBoardFiles();
                this.ShowBoardPartLines();
            }
        }

        private void OnAddBoardLinkClick(object? sender, RoutedEventArgs e)
        {
            this.thisBoardLinks.Add(new ComponentPaneBoardLinkRow
            {
                Category = this.thisBoardLinks.LastOrDefault()?.Category ?? string.Empty,
                CategoryChoices = BoardPartEdits.CategoriesOf(this.thisBoardLinks, row => row.Category),
            });

            this.ShowBoardPartLines();
        }

        private void OnRemoveBoardLinkClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneBoardLinkRow row } && this.thisBoardLinks.Remove(row))
            {
                this.CommitBoardLinks();
                this.ShowBoardPartLines();
            }
        }

        // ---- credits ------------------------------------------------------------------------------

        private void OnAddCreditClick(object? sender, RoutedEventArgs e)
        {
            this.thisCredits.Add(new ComponentPaneCreditRow
            {
                CategoryChoices = BoardPartEdits.CategoriesOf(this.thisCredits, row => row.Category),
            });

            this.ShowBoardPartLines();
        }

        // ###########################################################################################
        // "Add me" (owner decision, 2026-10-26: "Yes, let us try that"): a credit for the person at the
        // keyboard, by the name and contact CRT remembers (CreditIdentity - the host's; never the
        // contact email, since credits are shown to everybody). With none remembered yet the row starts
        // empty and says so; once it goes into the draft with a name, that name is remembered.
        // ###########################################################################################
        internal Func<(string Name, string Contact)>? CreditIdentity { get; set; }

        internal Action<string, string>? RememberCreditIdentity { get; set; }

        private void OnAddMeClick(object? sender, RoutedEventArgs e) => this.AddMe();

        internal void AddMe()
        {
            (string name, string contact) = this.CreditIdentity?.Invoke() ?? (string.Empty, string.Empty);

            this.thisCredits.Add(new ComponentPaneCreditRow
            {
                NameOrHandle = name,
                Contact = contact,
                IsMine = true,
                CategoryChoices = BoardPartEdits.CategoriesOf(this.thisCredits, row => row.Category),
            });

            this.CreditMeHintText.IsVisible = string.IsNullOrWhiteSpace(name);
            this.ShowBoardPartLines();
        }

        // The name and contact of a credit added with "Add me", once it is in the draft with a name.
        private void RememberMine()
        {
            if (this.thisCredits.FirstOrDefault(row => row.IsMine && row.NameOrHandle.Trim().Length > 0) is { } mine)
            {
                this.RememberCreditIdentity?.Invoke(mine.NameOrHandle.Trim(), mine.Contact.Trim());
                this.CreditMeHintText.IsVisible = false;
            }
        }

        private void OnRemoveCreditClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneCreditRow row } && this.thisCredits.Remove(row))
            {
                this.CommitCredits();
                this.ShowBoardPartLines();
            }
        }

        // ---- important signals --------------------------------------------------------------------

        private void OnAddSignalClick(object? sender, RoutedEventArgs e)
        {
            var row = new ComponentPaneSignalRow();
            row.UseChoices(this.thisNetChoices);
            this.thisSignals.Add(row);

            this.ShowBoardPartLines();
        }

        private void OnRemoveSignalClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneSignalRow row } && this.thisSignals.Remove(row))
            {
                this.CommitSignals();
                this.ShowBoardPartLines();
            }
        }
    }

    // ###########################################################################################
    // The rows of the board parts, as the pane holds them. A category box offers the board's own
    // categories of that list (CategoryChoices, set when the part is shown).
    // ###########################################################################################
    public abstract class ComponentPaneRowBase : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        // The value set, and said - false when it was already so.
        protected bool Set(ref string field, string? value, [CallerMemberName] string? name = null)
        {
            string wanted = value ?? string.Empty;

            if (string.Equals(field, wanted, StringComparison.Ordinal))
            {
                return false;
            }

            field = wanted;
            this.Raise(name);
            return true;
        }

        protected void Raise(string? name) => this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    public sealed class ComponentPaneBoardFileRow : ComponentPaneRowBase
    {
        private string thisCategory = string.Empty;
        private string thisName = string.Empty;
        private string thisFile = string.Empty;

        public string Category { get => this.thisCategory; set => this.Set(ref this.thisCategory, value); }

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value); }

        public string File { get => this.thisFile; set => this.Set(ref this.thisFile, value); }

        public IReadOnlyList<string> CategoryChoices { get; init; } = [];

        public static ComponentPaneBoardFileRow From(BoardLocalFileEntry row, IReadOnlyList<string> categories) => new()
        {
            Category = row.Category ?? string.Empty,
            Name = row.Name ?? string.Empty,
            File = row.File ?? string.Empty,
            CategoryChoices = categories,
        };

        public BoardLocalFileEntry ToEntry() => new() { Category = this.thisCategory, Name = this.thisName, File = this.thisFile };
    }

    public sealed class ComponentPaneBoardLinkRow : ComponentPaneRowBase
    {
        private string thisCategory = string.Empty;
        private string thisName = string.Empty;
        private string thisUrl = string.Empty;

        public string Category { get => this.thisCategory; set => this.Set(ref this.thisCategory, value); }

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value); }

        public string Url { get => this.thisUrl; set => this.Set(ref this.thisUrl, value); }

        public IReadOnlyList<string> CategoryChoices { get; init; } = [];

        public static ComponentPaneBoardLinkRow From(BoardLinkEntry row, IReadOnlyList<string> categories) => new()
        {
            Category = row.Category ?? string.Empty,
            Name = row.Name ?? string.Empty,
            Url = row.Url ?? string.Empty,
            CategoryChoices = categories,
        };

        public BoardLinkEntry ToEntry() => new() { Category = this.thisCategory, Name = this.thisName, Url = this.thisUrl };
    }

    public sealed class ComponentPaneCreditRow : ComponentPaneRowBase
    {
        private string thisCategory = string.Empty;
        private string thisSubCategory = string.Empty;
        private string thisNameOrHandle = string.Empty;
        private string thisContact = string.Empty;

        public string Category { get => this.thisCategory; set => this.Set(ref this.thisCategory, value); }

        public string SubCategory { get => this.thisSubCategory; set => this.Set(ref this.thisSubCategory, value); }

        public string NameOrHandle { get => this.thisNameOrHandle; set => this.Set(ref this.thisNameOrHandle, value); }

        public string Contact { get => this.thisContact; set => this.Set(ref this.thisContact, value); }

        public IReadOnlyList<string> CategoryChoices { get; init; } = [];

        // Added with "Add me" - its name and contact are the person's own, remembered once filled.
        public bool IsMine { get; init; }

        public static ComponentPaneCreditRow From(CreditEntry row, IReadOnlyList<string> categories) => new()
        {
            Category = row.Category ?? string.Empty,
            SubCategory = row.SubCategory ?? string.Empty,
            NameOrHandle = row.NameOrHandle ?? string.Empty,
            Contact = row.Contact ?? string.Empty,
            CategoryChoices = categories,
        };

        public CreditEntry ToEntry() => new()
        {
            Category = this.thisCategory,
            SubCategory = this.thisSubCategory,
            NameOrHandle = this.thisNameOrHandle,
            Contact = this.thisContact,
        };
    }

    // One important signal - its net picked from the board's KiCad nets, and what is wrong with it.
    public sealed class ComponentPaneSignalRow : ComponentPaneRowBase
    {
        private string thisDisplayName = string.Empty;
        private string thisNetName = string.Empty;
        private string thisProblem = string.Empty;
        private IReadOnlyList<string> thisNetChoices = [];

        public string DisplayName { get => this.thisDisplayName; set => this.Set(ref this.thisDisplayName, value); }

        public string KiCadNetName { get => this.thisNetName; set => this.Set(ref this.thisNetName, value); }

        public IReadOnlyList<string> NetChoices
        {
            get => this.thisNetChoices;
            private set
            {
                this.thisNetChoices = value;
                this.Raise(nameof(this.NetChoices));
            }
        }

        public string Problem
        {
            get => this.thisProblem;
            private set
            {
                if (this.Set(ref this.thisProblem, value))
                {
                    this.Raise(nameof(this.HasProblem));
                }
            }
        }

        public bool HasProblem => this.thisProblem.Length > 0;

        // The nets to pick from, and whether this row's net is one of them - nothing said while the
        // nets are still being read (null).
        public void UseChoices(KiCadNetChoices? choices)
        {
            IReadOnlyList<string> names = choices?.Names ?? [];

            if (!ReferenceEquals(this.thisNetChoices, names))
            {
                this.NetChoices = names;
            }

            this.Problem = choices is null || choices.Finds(this.thisNetName)
                ? string.Empty
                : ComponentPaneWording.NetNotFound(this.thisNetName);
        }

        public static ComponentPaneSignalRow From(KiCadImportantSignalEntry row, KiCadNetChoices? choices)
        {
            var shown = new ComponentPaneSignalRow
            {
                DisplayName = row.DisplayName ?? string.Empty,
                KiCadNetName = row.KiCadNetName ?? string.Empty,
            };

            shown.UseChoices(choices);
            return shown;
        }

        public KiCadImportantSignalEntry ToEntry() => new() { DisplayName = this.thisDisplayName, KiCadNetName = this.thisNetName };
    }
}
