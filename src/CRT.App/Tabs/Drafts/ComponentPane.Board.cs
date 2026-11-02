using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
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
    // The pane shows ONE of them in place of a component (ShowBoardPart): the board's files, its links
    // (one part, "Files and links", until 2026-10-29), its credits, or its important signals - and,
    // since slice 4b, its schematic images with their colours (ComponentPane.BoardPictures.cs). They
    // were reached only through their sheet of the table.
    //
    //   - EVERY CHANGE GOES INTO THE DRAFT AS A COMPONENT'S DOES - a box left, Enter, a typing pause,
    //     a row removed - each list one step (BoardPartEdits, the list replaced as a whole). A list is
    //     shown and written in the board's own order (grouped by category until 2026-10-30); the
    //     category box offers the board's own categories, so a new row joins one rather than starting
    //     a twin of it.
    //   - A FILE ADDED is copied in under a name of the board's own ("Board local files/<file>",
    //     ComponentPaneFileNames.ForBoardFile) and the draft saved at once, as a component's file is.
    //   - A FILE OR A LINK IS OPENED from its row, and A FILE, A LINK, A CREDIT OR AN IMPORTANT SIGNAL
    //     MOVES up or down, whatever its category (owner requests, 2026-10-29 and 2026-10-30: "the
    //     entire line moving up/down ... I also want to be able to rearrange the categories") -
    //     DRAGGED by its handle, a dashed placeholder where it will land (ListRowDrag; arrows until
    //     2026-10-30 - owner request: "a North/South icon and then drag the whole line up/down, just
    //     alike in the workbook for photos"), or Alt+Up / Alt+Down in its boxes (BoardPartEdits
    //     .CanMove/Moved). The new order is a change Submit sends (BoardRowOrder - owner decision,
    //     2026-10-30), and the part's "Draft" mark says it.
    //   - A FIELD A ROW NEEDS, LEFT EMPTY, IS MARKED ON ITS BOX (owner requests, 2026-10-30 and
    //     2026-11-01: "highlight the problematic cell with a helper what is missing ... These errors must
    //     be flagged in 'Board status' and then it should not be possible to submit"): a red frame and
    //     "Required before submitting" under it (BoardRowRequired, the one rule). The row goes into the
    //     draft as it is - an error there, said by the table, the draft's row and Board status, and
    //     Submit sends nothing while it is. Until then such a row was held back in the pane, a line
    //     under it, and nothing but the pane knew. A row with nothing typed in it is no row.
    //   - AN IMPORTANT SIGNAL's net is the KiCad data's, shown as text and never typed (owner request,
    //     2026-10-29): a new signal PICKS it from the board's nets no row has (UseNetChoices - the host
    //     reads them off the UI thread; ImportantSignalSuggestions.NetsToAdd), and one the Schematics
    //     tab's panel will not find says so under it (KiCadNetChoices, the panel's own rule). Without
    //     KiCad data there is nothing to pick.
    //   - THE PART'S ONE HEADLINE is the pane's title, with its count (PartTitle) - no heading of its
    //     own under it (owner request, 2026-10-29). "Add me" went the same day (owner decision).
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

        // Each list's drag - made once, by InitializeBoardParts.
        private ListRowDrag<ComponentPaneBoardFileRow>? thisBoardFilesDrag;
        private ListRowDrag<ComponentPaneBoardLinkRow>? thisBoardLinksDrag;
        private ListRowDrag<ComponentPaneCreditRow>? thisCreditsDrag;
        private ListRowDrag<ComponentPaneSignalRow>? thisSignalsDrag;
        private ListRowDrag<ComponentPaneCategoryRow>? thisCategoriesDrag;

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
            this.CategoriesItemsControl.ItemsSource = this.thisCategories;
            this.InitializeBoardPictures();
            this.InitializeSignalSuggestions();
            this.NewComponentMarkNoteText.Text = ComponentPaneWording.NewComponentMarkNote;

            // A row dropped somewhere else is the list in its new order, into the draft as one step.
            this.thisBoardFilesDrag = new ListRowDrag<ComponentPaneBoardFileRow>(
                this, this.BoardFilesItemsControl, this.BoardPartPanel, this.thisBoardFiles, (_, _, _) => this.CommitBoardFiles());
            this.thisBoardLinksDrag = new ListRowDrag<ComponentPaneBoardLinkRow>(
                this, this.BoardLinksItemsControl, this.BoardPartPanel, this.thisBoardLinks, (_, _, _) => this.CommitBoardLinks());
            this.thisCreditsDrag = new ListRowDrag<ComponentPaneCreditRow>(
                this, this.CreditsItemsControl, this.BoardPartPanel, this.thisCredits, (_, _, _) => this.CommitCredits());
            this.thisSignalsDrag = new ListRowDrag<ComponentPaneSignalRow>(
                this, this.SignalsItemsControl, this.BoardPartPanel, this.thisSignals, (_, _, _) => this.CommitSignals());
            this.thisCategoriesDrag = new ListRowDrag<ComponentPaneCategoryRow>(
                this, this.CategoriesItemsControl, this.BoardPartPanel, this.thisCategories, (_, _, _) => this.CommitCategoryOrder());

            // Alt+Up and Alt+Down in a row's boxes move it - ahead of the boxes, which would take the keys.
            this.BoardPartPanel.AddHandler(InputElement.KeyDownEvent, this.OnRowKeyDown, RoutingStrategies.Tunnel);
            this.SectionsScrollViewer.AddHandler(InputElement.KeyDownEvent, this.OnRowKeyDown, RoutingStrategies.Tunnel);
        }

        // Every drag ended - the rows are about to be replaced, and a drag must never outlive its row.
        private void ResetRowDrags()
        {
            this.thisBoardFilesDrag?.Reset();
            this.thisBoardLinksDrag?.Reset();
            this.thisCreditsDrag?.Reset();
            this.thisSignalsDrag?.Reset();
            this.thisCategoriesDrag?.Reset();
            this.thisFilesDrag?.Reset();
            this.thisLinksDrag?.Reset();
        }

        internal bool IsDraggingARowForTests =>
            this.thisBoardFilesDrag?.IsDragging == true || this.thisBoardLinksDrag?.IsDragging == true ||
            this.thisCreditsDrag?.IsDragging == true || this.thisSignalsDrag?.IsDragging == true ||
            this.thisCategoriesDrag?.IsDragging == true ||
            this.thisFilesDrag?.IsDragging == true || this.thisLinksDrag?.IsDragging == true;

        // ###########################################################################################
        // Shows `part` of the draft held in `session`. `published` and its file as for Show - what the
        // "Draft" mark compares with. What the pane held for a component goes in first.
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
            this.ForgetKiCadPins();
            this.thisPublished = published;
            this.thisPublishedBoardFile = publishedBoardFile ?? string.Empty;

            // The KiCad nets are read again for the part each time it opens (UseNetChoices), and so are
            // the KiCad views the pictures pick from (UseKiCadViews).
            if (part == PaneBoardPart.ImportantSignals)
            {
                this.thisNetChoices = null;
            }

            if (part == PaneBoardPart.Pictures)
            {
                this.thisCadViewNames = null;
                this.thisCadBoardSides = null;
                this.thisCadViewComponents = null;

                // With the others (code review, 2026-11-02): left, a card on another board with the same
                // view names counted its marks by the previous board's footprints until its read answered.
                this.thisCadViewFootprints = null;

                // The part opens fresh: no image's traces ticked, none decoded sharp.
                this.ForgetPreviews();
            }

            // What "Mark all automatically" said was of the part as it was open before - the part read
            // again after its marks keeps it (ShowBoardPartRows).
            this.ShowPicturesNote(null);

            this.ForgetComponentRows();
            this.ShowBoardPartRows();

            if (part == PaneBoardPart.Pictures)
            {
                this.AskForKiCadViews();
            }
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
            this.ShowSignalSuggestions();
            this.ShowNewSignalPicker();
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
            this.ResetRowDrags();

            try
            {
                IReadOnlyList<string> fileCategories = BoardPartEdits.CategoriesOf(board.BoardLocalFiles, row => row.Category);
                IReadOnlyList<string> linkCategories = BoardPartEdits.CategoriesOf(board.BoardLinks, row => row.Category);
                IReadOnlyList<string> creditCategories = BoardPartEdits.CategoriesOf(board.Credits, row => row.Category);

                ComponentPane.Fill(
                    this.thisBoardFiles,
                    board.BoardLocalFiles.Select(row => ComponentPaneBoardFileRow.From(row, fileCategories)));
                ComponentPane.Fill(
                    this.thisBoardLinks,
                    board.BoardLinks.Select(row => ComponentPaneBoardLinkRow.From(row, linkCategories)));
                ComponentPane.Fill(
                    this.thisCredits,
                    board.Credits.Select(row => ComponentPaneCreditRow.From(row, creditCategories)));
                ComponentPane.Fill(
                    this.thisSignals,
                    board.KiCadImportantSignals.Select(row => ComponentPaneSignalRow.From(row, this.thisNetChoices)));

                this.UseBoardRegions(board);

                if (part == PaneBoardPart.Pictures)
                {
                    this.DropPicturesNoteIfMovedOn();
                    this.ShowBoardPictures(board);
                }

                if (part == PaneBoardPart.Regions)
                {
                    this.ShowRegionsPart(board);
                }

                if (part == PaneBoardPart.Categories)
                {
                    this.ShowCategoriesPart(board);
                }
            }
            finally
            {
                this.thisShowing = false;
            }

            this.RegionText.IsVisible = false;
            this.ActionsPanel.IsVisible = false;
            this.GonePanel.IsVisible = false;
            this.NewComponentPanel.IsVisible = false;
            this.SectionsScrollViewer.IsVisible = false;
            this.BoardPartPanel.IsVisible = true;
            this.ShowPageTabs();

            this.RegionsPart.IsVisible = part == PaneBoardPart.Regions;
            this.CategoriesPart.IsVisible = part == PaneBoardPart.Categories;
            this.BoardFilesPart.IsVisible = part == PaneBoardPart.Files;
            this.BoardLinksPart.IsVisible = part == PaneBoardPart.Links;
            this.CreditsPart.IsVisible = part == PaneBoardPart.Credits;
            this.ImportantSignalsPart.IsVisible = part == PaneBoardPart.ImportantSignals;
            this.PicturesPart.IsVisible = part == PaneBoardPart.Pictures;

            // The lists keep a readable width; the schematic images take the room there is - capped, a
            // wide preview took its card's settings down to nothing (owner report, 2026-10-30).
            ((Grid)this.BoardPartPanel.Content!).ColumnDefinitions[0].MaxWidth =
                part == PaneBoardPart.Pictures ? double.PositiveInfinity : ComponentPane.ReadableWidth;

            this.ShowBoardPartLines();
            this.ShowBoardPartChanged();
            this.ShowSignalSuggestions();
        }

        // The title with the part's count, the line for a list with nothing in it, which rows can move,
        // and what a new signal may pick.
        private void ShowBoardPartLines()
        {
            this.ShowBoardPartTitle();

            this.NoBoardFilesText.IsVisible = this.thisBoardFiles.Count == 0;
            this.BoardFilesHeadings.IsVisible = this.thisBoardFiles.Count > 0;

            this.NoBoardLinksText.IsVisible = this.thisBoardLinks.Count == 0;
            this.BoardLinksHeadings.IsVisible = this.thisBoardLinks.Count > 0;

            this.NoCreditsText.IsVisible = this.thisCredits.Count == 0;
            this.CreditsHeadings.IsVisible = this.thisCredits.Count > 0;

            this.NoSignalsText.IsVisible = this.thisSignals.Count == 0;
            this.SignalsHeadings.IsVisible = this.thisSignals.Count > 0;

            // What each file and link applies to, on a board with regions (owner request, 2026-11-02) - a
            // list of the whole board, so every row shows whatever region CRT shows.
            foreach (ComponentPaneRegionalRow row in this.thisBoardFiles.Cast<ComponentPaneRegionalRow>().Concat(this.thisBoardLinks))
            {
                row.ShowsAppliesTo = this.thisBoardHasRegions;
            }

            this.BoardFilesAppliesToHeading.IsVisible = this.thisBoardHasRegions;
            this.BoardLinksAppliesToHeading.IsVisible = this.thisBoardHasRegions;

            this.ShowNetsLine();
            this.ShowNewSignalPicker();
        }

        // The pane's title over a board part: its name and how many it holds - the part's one headline.
        // A row not in the draft - one with nothing typed in it yet - is not counted.
        private void ShowBoardPartTitle()
        {
            if (this.thisBoardPart is not { } part)
            {
                return;
            }

            // A page of settings has no count of its own - the tick says it.
            if (part == PaneBoardPart.Regions)
            {
                this.TitleText.Text = ComponentPaneWording.Title(part);
                return;
            }

            int count = part switch
            {
                PaneBoardPart.Categories => this.thisCategories.Count(row => !row.NeedsOne),
                PaneBoardPart.Files => this.thisBoardFiles.Count(row => row.Committed is not null),
                PaneBoardPart.Links => this.thisBoardLinks.Count(row => row.Committed is not null),
                PaneBoardPart.Credits => this.thisCredits.Count(row => row.Committed is not null),
                PaneBoardPart.Pictures => this.thisBoardPictures.Count,
                _ => this.thisSignals.Count(row => row.Committed is not null),
            };

            this.TitleText.Text = ComponentPaneWording.PartTitle(part, count);
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

        // ###########################################################################################
        // "Draft" while the draft's list differs from the published board's - by the rules Submit
        // counts by (code review, 2026-10-30): BoardDataDiffer's for the rows, and BoardRowOrder's for
        // their order, which is a change since the owner's decision the same day ("a change of (any)
        // order should also be submitted as a change"). A new board has nothing published to differ from.
        // ###########################################################################################
        internal bool IsBoardPartChanged =>
            this.thisSession is { } session && this.thisBoardPart is { } part && this.thisPublished is not null &&
            ComponentPane.BoardPartDiffers(part, session.Content.Board, this.thisPublished);

        internal static bool BoardPartDiffers(PaneBoardPart part, BoardData board, BoardData? published) => part switch
        {
            PaneBoardPart.Files => ComponentPane.ListDiffers(published?.BoardLocalFiles, board.BoardLocalFiles),
            PaneBoardPart.Links => ComponentPane.ListDiffers(published?.BoardLinks, board.BoardLinks),
            PaneBoardPart.Credits => ComponentPane.ListDiffers(published?.Credits, board.Credits),
            PaneBoardPart.Pictures => ComponentPane.ListDiffers(published?.Schematics, board.Schematics),
            PaneBoardPart.Regions => BoardRegions.Differ(published, board),
            PaneBoardPart.Categories => !ComponentPane.SameCategories(published, board),
            _ => ComponentPane.ListDiffers(published?.KiCadImportantSignals, board.KiCadImportantSignals),
        };

        // ###########################################################################################
        // Whether two boards sort their components under the same words, in the same order (order is a
        // change - owner decision, 2026-10-30). The names alone (code review, 2026-11-01): how many carry
        // each changes with any component added.
        // ###########################################################################################
        private static bool SameCategories(BoardData? published, BoardData board) =>
            (published is null ? [] : BoardCategoryEdits.Of(published)).Select(row => row.Name)
                .SequenceEqual(BoardCategoryEdits.Of(board).Select(row => row.Name), StringComparer.Ordinal);

        private static bool ListDiffers<T>(IReadOnlyList<T>? published, IReadOnlyList<T> draft)
            where T : class =>
            BoardDataDiffer.SectionDiffers(published, draft) || (published is not null && BoardRowOrder.Reordered(published, draft));

        private void ShowBoardPartChanged()
        {
            this.ChangedMark.IsVisible = this.IsBoardPartChanged;
            ToolTip.SetTip(this.ChangedMark, ComponentPaneWording.BoardPartChangedTip);
            this.ShowBoardFilesAlsoOffer();
        }

        // ---- the part's changes into the draft ----------------------------------------------------

        // Everything the part's boxes hold that the draft does not have yet - each list one step. A
        // picture's name never on a typing pause (`fromTyping` - see CommitHeld).
        private void CommitBoardPart(bool fromTyping)
        {
            switch (this.thisBoardPart)
            {
                case PaneBoardPart.Files:
                    this.CommitBoardFiles();
                    break;
                case PaneBoardPart.Links:
                    this.CommitBoardLinks();
                    break;
                case PaneBoardPart.Credits:
                    this.CommitCredits();
                    break;
                case PaneBoardPart.ImportantSignals:
                    this.CommitSignals();
                    break;
                case PaneBoardPart.Pictures:
                    // A name typed first - the colours are filed by the pictures' names.
                    if (!fromTyping)
                    {
                        this.CommitPictureNames();
                    }

                    this.CommitPictureColours();
                    break;
                case PaneBoardPart.Categories:
                    // A name never on a typing pause, as a picture's.
                    if (!fromTyping)
                    {
                        this.CommitCategoryNames();
                    }

                    break;
            }
        }

        private void CommitBoardFiles() =>
            this.CommitList(this.thisBoardFiles, DraftStepWording.BoardFiles, BoardPartEdits.WithBoardFiles);

        private void CommitBoardLinks() =>
            this.CommitList(this.thisBoardLinks, DraftStepWording.BoardLinks, BoardPartEdits.WithBoardLinks);

        private void CommitCredits() =>
            this.CommitList(this.thisCredits, DraftStepWording.Credits, BoardPartEdits.WithCredits);

        private void CommitSignals()
        {
            this.CommitList(this.thisSignals, DraftStepWording.ImportantSignals, BoardPartEdits.WithImportantSignals);

            // Each row's net judged again as it now reads, and what is suggested follows.
            foreach (ComponentPaneSignalRow row in this.thisSignals)
            {
                row.UseChoices(this.thisNetChoices);
            }

            this.ShowSignalSuggestions();
            this.ShowNewSignalPicker();
        }

        // ###########################################################################################
        // A list's rows into the draft as they read, as one step - a field a row needs left empty
        // included: an error of the draft's then, its box marked (owner request, 2026-11-01 - it was held
        // back until then). A row with nothing typed in it is no row (the table drops a blank row too):
        // it waits in the pane, and one emptied of everything leaves the draft.
        // ###########################################################################################
        private void CommitList<TEntry>(
            IReadOnlyList<ComponentPaneListRow<TEntry>> rows,
            string description,
            Func<BoardData, IReadOnlyList<TEntry>, BoardData> with)
            where TEntry : class
        {
            var entries = new List<TEntry>(rows.Count);

            foreach (ComponentPaneListRow<TEntry> row in rows)
            {
                if (row.IsBlank)
                {
                    row.Committed = null;
                    continue;
                }

                TEntry wanted = row.ToEntry();
                row.Committed = wanted;
                entries.Add(wanted);
            }

            this.ApplyBoardPart(description, board => with(board, entries));
        }

        // One change of a board part, as the pane's own step - "Draft" follows.
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
                this.ShowBoardPartTitle();
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
        // listed (so it joins a group) - every field filled in, BoardPartRequired.NewFile - the draft
        // saved at once.
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
                BoardLocalFileEntry added = BoardPartRequired.NewFile(path, name, this.thisBoardFiles.LastOrDefault()?.Category);

                this.thisBoardFiles.Add(new ComponentPaneBoardFileRow
                {
                    Category = added.Category,
                    Name = added.Name,
                    File = added.File,
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

        // "Open" (owner request, 2026-10-29: "a possibility to open the file so I can view it") - found
        // as every file of the draft is, a "Scope measurements" note's way (OpenFileOfBoard).
        private void OnOpenBoardFileClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneBoardFileRow row })
            {
                this.OpenFileOfBoard(row.File, row.Name.Trim().Length > 0 ? row.Name : row.File);
            }
        }

        // A row's handle pressed - the drag starts once the pointer moves (ListRowDrag.Press).
        private void OnBoardFileHandlePressed(object? sender, PointerPressedEventArgs e) => this.thisBoardFilesDrag?.Press(sender, e);

        internal bool MoveBoardFile(ComponentPaneBoardFileRow row, int by)
        {
            if (!this.MoveRow(this.thisBoardFiles, row, by))
            {
                return false;
            }

            this.CommitBoardFiles();
            return true;
        }

        private void OnAddBoardLinkClick(object? sender, RoutedEventArgs e)
        {
            this.thisBoardLinks.Add(new ComponentPaneBoardLinkRow
            {
                Category = this.thisBoardLinks.LastOrDefault()?.Category ?? string.Empty,
                OfferedCategory = this.thisBoardLinks.LastOrDefault()?.Category ?? string.Empty,
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

        // For tests, in place of opening a web page in the browser.
        internal Action<string>? OpenLinkOverrideForTests { get; set; }

        // "Open" (owner request, 2026-10-29: "a possibility to open a link in a browser") - the address
        // by LinkTarget's rule, the PDF export's own: one typed without "https://" gets it.
        private void OnOpenBoardLinkClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneBoardLinkRow row })
            {
                this.OpenBoardLink(row);
            }
        }

        internal void OpenBoardLink(ComponentPaneBoardLinkRow row)
        {
            this.ShowMessage(null);

            if (LinkTarget.For(row.Url) is not { } target)
            {
                this.ShowMessage(ComponentPaneWording.LinkNotOpened(row.Name.Trim().Length > 0 ? row.Name : row.Url));
                return;
            }

            if (this.OpenLinkOverrideForTests is { } open)
            {
                open(target);
                return;
            }

            if (!ExternalTargetLauncher.TryOpen(target))
            {
                this.ShowMessage(ComponentPaneWording.LinkNotOpened(row.Name.Trim().Length > 0 ? row.Name : row.Url));
            }
        }

        private void OnBoardLinkHandlePressed(object? sender, PointerPressedEventArgs e) => this.thisBoardLinksDrag?.Press(sender, e);

        internal bool MoveBoardLink(ComponentPaneBoardLinkRow row, int by)
        {
            if (!this.MoveRow(this.thisBoardLinks, row, by))
            {
                return false;
            }

            this.CommitBoardLinks();
            return true;
        }

        // ###########################################################################################
        // Alt+Up / Alt+Down in a row's boxes moves the row one place - the keyboard's way to what the
        // handle's drag does. The focus is put back in the same box of the row where it now stands: a
        // moved row's controls are made again. Not on a drop-down list: there the keys open and close it.
        // ###########################################################################################
        private void OnRowKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.KeyModifiers != KeyModifiers.Alt || e.Key is not (Key.Up or Key.Down) || e.Source is ComboBox ||
                e.Source is not Control { DataContext: ComponentPaneDraggableRow row } source)
            {
                return;
            }

            int by = e.Key == Key.Up ? -1 : +1;
            ItemsControl? list = row switch
            {
                ComponentPaneBoardFileRow => this.BoardFilesItemsControl,
                ComponentPaneBoardLinkRow => this.BoardLinksItemsControl,
                ComponentPaneCreditRow => this.CreditsItemsControl,
                ComponentPaneCategoryRow => this.CategoriesItemsControl,
                ComponentPaneSignalRow => this.SignalsItemsControl,
                ComponentPaneFileRow => this.FilesItemsControl,
                ComponentPaneLinkRow => this.LinksItemsControl,
                _ => null,
            };

            if (list is null)
            {
                return;
            }

            int focusAt = ComponentPane.FocusableIndexIn(list.ContainerFromItem(row), source);

            bool moved = row switch
            {
                ComponentPaneBoardFileRow file => this.MoveBoardFile(file, by),
                ComponentPaneBoardLinkRow link => this.MoveBoardLink(link, by),
                ComponentPaneCreditRow credit => this.MoveCredit(credit, by),
                ComponentPaneSignalRow signal => this.MoveSignal(signal, by),
                ComponentPaneCategoryRow category => this.MoveCategory(category, by),
                ComponentPaneFileRow file => this.MoveFile(file, by),
                ComponentPaneLinkRow link => this.MoveLink(link, by),
                _ => false,
            };

            if (!moved)
            {
                return;
            }

            e.Handled = true;

            Avalonia.Threading.Dispatcher.UIThread.Post(
                () =>
                {
                    if (focusAt >= 0 &&
                        list.ContainerFromItem(row) is Control container &&
                        ComponentPane.FocusablesIn(container).ElementAtOrDefault(focusAt) is { } box)
                    {
                        box.Focus();
                    }
                },
                Avalonia.Threading.DispatcherPriority.Loaded);
        }

        // The controls of a row the keys can reach, in order - and which of them holds `source`.
        private static List<InputElement> FocusablesIn(Control container) =>
            container.GetVisualDescendants().OfType<InputElement>().Where(element => element.Focusable && element.IsEffectivelyVisible).ToList();

        private static int FocusableIndexIn(Control? container, Control source)
        {
            if (container is null)
            {
                return -1;
            }

            List<InputElement> focusables = ComponentPane.FocusablesIn(container);

            for (int index = 0; index < focusables.Count; index++)
            {
                if (ReferenceEquals(focusables[index], source) || (focusables[index] is Visual visual && source.GetVisualAncestors().Contains(visual)))
                {
                    return index;
                }
            }

            return -1;
        }

        // ###########################################################################################
        // A row of files, links, credits or important signals moved one place up (-1) or down (+1),
        // past the row next to it ON SCREEN whatever its category (owner request, 2026-10-30 -
        // BoardPartEdits.CanMove). What the boxes hold goes in first; the rows are put in the order
        // BoardPartEdits.Moved gives - in place, so the list keeps every row's boxes. False when it
        // cannot move.
        // ###########################################################################################
        private bool MoveRow<TRow>(ObservableCollection<TRow> rows, TRow row, int by)
            where TRow : ComponentPaneDraggableRow
        {
            this.CommitHeld();

            int index = rows.IndexOf(row);

            if (!BoardPartEdits.CanMove(rows, index, by))
            {
                return false;
            }

            List<TRow> moved = BoardPartEdits.Moved(rows, index, by);

            for (int wanted = 0; wanted < moved.Count; wanted++)
            {
                int at = rows.IndexOf(moved[wanted]);

                if (at != wanted)
                {
                    rows.Move(at, wanted);
                }
            }

            return true;
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

        // "Add me" - a credit by the name CRT remembered - went on 2026-10-29 (owner decision: "Remove
        // the button entirely - there is already the 'Add credit' button, which can be used").

        private void OnCreditHandlePressed(object? sender, PointerPressedEventArgs e) => this.thisCreditsDrag?.Press(sender, e);

        internal bool MoveCredit(ComponentPaneCreditRow row, int by)
        {
            if (!this.MoveRow(this.thisCredits, row, by))
            {
                return false;
            }

            this.CommitCredits();
            return true;
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

        // ###########################################################################################
        // A NEW SIGNAL PICKS ITS NET (owner request, 2026-10-29: the KiCad net is the KiCad data's -
        // "they should not be editable, but rather just show as texts"): the picker offers the board's
        // nets no row has, "Add signal" is on only for one of them, and the row comes named as it would
        // be suggested (ImportantSignalSuggestions.StartingNameOf) - the name is changed in its row.
        // ###########################################################################################
        private IReadOnlyList<string> thisNetsToAdd = [];

        // The picker's list is replaced only when the nets in it change (code review, 2026-10-30):
        // NetsToAdd answers a new list every time, so every commit of the signals replaced it, and the
        // drop-down cleared and refilled under a net being picked.
        private void ShowNewSignalPicker()
        {
            IReadOnlyList<string> nets = this.thisNetChoices is { HasNets: true } choices
                ? ImportantSignalSuggestions.NetsToAdd(choices, this.thisSignals.Select(row => row.ToEntry()))
                : [];

            if (this.NewSignalNetBox.ItemsSource is IReadOnlyList<string> shown && shown.SequenceEqual(nets, StringComparer.Ordinal))
            {
                nets = shown;
            }

            this.thisNetsToAdd = nets;

            if (!ReferenceEquals(this.NewSignalNetBox.ItemsSource, nets))
            {
                this.NewSignalNetBox.ItemsSource = nets;
            }

            this.NewSignalNetBox.IsEnabled = this.thisNetChoices is { HasNets: true };
            this.ShowAddSignalButton();
        }

        // The net the picker holds, as the board's KiCad data spells it - null for none of the nets a
        // new signal may pick.
        private string? NetPicked()
        {
            string? net = this.thisNetChoices?.NormalizedOf(this.NewSignalNetBox.Text);

            return net is not null && this.thisNetsToAdd.Contains(net, StringComparer.OrdinalIgnoreCase) ? net : null;
        }

        private void ShowAddSignalButton()
        {
            bool hasNets = this.thisNetChoices is { HasNets: true };

            this.AddSignalButton.IsEnabled = hasNets && this.NetPicked() is not null;
            ToolTip.SetTip(
                this.AddSignalButton,
                !hasNets ? ComponentPaneWording.AddSignalNeedsKiCadTip
                    : this.AddSignalButton.IsEnabled ? ComponentPaneWording.AddSignalTip
                    : ComponentPaneWording.PickANetTip);
        }

        private void OnNewSignalNetChanged(object? sender, TextChangedEventArgs e) => this.ShowAddSignalButton();

        // Enter in the picker adds the net picked, as the button does.
        private void OnNewSignalNetKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Enter && this.AddSignalForNet(this.NewSignalNetBox.Text))
            {
                e.Handled = true;
            }
        }

        private void OnAddSignalClick(object? sender, RoutedEventArgs e) => this.AddSignalForNet(this.NewSignalNetBox.Text);

        // A signal on `net` - one of the board's nets no row has - as ONE step, under `name` or else the
        // name it starts with; false when it is none.
        internal bool AddSignalForNet(string? net, string? name = null)
        {
            if (this.thisSession is null || this.thisBoardPart != PaneBoardPart.ImportantSignals)
            {
                return false;
            }

            this.CommitHeld();
            this.NewSignalNetBox.Text = net;

            if (this.NetPicked() is not { } picked)
            {
                return false;
            }

            var row = new ComponentPaneSignalRow
            {
                DisplayName = name ?? ImportantSignalSuggestions.StartingNameOf(picked),
                KiCadNetName = picked,
            };

            row.UseChoices(this.thisNetChoices);
            this.thisSignals.Add(row);

            this.CommitSignals();
            this.NewSignalNetBox.Text = string.Empty;
            this.ShowBoardPartLines();
            return true;
        }

        private void OnRemoveSignalClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneSignalRow row } && this.thisSignals.Remove(row))
            {
                this.CommitSignals();
                this.ShowBoardPartLines();
            }
        }

        // An important signal moved up or down (owner request, 2026-10-30: "I need a possibility to
        // reorder the 'Important signals' entries") - the Schematics tab's panel lists them so.
        private void OnSignalHandlePressed(object? sender, PointerPressedEventArgs e) => this.thisSignalsDrag?.Press(sender, e);

        internal bool MoveSignal(ComponentPaneSignalRow row, int by)
        {
            if (!this.MoveRow(this.thisSignals, row, by))
            {
                return false;
            }

            this.CommitSignals();
            return true;
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

    // ###########################################################################################
    // A row DRAGGED up or down by its handle, a dashed placeholder where it will land (ListRowDrag - owner
    // request, 2026-10-30: "a North/South icon and then drag the whole line up/down, just alike in the
    // workbook for photos"): the files, links, credits and important signals - the board's, and a
    // component's files and links (owner request, 2026-11-02) - and the categories.
    // ###########################################################################################
    public abstract class ComponentPaneDraggableRow : ComponentPaneRowBase, IDraggableRow
    {
        private bool thisIsDropPlaceholder;
        private double thisPlaceholderHeight;

        public bool IsDropPlaceholder
        {
            get => this.thisIsDropPlaceholder;
            set
            {
                if (this.thisIsDropPlaceholder != value)
                {
                    this.thisIsDropPlaceholder = value;
                    this.Raise(nameof(this.IsDropPlaceholder));
                    this.Raise(nameof(this.IsNotDropPlaceholder));
                }
            }
        }

        public bool IsNotDropPlaceholder => !this.thisIsDropPlaceholder;

        public double PlaceholderHeight
        {
            get => this.thisPlaceholderHeight;
            set
            {
                if (this.thisPlaceholderHeight != value)
                {
                    this.thisPlaceholderHeight = value;
                    this.Raise(nameof(this.PlaceholderHeight));
                }
            }
        }
    }

    // ###########################################################################################
    // A row of the files, links, credits or important signals, each field it needs MARKED while empty
    // (BoardRowRequired - owner request, 2026-11-01: "highlight the problematic cell with a helper what
    // is missing"): its "...Missing" says so to the markup, raised whenever a field changes.
    // ###########################################################################################
    public abstract class ComponentPaneListRow : ComponentPaneRegionalRow
    {
        // Nothing typed in it - no row of the draft (CommitList). BoardPartEdits.IsBlank's rule, the one
        // the list written goes by: two copies could count a row the edit then dropped.
        public abstract bool IsBlank { get; }

        // The columns it needs and has empty (BoardRowRequired), and whether `column` is one of them.
        protected abstract IReadOnlyList<string> EmptyColumns();

        protected bool IsMissing(string column) => this.EmptyColumns().Contains(column);

        // A field changed: every "...Missing" read again.
        protected abstract void RaiseMissing();

        protected void SetField(ref string field, string? value, [CallerMemberName] string? name = null)
        {
            if (this.Set(ref field, value, name))
            {
                this.RaiseMissing();
            }
        }
    }

    // A list row with the row as the draft has it - null for a row not in the draft (one with nothing
    // typed in it yet).
    public abstract class ComponentPaneListRow<TEntry> : ComponentPaneListRow
        where TEntry : class
    {
        public TEntry? Committed { get; set; }

        public abstract TEntry ToEntry();
    }

    public sealed class ComponentPaneBoardFileRow : ComponentPaneListRow<BoardLocalFileEntry>
    {
        private string thisCategory = string.Empty;
        private string thisName = string.Empty;
        private string thisFile = string.Empty;

        public string Category { get => this.thisCategory; set => this.SetField(ref this.thisCategory, value); }

        public string Name { get => this.thisName; set => this.SetField(ref this.thisName, value); }

        public string File
        {
            get => this.thisFile;
            set
            {
                if (this.Set(ref this.thisFile, value))
                {
                    this.Raise(nameof(this.HasFile));
                    this.RaiseMissing();
                }
            }
        }

        // "Open" is on while the row names a file.
        public bool HasFile => this.thisFile.Trim().Length > 0;

        public bool CategoryMissing => this.IsMissing(BoardWorkbookSchema.ColCategory);

        public bool NameMissing => this.IsMissing(BoardWorkbookSchema.ColName);

        public bool FileMissing => this.IsMissing(BoardWorkbookSchema.ColFile);

        public override bool IsBlank => BoardPartEdits.IsBlank(this.ToEntry());

        protected override IReadOnlyList<string> EmptyColumns() => BoardRowRequired.EmptyColumns(this.ToEntry());

        protected override void RaiseMissing()
        {
            this.Raise(nameof(this.CategoryMissing));
            this.Raise(nameof(this.NameMissing));
            this.Raise(nameof(this.FileMissing));
        }

        public IReadOnlyList<string> CategoryChoices { get; init; } = [];

        public static ComponentPaneBoardFileRow From(BoardLocalFileEntry row, IReadOnlyList<string> categories) => new()
        {
            Category = row.Category ?? string.Empty,
            Name = row.Name ?? string.Empty,
            File = row.File ?? string.Empty,
            Region = row.Region ?? string.Empty,
            CategoryChoices = categories,
            Committed = row,
        };

        public override BoardLocalFileEntry ToEntry() => new() { Category = this.thisCategory, Name = this.thisName, File = this.thisFile, Region = this.Region };
    }

    public sealed class ComponentPaneBoardLinkRow : ComponentPaneListRow<BoardLinkEntry>
    {
        private string thisCategory = string.Empty;
        private string thisName = string.Empty;
        private string thisUrl = string.Empty;

        public string Category { get => this.thisCategory; set => this.SetField(ref this.thisCategory, value); }

        public string Name { get => this.thisName; set => this.SetField(ref this.thisName, value); }

        public string Url
        {
            get => this.thisUrl;
            set
            {
                if (this.Set(ref this.thisUrl, value))
                {
                    this.Raise(nameof(this.CanOpen));
                    this.RaiseMissing();
                }
            }
        }

        // "Open" is on while the address is one a browser can be sent to (LinkTarget).
        public bool CanOpen => LinkTarget.For(this.thisUrl) is not null;

        public bool CategoryMissing => this.IsMissing(BoardWorkbookSchema.ColCategory);

        public bool NameMissing => this.IsMissing(BoardWorkbookSchema.ColName);

        public bool UrlMissing => this.IsMissing(BoardWorkbookSchema.ColUrl);

        // An address the checks refuse - marked as an empty one is, the helper saying what to write:
        // "Open link" opens "www.lemon64.com", and Submit refused it (code review, 2026-11-01).
        public bool UrlNotWeb => BoardDataChecks.IsRefusedLink(this.thisUrl);

        public bool UrlMarked => this.UrlMissing || this.UrlNotWeb;

        public string NotWebText => ComponentPaneWording.LinkNotWeb;

        // The category "Add link" filled in from the row above. Nobody typed it, so a row holding only
        // that is a row with nothing typed in it (code review, 2026-11-01): pressed and thought better of,
        // it went into the draft as two errors stopping Submit.
        public string OfferedCategory { get; init; } = string.Empty;

        public override bool IsBlank =>
            BoardPartEdits.IsBlank(this.ToEntry()) ||
            (string.IsNullOrWhiteSpace(this.thisName) && string.IsNullOrWhiteSpace(this.thisUrl) &&
             string.Equals(this.thisCategory.Trim(), this.OfferedCategory.Trim(), StringComparison.Ordinal));

        protected override IReadOnlyList<string> EmptyColumns() => BoardRowRequired.EmptyColumns(this.ToEntry());

        protected override void RaiseMissing()
        {
            this.Raise(nameof(this.CategoryMissing));
            this.Raise(nameof(this.NameMissing));
            this.Raise(nameof(this.UrlMissing));
            this.Raise(nameof(this.UrlNotWeb));
            this.Raise(nameof(this.UrlMarked));
        }

        public IReadOnlyList<string> CategoryChoices { get; init; } = [];

        public static ComponentPaneBoardLinkRow From(BoardLinkEntry row, IReadOnlyList<string> categories) => new()
        {
            Category = row.Category ?? string.Empty,
            Name = row.Name ?? string.Empty,
            Url = row.Url ?? string.Empty,
            Region = row.Region ?? string.Empty,
            CategoryChoices = categories,
            Committed = row,
        };

        public override BoardLinkEntry ToEntry() => new() { Category = this.thisCategory, Name = this.thisName, Url = this.thisUrl, Region = this.Region };
    }

    public sealed class ComponentPaneCreditRow : ComponentPaneListRow<CreditEntry>
    {
        private string thisCategory = string.Empty;
        private string thisSubCategory = string.Empty;
        private string thisNameOrHandle = string.Empty;
        private string thisContact = string.Empty;

        public string Category { get => this.thisCategory; set => this.SetField(ref this.thisCategory, value); }

        public string SubCategory { get => this.thisSubCategory; set => this.SetField(ref this.thisSubCategory, value); }

        public string NameOrHandle { get => this.thisNameOrHandle; set => this.SetField(ref this.thisNameOrHandle, value); }

        public string Contact { get => this.thisContact; set => this.SetField(ref this.thisContact, value); }

        // "What" and "Name or handle" - which part of it and the contact may stay empty.
        public bool CategoryMissing => this.IsMissing(BoardWorkbookSchema.ColCategory);

        public bool NameOrHandleMissing => this.IsMissing(BoardWorkbookSchema.ColNameOrHandle);

        public override bool IsBlank => BoardPartEdits.IsBlank(this.ToEntry());

        protected override IReadOnlyList<string> EmptyColumns() => BoardRowRequired.EmptyColumns(this.ToEntry());

        protected override void RaiseMissing()
        {
            this.Raise(nameof(this.CategoryMissing));
            this.Raise(nameof(this.NameOrHandleMissing));
        }

        public IReadOnlyList<string> CategoryChoices { get; init; } = [];

        public static ComponentPaneCreditRow From(CreditEntry row, IReadOnlyList<string> categories) => new()
        {
            Category = row.Category ?? string.Empty,
            SubCategory = row.SubCategory ?? string.Empty,
            NameOrHandle = row.NameOrHandle ?? string.Empty,
            Contact = row.Contact ?? string.Empty,
            CategoryChoices = categories,
            Committed = row,
        };

        public override CreditEntry ToEntry() => new()
        {
            Category = this.thisCategory,
            SubCategory = this.thisSubCategory,
            NameOrHandle = this.thisNameOrHandle,
            Contact = this.thisContact,
        };
    }

    // One important signal - its net picked from the board's KiCad nets, and what is wrong with it.
    public sealed class ComponentPaneSignalRow : ComponentPaneListRow<KiCadImportantSignalEntry>
    {
        private string thisDisplayName = string.Empty;
        private string thisNetName = string.Empty;
        private string thisProblem = string.Empty;
        private IReadOnlyList<string> thisNetChoices = [];

        public string DisplayName { get => this.thisDisplayName; set => this.SetField(ref this.thisDisplayName, value); }

        public string KiCadNetName
        {
            get => this.thisNetName;
            set
            {
                if (this.Set(ref this.thisNetName, value))
                {
                    this.Raise(nameof(this.CanShowOnBoard));
                    this.RaiseMissing();
                }
            }
        }

        public bool NetMissing => this.IsMissing(BoardWorkbookSchema.ColKiCadNetName);

        public bool DisplayNameMissing => this.IsMissing(BoardWorkbookSchema.ColDisplayName);

        public override bool IsBlank => BoardPartEdits.IsBlank(this.ToEntry());

        protected override IReadOnlyList<string> EmptyColumns() => BoardRowRequired.EmptyColumns(this.ToEntry());

        protected override void RaiseMissing()
        {
            this.Raise(nameof(this.NetMissing));
            this.Raise(nameof(this.DisplayNameMissing));
        }

        // "Show on board" lights the net up - while there is one, and the board's KiCad data has it: a
        // board without KiCad data has nothing to light (code review, 2026-10-26).
        public bool CanShowOnBoard => this.thisNetChoices.Count > 0 && this.thisNetName.Trim().Length > 0 && this.thisProblem.Length == 0;

        public IReadOnlyList<string> NetChoices
        {
            get => this.thisNetChoices;
            private set
            {
                this.thisNetChoices = value;
                this.Raise(nameof(this.NetChoices));
                this.Raise(nameof(this.CanShowOnBoard));
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
                    this.Raise(nameof(this.CanShowOnBoard));
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
                Committed = row,
            };

            shown.UseChoices(choices);
            return shown;
        }

        public override KiCadImportantSignalEntry ToEntry() => new() { DisplayName = this.thisDisplayName, KiCadNetName = this.thisNetName };
    }
}
