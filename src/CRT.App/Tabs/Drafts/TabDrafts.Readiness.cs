using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Layout;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // "BOARD STATUS" - A PANEL PER PART OF THE BOARD (owner request, 2026-11-01: "each point in its own
    // panel with a headline, then an explanation what this is about, a button for it and lastly also a
    // status where it will count the data for it, if any. Maybe some coloring and the checkmark icon or
    // a 'missing to do' icon"). It replaced the new board's checklist and the published board's "What's
    // missing"; what a panel says is BoardStatusPage's.
    //
    // On the right of the Contribute tab while nothing else is open there (the start panel):
    //
    //   - THE PANELS follow the left column's entries, each with the button to its part.
    //   - SUGGESTED MATCHES (slice 3d - the plan: "Match: suggested pairs ("...Top.png" with "PCB Top")
    //     to confirm") are under Schematic images, each picture with the KiCad view its name and its
    //     file's name say it shows (SchematicCadNameSuggestions); "Use" takes one, "Use all N" every
    //     one - one step either way.
    //   - WHAT EACH COMPONENT LACKS is under Components, a button per component that opens it
    //     (BoardReadiness.Missing).
    //   - "ADD N FROM KICAD" (slice 3b) puts every component of the board's KiCad data the draft lacks
    //     into it in one step - with their values, categories, packages, pins and datasheet links. It
    //     is the KiCad import's since 2026-11-01 (KiCadDataPanel).
    //
    // Worked out ONCE for each board shown and each time the right side goes back to the start panel -
    // never after every refresh of the drafts list, which runs whenever CRT comes to the front, and not
    // at all while the tab is off screen: then when it is shown (code review, 2026-10-26 - twice per
    // board load, in any tab). EVERY PANEL AND BUTTON IS MADE ONCE AND CHANGED IN PLACE: the KiCad data
    // arrives a moment after the rest, Main reads the board again after a change, and a control made
    // again under a press threw that click away (the trap the drafts list met, 2026-10-10). A board's
    // KiCad facts are kept for it, so the page coming back shows them while they are read again rather
    // than flicking to "not known yet".
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // Which time the page was asked for - a read answers only the latest ask.
        private int thisReadinessAsk;

        // Asked for while the tab was off screen - worked out when it is shown.
        private bool thisReadinessStale;

        // One panel's controls, made once: its icon, words and button - and Extra, under its lines, for
        // what only that part has (the suggested matches, what each component lacks).
        private sealed record StatusRow(Border Panel, TextBlock Icon, TextBlock Title, TextBlock Explanation, StackPanel Lines, Button Action, Button Table, StackPanel Extra, StatusProblems Problems);

        // A panel's errors and warnings, in a panel of their own inside it: how many, then each. Inner is what
        // stands in it, which the components marked nowhere and "Show in the table" join.
        private sealed record StatusProblems(Border Block, TextBlock Icon, TextBlock Title, StackPanel Inner, StackPanel Rows, TextBlock More);

        private readonly Dictionary<BoardStatusPart, StatusRow> thisStatusRows = new();

        // ###########################################################################################
        // The matches suggested under "Match pictures to KiCad views" (slice 3d), made once in its
        // row: a line saying where they come from, a row per pair with its "Use", and "Use all N" -
        // and the pairs shown, which "Use all" takes.
        // ###########################################################################################
        private sealed record MatchControls(StackPanel Panel, StackPanel Rows, Button UseAll);

        private MatchControls? thisMatchControls;
        private IReadOnlyList<SchematicCadNameSuggestion> thisShownMatches = [];

        // The panels shown, and the KiCad files last counted for a board.
        private IReadOnlyList<BoardStatusPanel> thisStatusPanels = [];

        private (string Board, IReadOnlyList<string> Files)? thisStatusFiles;

        // The last KiCad facts read, and the board they are of.
        private (string Board, BoardKiCadFacts Facts)? thisReadinessFacts;

        // The draft's errors and warnings as last checked, with the rows they are numbered by, and the board
        // they are of.
        private (string Board, BoardProblemsRead Read)? thisStatusProblems;

        // How many credits' rows a new board would be given, as last worked out, and the board it is of.
        private (string Board, int Rows)? thisStatusCreditsToOffer;

        // ###########################################################################################
        // What a button did - said only while the draft is as it left it (code review,
        // 2026-10-26: after an Undo it still said the components were added): its board, the draft's
        // content it was said of, and - for a save that failed while something else waited to be
        // saved - whether it goes once the draft is saved.
        // ###########################################################################################
        private sealed record ReadinessNote(string Board, string Text, BoardContent Content, bool GoesWhenSaved);

        private ReadinessNote? thisReadinessNote;

        // What the components lack, under the Components panel, made once: a heading, the first labels as
        // buttons, "and N more".
        private sealed record MissingGroup(StackPanel Panel, TextBlock Heading, WrapPanel Buttons, TextBlock More);

        private MissingGroup[]? thisMissingGroups;

        // The page's last work - for a test to wait for.
        private Task thisReadinessShown = Task.CompletedTask;

        // Stands in for reading a board's KiCad data - a test has no KiCad files.
        internal Func<string, Task<KiCadProjectRoot?>>? BoardKiCadOverrideForTests { get; set; }

        internal IReadOnlyList<BoardStatusPanel> StatusPanelsForTests => this.thisStatusPanels;

        internal BoardStatusPanel StatusPanelForTests(BoardStatusPart part) => this.thisStatusPanels.Single(panel => panel.Part == part);

        // A panel's controls, and its button - for a test to press.
        internal Border StatusBorderForTests(BoardStatusPart part) => this.thisStatusRows[part].Panel;

        internal Button StatusButtonForTests(BoardStatusPart part) => this.thisStatusRows[part].Action;

        internal Button StatusTableButtonForTests(BoardStatusPart part) => this.thisStatusRows[part].Table;

        // A panel's own panel of errors and warnings as it stands: not shown is null.
        internal (string Title, List<Grid> Rows, string More)? StatusProblemsForTests(BoardStatusPart part)
        {
            StatusProblems shown = this.thisStatusRows[part].Problems;

            return shown.Block.IsVisible
                ? (shown.Title.IsVisible ? shown.Title.Text ?? string.Empty : string.Empty,
                   shown.Rows.Children.OfType<Grid>().ToList(),
                   shown.More.IsVisible ? shown.More.Text ?? string.Empty : string.Empty)
                : null;
        }

        internal Border StatusProblemsBlockForTests(BoardStatusPart part) => this.thisStatusRows[part].Problems.Block;

        internal string? ReadinessNoteForTests => this.ReadinessNoteText.IsVisible ? Handlers.Theming.BoldRuns.TextOf(this.ReadinessNoteText) : null;

        internal BoardMissing? MissingForTests { get; private set; }

        internal Task ReadinessShownForTests => this.thisReadinessShown;

        // How many times the lists were worked out - each is a session opened and KiCad data read.
        internal int ReadinessAsksForTests => this.thisReadinessAsk;

        // Stands in for whether the tab is on screen - a test's tab is in no window.
        internal Func<bool>? IsOnScreenOverrideForTests { get; set; }

        // ###########################################################################################
        // The start panel shown or hidden - and its panels worked out again each time it comes back
        // (the component pane or the table closed).
        // ###########################################################################################
        private void ShowStartPanel(bool visible)
        {
            bool cameBack = visible && !this.StartPanel.IsVisible;
            this.StartPanel.IsVisible = visible;

            if (cameBack)
            {
                this.ShowReadiness();
            }
        }

        // ###########################################################################################
        // Whether the tab is on screen: in CRT, in the window - a TabControl takes a tab not selected
        // out of it. A tab no Main hosts (a test's) is taken as on screen.
        // ###########################################################################################
        private bool IsOnScreen() =>
            this.IsOnScreenOverrideForTests?.Invoke() ?? (this.thisMainWindow is null || TopLevel.GetTopLevel(this) is not null);

        // The tab shown: what was asked for while it was not.
        private void ShowReadinessIfStale()
        {
            if (this.thisReadinessStale)
            {
                this.ShowReadiness();
            }
        }

        // ###########################################################################################
        // The panels, for the board shown: at once from what is known - the draft held in memory, the
        // KiCad facts last read for the board - then again once all of it is read (ShowStatusOfAsync).
        // ###########################################################################################
        private void ShowReadiness()
        {
            if (!this.IsOnScreen())
            {
                this.thisReadinessStale = true;
                return;
            }

            this.thisReadinessStale = false;
            int ask = ++this.thisReadinessAsk;

            if (this.thisReadinessNote is { } note && !string.Equals(note.Board, this.thisCurrentEntry?.ExcelDataFile, StringComparison.OrdinalIgnoreCase))
            {
                this.thisReadinessNote = null;
            }

            if (!this.StartPanel.IsVisible || this.thisCurrentEntry is not { } entry || this.thisCurrentBoard is not { } board)
            {
                this.HideStatus();
                return;
            }

            string identity = entry.ExcelDataFile;
            DraftSession? held = DraftManager.Sessions.Find(DraftManager.DraftsRoot, identity);

            this.ShowStatus(entry, held, held?.Content.Board ?? board, held?.Content.Calibrations, this.KnownKiCadFiles(identity), this.KnownKiCadFacts(identity), this.KnownProblems(identity), this.KnownCreditsToOffer(identity));

            this.thisReadinessShown = this.ShowStatusOfAsync(entry, board, ask);
        }

        private void HideStatus()
        {
            this.StatusPanelsHost.IsVisible = false;
            this.thisStatusPanels = [];
        }

        // ###########################################################################################
        // THE PANELS, ONCE ALL OF IT IS READ: the board as its draft has it - or, with no draft, as its
        // published board file has it (its calibrations are there) - its KiCad files, and what its KiCad
        // data says, read off the UI thread (KiCadProjectLoader keeps the parse, so the board on screen
        // has usually read it already). A newer ask wins; a new board whose draft cannot be opened
        // shows nothing. Reading it changes nothing in the draft (TabDrafts.NewBoardCredits.cs).
        // ###########################################################################################
        private async Task ShowStatusOfAsync(HardwareBoardEntry entry, BoardData shown, int ask)
        {
            try
            {
                string draftsRoot = DraftManager.DraftsRoot;
                string identity = entry.ExcelDataFile;
                string published = entry.IsDraftOnly ? string.Empty : this.PublishedBoardFileOf(entry);

                DraftSession? session = DraftManager.Sessions.Find(draftsRoot, identity)
                                        ?? await Task.Run(() => entry.IsDraftOnly || DraftBoardSource.HasDraft(draftsRoot, identity)
                                            ? DraftManager.Sessions.Get(draftsRoot, identity)
                                            : null);

                // The page gone meanwhile, nothing below is wanted. The page coming back asks again
                // (ShowStartPanel).
                if (ask != this.thisReadinessAsk || !this.StartPanel.IsVisible)
                {
                    return;
                }

                if (session is null && entry.IsDraftOnly)
                {
                    this.HideStatus();
                    return;
                }

                IReadOnlyList<string>? listed = session?.ListedFiles.Keys.ToList();

                string dataRoot = DataManager.DataRoot;
                string draftFolder = DraftFolderLayout.GetBoardFolder(draftsRoot, identity);

                // The draft's errors and warnings, as the table's checks find them on what it holds now - and a board
                // with no draft, its published board's (owner decision, 2026-11-07: the launch check's findings "should
                // show just like any other error or warning in the "Contribute" tab"), the very check its badges wait for
                // (PublishedProblemsOf; code review, 2026-11-08: a second one ran beside it). They go with the very board
                // they ran on.
                Task<BoardProblemsRead>? publishedCheck = session is null && !entry.IsDraftOnly ? this.PublishedProblemsOf(entry) : null;

                (BoardContent? content, List<string> files, BoardProblemsRead? ofDraft, int creditsToOffer) = await Task.Run(() =>
                {
                    BoardContent? read = session?.Content ?? TabDrafts.PublishedContentOf(published);

                    BoardProblemsRead? found = null;

                    if (session is not null && read is not null)
                    {
                        IReadOnlyList<BoardDataProblem> problemsOfDraft = DraftStatusReader.ProblemsOf(read.Board, dataRoot, draftFolder);
                        found = new BoardProblemsRead(BoardProblemCounts.Of(problemsOfDraft), problemsOfDraft, read.Board);
                    }

                    return (read,
                            TabDrafts.KiCadFileNamesOf(draftsRoot, identity, listed, published),
                            found,
                            session is null ? 0 : TabDrafts.NewBoardCreditsToOffer(entry, session));
                });

                BoardProblemsRead problems = ofDraft ?? (publishedCheck is null ? BoardProblemsRead.None : await publishedCheck);

                if (ask != this.thisReadinessAsk)
                {
                    return;
                }

                this.thisStatusFiles = (identity, files);
                this.thisStatusProblems = (identity, problems);
                this.thisStatusCreditsToOffer = (identity, creditsToOffer);

                BoardKiCadFacts? known = files.Count > 0 ? this.KnownKiCadFacts(identity) : null;
                this.ShowStatus(entry, session, session?.Content.Board ?? content?.Board ?? shown, session?.Content.Calibrations ?? content?.Calibrations, files, known, problems, creditsToOffer);

                if (files.Count == 0)
                {
                    return;
                }

                KiCadProjectRoot? root = await this.ReadKiCadForReadinessAsync(identity, "its board status");

                if (ask != this.thisReadinessAsk)
                {
                    return;
                }

                BoardKiCadFacts facts = await Task.Run(() => BoardKiCadFacts.Of(root));

                if (ask == this.thisReadinessAsk)
                {
                    this.thisReadinessFacts = (identity, facts);
                    this.ShowStatus(entry, session, session?.Content.Board ?? content?.Board ?? shown, session?.Content.Calibrations ?? content?.Calibrations, files, facts, problems, creditsToOffer);
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not work out the board status of [{entry.ExcelDataFile}] - [{ex.Message}]");
            }
        }

        // A published board file's content - null when it cannot be read.
        private static BoardContent? PublishedContentOf(string boardFile)
        {
            if (!System.IO.File.Exists(boardFile) || !DataTreeLayout.IsBoardFile(boardFile))
            {
                return null;
            }

            try
            {
                return BoardFiles.ReadDocument(boardFile) is { } document ? BoardDocumentMapping.ToContent(document) : null;
            }
            catch (Exception ex) when (ex is System.IO.IOException or UnauthorizedAccessException or System.Text.Json.JsonException or InvalidOperationException)
            {
                Logger.Warning($"Could not read the published board file [{boardFile}] for its board status - [{ex.Message}]");
                return null;
            }
        }

        // The KiCad facts last read for a board - null when none were.
        private BoardKiCadFacts? KnownKiCadFacts(string identity) =>
            this.thisReadinessFacts is { } last && string.Equals(last.Board, identity, StringComparison.OrdinalIgnoreCase) ? last.Facts : null;

        // The KiCad files last counted for a board - null while they are not known.
        private IReadOnlyList<string>? KnownKiCadFiles(string identity) =>
            this.thisStatusFiles is { } last && string.Equals(last.Board, identity, StringComparison.OrdinalIgnoreCase) ? last.Files : null;

        // The draft's errors and warnings as last checked - none known yet is none said.
        private BoardProblemsRead KnownProblems(string identity) =>
            this.thisStatusProblems is { } last && string.Equals(last.Board, identity, StringComparison.OrdinalIgnoreCase) ? last.Read : BoardProblemsRead.None;

        // The credits' rows a new board would be given, as last worked out - none known yet is none said.
        private int KnownCreditsToOffer(string identity) =>
            this.thisStatusCreditsToOffer is { } last && string.Equals(last.Board, identity, StringComparison.OrdinalIgnoreCase) ? last.Rows : 0;

        // What the last button here did - said while the draft is the content it was said of.
        private void ShowReadinessNote(string board, DraftSession? session)
        {
            bool noteShown = session is not null &&
                             this.thisReadinessNote is { } note &&
                             string.Equals(note.Board, board, StringComparison.OrdinalIgnoreCase) &&
                             ReferenceEquals(note.Content, session.Content) &&
                             !(note.GoesWhenSaved && session.SaveState == DraftSaveState.Saved);

            Handlers.Theming.BoldRuns.ShowMarked(this.ReadinessNoteText, noteShown ? this.thisReadinessNote!.Text : null);
            this.ReadinessNoteText.IsVisible = noteShown;
        }

        // The board's KiCad project, as the Schematics tab draws it (ReadBoardKiCadAsync) - null without.
        private Task<KiCadProjectRoot?> ReadKiCadForReadinessAsync(string identity, string what) =>
            this.BoardKiCadOverrideForTests is { } read
                ? read(identity)
                : TabDrafts.ReadBoardKiCadAsync(DataManager.DataRoot, DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, identity), identity, what);

        // ###########################################################################################
        // The panels shown - each made once, its words and state changed in place; one with nothing to
        // say hidden.
        // ###########################################################################################
        private void ShowStatus(
            HardwareBoardEntry entry,
            DraftSession? session,
            BoardData board,
            IReadOnlyList<KiCadCalibrationEntry>? calibrations,
            IReadOnlyList<string>? files,
            BoardKiCadFacts? kiCad,
            BoardProblemsRead problems,
            int creditsToOffer)
        {
            this.CreateStatusRows();

            BoardMissing missing = BoardReadiness.Missing(board, hasKiCadFiles: files is { Count: > 0 }, kiCad);
            List<BoardStatusPanel> panels = BoardStatusPage.Panels(board, missing, calibrations, files, kiCad, problems.Problems, creditsToOffer, problems.Board);

            this.MissingForTests = missing;
            this.thisStatusPanels = panels;

            foreach ((BoardStatusPart part, StatusRow row) in this.thisStatusRows)
            {
                if (panels.FirstOrDefault(panel => panel.Part == part) is { } panel)
                {
                    TabDrafts.ShowStatusRow(row, panel);
                    this.ShowStatusProblems(row.Problems, panel);
                }
                else
                {
                    row.Panel.IsVisible = false;
                }
            }

            this.ShowMissingComponents(missing);
            this.ShowMatches(
                this.thisStatusRows[BoardStatusPart.Pictures].Extra,
                kiCad is { Read: true } facts && files is { Count: > 0 }
                    ? SchematicCadNameSuggestions.For(board.Schematics, facts.ToMatch ?? [])
                    : []);

            this.ShowReadinessNote(entry.ExcelDataFile, session);
            this.StatusPanelsHost.IsVisible = true;
        }

        // A panel's icon, which differs in shape too, so nothing rests on the color alone.
        private static string StatusIconOf(BoardStatusState state) => state switch
        {
            BoardStatusState.Done => "\uf058",
            BoardStatusState.ToDo => "\uf06a",
            BoardStatusState.Error => "\uf057",
            _ => "\uf056",
        };

        private static void ShowStatusRow(StatusRow row, BoardStatusPanel panel)
        {
            row.Panel.IsVisible = true;
            row.Title.Text = panel.Title;
            Handlers.Theming.BoldRuns.ShowMarked(row.Explanation, panel.Explanation);
            row.Action.Content = panel.Action;

            row.Icon.Text = TabDrafts.StatusIconOf(panel.State);

            TabDrafts.UseStateClasses(row.Panel, panel.State);
            TabDrafts.UseStateClasses(row.Icon, panel.State);

            Avalonia.Controls.Controls lines = row.Lines.Children;

            for (int index = 0; index < panel.Lines.Count; index++)
            {
                if (index == lines.Count)
                {
                    lines.Add(new TextBlock { Classes = { "StatusLine" } });
                }

                BoardStatusLine line = panel.Lines[index];
                var block = (TextBlock)lines[index];

                block.Text = line.Depth > 0 ? $"- {line.Text}" : line.Text;
                block.Margin = new Avalonia.Thickness(line.Depth * 16, 0, 0, 0);
                block.Classes.Set("todo", line.Tone == BoardStatusState.ToDo);
                block.Classes.Set("error", line.Tone == BoardStatusState.Error);
            }

            while (lines.Count > panel.Lines.Count)
            {
                lines.RemoveAt(lines.Count - 1);
            }

            row.Lines.IsVisible = panel.Lines.Count > 0;

            // The table a button away from what its checks found - the panel of its own's own button is it.
            row.Table.IsVisible = panel.TableSheet is not null && panel.Part != BoardStatusPart.Problems;
        }

        // ###########################################################################################
        // WHAT THE CHECKS FOUND ON A PANEL'S PART, IN A PANEL OF THEIR OWN INSIDE IT (owner requests, 2026-11-04:
        // "the error and warning panels should be shown specifically inside those panels and then detail the
        // issues in there also"; "encapsulated in its own panel, as now the text is kind of mixed with the normal
        // text"), right under the headline: how many, then each - a component's with the way to it - as the panel
        // of a page has them. Changed in place, as everything here: a row made again under a press threw the
        // click away.
        // ###########################################################################################
        private void ShowStatusProblems(StatusProblems shown, BoardStatusPanel panel)
        {
            BoardProblemCounts counts = panel.ProblemCounts;
            bool any = counts.Errors > 0 || counts.Warnings > 0;
            BoardStatusState state = counts.Errors > 0 ? BoardStatusState.Error : BoardStatusState.ToDo;

            shown.Block.IsVisible = any;
            shown.Title.IsVisible = any;
            shown.Title.Text = ContributeProblems.Said(counts);
            shown.Icon.Text = TabDrafts.StatusIconOf(state);

            TabDrafts.UseStateClasses(shown.Block, state);
            TabDrafts.UseStateClasses(shown.Icon, state);

            InPlaceList.Sync(
                shown.Rows.Children,
                panel.Problems ?? [],
                (control, item) => control is Grid { Tag: EntryProblem row } && row == item,
                this.StatusProblemRow);

            shown.More.Text = panel.ProblemsMore;
            shown.More.IsVisible = panel.ProblemsMore.Length > 0;
        }

        private Control StatusProblemRow(EntryProblem item)
        {
            (Border level, TextBlock levelText) = TabDrafts.Badge(item.Problem.Level == BoardProblemLevel.Error ? "error" : "warning");

            levelText.Text = ContributeProblems.LevelWord(item.Problem.Level);

            var text = new TextBlock
            {
                Text = item.Text,
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
                Margin = new Avalonia.Thickness(8, 0),
            };

            level.IsVisible = true;
            Grid.SetColumn(text, 1);

            var row = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*,Auto"), Tag = item, Children = { level, text } };

            if (item.ComponentLabel.Length > 0)
            {
                var open = new Button { Content = ContributeProblems.OpenLabel(item.ComponentLabel), VerticalAlignment = VerticalAlignment.Center };

                open.Click += (_, _) => this.RunGuarded(() => this.OpenComponentOfProblemAsync(item), "Open the component of a problem");
                Grid.SetColumn(open, 2);
                row.Children.Add(open);
            }

            return row;
        }

        private static void UseStateClasses(Control control, BoardStatusState state)
        {
            control.Classes.Set("done", state == BoardStatusState.Done);
            control.Classes.Set("todo", state == BoardStatusState.ToDo);
            control.Classes.Set("error", state == BoardStatusState.Error);
            control.Classes.Set("optional", state == BoardStatusState.Optional);
        }

        // Every panel, once, in the page's order.
        private void CreateStatusRows()
        {
            if (this.thisStatusRows.Count > 0)
            {
                return;
            }

            foreach (BoardStatusPart part in Enum.GetValues<BoardStatusPart>())
            {
                var icon = new TextBlock { Classes = { "StatusIcon" } };
                var title = new TextBlock { Classes = { "StatusTitle" }, TextWrapping = Avalonia.Media.TextWrapping.Wrap };
                var explanation = new TextBlock { Classes = { "StatusExplanation" } };
                var lines = new StackPanel { Spacing = 2, Margin = new Avalonia.Thickness(0, 6, 0, 0) };
                var extra = new StackPanel { Spacing = 8 };

                // Opens where the part is seen to - plain (CLAUDE.md, button colors).
                var action = new Button { Tag = part, VerticalAlignment = VerticalAlignment.Top, Margin = new Avalonia.Thickness(8, 0, 0, 0) };
                action.Click += this.OnStatusActionClick;

                var table = new Button { Content = BoardReadiness.TableAction, Tag = part, HorizontalAlignment = HorizontalAlignment.Left, Margin = new Avalonia.Thickness(0, 6, 0, 0), IsVisible = false };
                table.Click += this.OnStatusTableClick;

                StatusProblems problems = TabDrafts.MakeStatusProblems();

                // "Show in the table" is about what the checks found, so it stands with them.
                problems.Inner.Children.Add(table);

                var text = new StackPanel { Spacing = 1 };
                text.Children.Add(title);
                text.Children.Add(explanation);
                text.Children.Add(problems.Block);
                text.Children.Add(lines);
                text.Children.Add(extra);

                var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("28,*,Auto") };
                Grid.SetColumn(icon, 0);
                Grid.SetColumn(text, 1);
                Grid.SetColumn(action, 2);
                grid.Children.Add(icon);
                grid.Children.Add(text);
                grid.Children.Add(action);

                var panel = new Border { Classes = { "StatusPanel" }, Child = grid, Tag = part };
                this.StatusPanelsHost.Children.Add(panel);

                this.thisStatusRows[part] = new StatusRow(panel, icon, title, explanation, lines, action, table, extra, problems);
            }
        }

        private static StatusProblems MakeStatusProblems()
        {
            var icon = new TextBlock { Classes = { "StatusIcon" }, Margin = new Avalonia.Thickness(0, 2, 10, 0) };
            var title = new TextBlock { Classes = { "StatusTitle" } };
            var rows = new StackPanel { Spacing = 4 };
            var more = new TextBlock { Classes = { "StatusExplanation" }, IsVisible = false };
            var inner = new StackPanel { Spacing = 6, Children = { title, rows, more } };

            Grid.SetColumn(inner, 1);

            var block = new Border
            {
                Classes = { "StatusProblems" },
                IsVisible = false,
                Child = new Grid { ColumnDefinitions = new ColumnDefinitions("Auto,*"), Children = { icon, inner } },
            };

            return new StatusProblems(block, icon, title, inner, rows, more);
        }

        private void OnStatusActionClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: BoardStatusPart part })
            {
                this.DoStatusPart(part);
            }
        }

        // "Show in the table" under a panel's errors and warnings: the sheet they are on, marked there.
        private void OnStatusTableClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: BoardStatusPart part } &&
                this.thisStatusPanels.FirstOrDefault(panel => panel.Part == part)?.TableSheet is { } sheet)
            {
                this.RunGuarded(() => this.ShowTableOnSheetAsync(sheet), "Table");
            }
        }

        // A panel's button: the part it is about, as the left column opens it.
        internal void DoStatusPart(BoardStatusPart part)
        {
            switch (part)
            {
                case BoardStatusPart.Problems:
                    this.RunGuarded(() => this.OnCurrentDraftAsync(this.ShowTableAsync), "Table");
                    break;
                case BoardStatusPart.Name:
                    this.RunGuarded(this.OpenBoardNameAsync, "Board name");
                    break;
                case BoardStatusPart.Regions:
                    this.RunGuarded(() => this.ShowEntryAsync(ContributeEntry.Regions), "Regions");
                    break;
                case BoardStatusPart.Categories:
                    this.RunGuarded(() => this.ShowEntryAsync(ContributeEntry.Categories), "Categories");
                    break;
                case BoardStatusPart.KiCad:
                    this.RunGuarded(this.OpenKiCadAsync, "KiCad data");
                    break;
                case BoardStatusPart.Pictures:
                    this.RunGuarded(() => this.ShowEntryAsync(ContributeEntry.SchematicImages), "Schematic images");
                    break;
                case BoardStatusPart.Signals:
                    this.RunGuarded(() => this.ShowEntryAsync(ContributeEntry.ImportantSignals), "Important signals");
                    break;
                case BoardStatusPart.Components:
                    this.PickComponents();
                    break;
                case BoardStatusPart.Files:
                    this.RunGuarded(() => this.ShowEntryAsync(ContributeEntry.Files), "Files");
                    break;
                case BoardStatusPart.Links:
                    this.RunGuarded(() => this.ShowEntryAsync(ContributeEntry.Links), "Links");
                    break;
                case BoardStatusPart.Credits:
                    this.RunGuarded(() => this.ShowEntryAsync(ContributeEntry.Credits), "Credits");
                    break;
            }
        }

        // ###########################################################################################
        // The suggested matches under Schematic images - a row per pair, changed in place (a row made
        // again under a press threw that click away), "Use all N" for more than one.
        // ###########################################################################################
        private void ShowMatches(StackPanel text, IReadOnlyList<SchematicCadNameSuggestion> matches)
        {
            if (this.thisMatchControls is not { } controls)
            {
                if (matches.Count == 0)
                {
                    this.thisShownMatches = [];
                    return;
                }

                var panel = new StackPanel { Spacing = 4, Margin = new Avalonia.Thickness(0, 6, 0, 0) };
                var rows = new StackPanel { Spacing = 4 };
                var useAll = new Button { HorizontalAlignment = HorizontalAlignment.Left, Classes = { "Accept" } };
                ToolTip.SetTip(useAll, "Give every schematic image listed its suggested KiCad view - one step, so \"Undo\" takes them all back");
                useAll.Click += (_, _) => this.RunGuarded(() => this.UseSuggestedMatchesAsync(this.thisShownMatches), "Use all suggestions");

                panel.Children.Add(new TextBlock { Text = BoardReadiness.MatchesHeading, Classes = { "ReadinessDetail" } });
                panel.Children.Add(rows);
                panel.Children.Add(useAll);
                text.Children.Add(panel);

                controls = this.thisMatchControls = new MatchControls(panel, rows, useAll);
            }

            this.thisShownMatches = matches;
            controls.Panel.IsVisible = matches.Count > 0;

            InPlaceList.Sync(
                controls.Rows.Children,
                matches,
                (control, match) => control is Grid { Tag: SchematicCadNameSuggestion shown } && shown == match,
                this.MatchRow);

            controls.UseAll.Content = BoardReadiness.UseAllMatches(matches.Count);
            controls.UseAll.IsVisible = matches.Count > 1;
        }

        private Control MatchRow(SchematicCadNameSuggestion match)
        {
            var text = new TextBlock
            {
                Text = BoardReadiness.SuggestedMatch(match),
                TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                VerticalAlignment = VerticalAlignment.Center,
            };

            var use = new Button { Content = BoardReadiness.UseMatch, Tag = match, Margin = new Avalonia.Thickness(8, 0, 0, 0), Classes = { "Accept" } };
            ToolTip.SetTip(use, BoardReadiness.UseMatchTip(match));
            use.Click += (_, _) => this.RunGuarded(() => this.UseSuggestedMatchesAsync([match]), "Use suggestion");

            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("*,Auto"), Tag = match };
            Grid.SetColumn(use, 1);
            grid.Children.Add(text);
            grid.Children.Add(use);
            return grid;
        }

        // ###########################################################################################
        // A suggested match used - one, or every one shown: the draft opened as every change opens it,
        // and only the pairs it can still take (no CAD name on the picture, no picture with the view -
        // the draft can have moved) taken as ONE step named for them, saved at once under the
        // BusyOverlay. What came of it is said over the panels; a change reads the board on screen
        // again, so the Schematics tab draws the views.
        // ###########################################################################################
        internal Task UseSuggestedMatchesAsync(IReadOnlyList<SchematicCadNameSuggestion> wanted) =>
            this.OnCurrentDraftAsync(async entry =>
            {
                string identity = entry.ExcelDataFile;
                bool changed = false;

                ReadinessNote? note = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.MatchingPictures, async () =>
                {
                    if (await DraftEdits.OpenForChangeAsync(identity) is not DraftSession session)
                    {
                        return null;
                    }

                    List<SchematicCadNameSuggestion> open = SchematicCadNameSuggestions.StillOpen(session.Content.Board, wanted);

                    if (open.Count == 0)
                    {
                        return new ReadinessNote(identity, BoardReadiness.MatchesUsed([]), session.Content, GoesWhenSaved: false);
                    }

                    BoardContent before = session.Content;
                    IReadOnlyList<SchematicCadNameSuggestion> taken = [];

                    // One matching pass per working-out of the change (code review, 2026-10-28: two).
                    DraftWorkbookEditOutcome outcome = await DraftEdits.ApplyAndSaveAsync(
                        identity,
                        DraftStepWording.PicturesMatched(open.Select(match => match.SchematicName).ToList()),
                        board =>
                        {
                            BoardData result = SchematicCadNameSuggestions.With(board, open, out IReadOnlyList<SchematicCadNameSuggestion> took);
                            taken = took;
                            return result;
                        });

                    if (outcome != DraftWorkbookEditOutcome.Saved)
                    {
                        return new ReadinessNote(identity, DraftSaveWording.EditNotSaved(outcome), session.Content, GoesWhenSaved: session.SaveState != DraftSaveState.Saved);
                    }

                    changed = !ReferenceEquals(session.Content, before);

                    return new ReadinessNote(identity, BoardReadiness.MatchesUsed(changed ? taken : []), session.Content, GoesWhenSaved: false);
                });

                if (note is null)
                {
                    this.ShowDraftProblem(DraftSaveWording.EditNotSaved(DraftWorkbookEditOutcome.NoDraft));
                    return;
                }

                this.thisReadinessNote = note;

                if (changed)
                {
                    this.AfterDraftChanged(entry);
                }

                this.ShowReadiness();
            });

        // ###########################################################################################
        // "ADD N FROM KICAD" (the plan's step 6: "one press adds every KiCad component that isn't on the
        // board yet"): the draft opened as every change opens it (DraftEdits.OpenForChangeAsync - what
        // the editors hold handed over, a file written outside read), the board's KiCad data read again
        // (KiCadProjectLoader keeps the parse), and every component the draft lacks AS IT IS THEN added
        // as ONE step named for that many (code review, 2026-10-26: counted on the draft before it was
        // read again, the step named more than it added) - with pins and datasheet links, saved at once,
        // all under the BusyOverlay. What came of it is said over the panels; a draft that cannot
        // be opened, on the tab's problem line.
        //
        // *** SINCE 2026-11-01 IT IS THE KiCad IMPORT'S, NOT THE CHECKLIST'S (owner request: "it should
        // add all components when importing the KiCad data. Then you can delete this option from 'Board
        // status'"). *** "KiCad data" runs it after an import while "Add the KiCad components to the
        // board" is ticked, and from its summary's button (KiCadDataPanel.AddKiCadComponents) - its own
        // change, so the panel catches up with a change as it ends, the board on screen read again only
        // then; what came of it is said there, and over the panels when Board status comes back.
        //
        // What came of it - null when no draft could be opened - and whether the draft changed.
        // ###########################################################################################
        private async Task<(string? Note, bool Changed)> AddFromKiCadToAsync(HardwareBoardEntry entry)
        {
            string identity = entry.ExcelDataFile;
            bool changed = false;

            ReadinessNote? note = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.AddingKiCadComponents, async () =>
            {
                if (await DraftEdits.OpenForChangeAsync(identity) is not DraftSession session)
                {
                    return null;
                }

                if (await this.ReadKiCadForReadinessAsync(identity, "adding its components") is not { } root)
                {
                    return new ReadinessNote(identity, BoardReadiness.KiCadNotRead, session.Content, GoesWhenSaved: false);
                }

                List<KiCadBoardComponent> components = await Task.Run(() => KiCadComponentTable.Of(root));
                int missing = KiCadBoardComponents.NotOn(session.Content.Board, components).Count;

                if (missing == 0)
                {
                    return new ReadinessNote(identity, BoardReadiness.AddedFromKiCad(new KiCadComponentsAdded(0, 0, 0)), session.Content, GoesWhenSaved: false);
                }

                BoardContent before = session.Content;
                KiCadComponentsAdded added = new(0, 0, 0);

                DraftWorkbookEditOutcome outcome = await DraftEdits.ApplyAndSaveAsync(
                    identity,
                    DraftStepWording.ComponentsFromKiCad(missing),
                    board =>
                    {
                        BoardData result = KiCadBoardComponents.WithComponents(board, components, label => KiCadPinTable.PinsOf(root, label), out KiCadComponentsAdded counts);
                        added = counts;
                        return result;
                    });

                if (outcome != DraftWorkbookEditOutcome.Saved)
                {
                    // Taken back out; with something else still waiting to be saved, the words go
                    // once it is.
                    return new ReadinessNote(identity, DraftSaveWording.EditNotSaved(outcome), session.Content, GoesWhenSaved: session.SaveState != DraftSaveState.Saved);
                }

                // Saved - and changed, unless the change came to nothing as it was applied.
                changed = !ReferenceEquals(session.Content, before);

                return new ReadinessNote(
                    identity,
                    BoardReadiness.AddedFromKiCad(changed ? added : new KiCadComponentsAdded(0, 0, 0)),
                    session.Content,
                    GoesWhenSaved: false);
            });

            if (note is null)
            {
                this.ShowDraftProblem(DraftSaveWording.EditNotSaved(DraftWorkbookEditOutcome.NoDraft));
                return (null, false);
            }

            this.thisReadinessNote = note;

            this.ShowReadiness();
            return (note.Text, changed);
        }

        // ###########################################################################################
        // What each component lacks, under the Components panel: a group for each kind with any, each
        // component a button that opens it - the first ones named, the rest counted. Every group and
        // button made once, kept while it shows the same (InPlaceList.Sync).
        // ###########################################################################################
        private void ShowMissingComponents(BoardMissing missing)
        {
            MissingGroup[] groups = this.MissingControls();

            this.ShowMissingGroup(groups[0], BoardReadiness.NotInKiCadHeading(missing.NotInKiCadLabels.Count), missing.NotInKiCadLabels);
            this.ShowMissingGroup(groups[1], BoardReadiness.WithoutDescriptionHeading(missing.WithoutDescription.Count), missing.WithoutDescription);
            this.ShowMissingGroup(groups[TabDrafts.UnmarkedGroup], BoardReadiness.UnmarkedHeading(missing.Unmarked.Count), missing.Unmarked);
            this.ShowMissingGroup(groups[3], BoardReadiness.PinsWithoutBaselineHeading(missing.PinsWithoutBaseline.Count), missing.PinsWithoutBaseline);

            // The components marked nowhere stand in the frame of the panel's warnings, being one
            // (MissingControls), so the frame is there for them as the checks come in.
            Border frame = this.thisStatusRows[BoardStatusPart.Components].Problems.Block;

            frame.IsVisible = frame.IsVisible || groups[TabDrafts.UnmarkedGroup].Panel.IsVisible;
        }

        // Which of the groups below is the components marked nowhere.
        private const int UnmarkedGroup = 2;

        private void ShowMissingGroup(MissingGroup group, string heading, IReadOnlyList<string> labels)
        {
            group.Panel.IsVisible = labels.Count > 0;
            group.Heading.Text = heading;

            InPlaceList.Sync(
                group.Buttons.Children,
                labels.Take(BoardReadiness.ListedAtMost).ToList(),
                (control, label) => control is Button { Tag: string shown } && string.Equals(shown, label, StringComparison.Ordinal),
                this.MissingComponentButton);

            string more = BoardReadiness.MoreThanListed(labels.Count);
            group.More.Text = more;
            group.More.IsVisible = more.Length > 0;
        }

        // A component's button under the Components panel - it opens the component in the pane.
        private Control MissingComponentButton(string label)
        {
            var button = new Button { Content = label, Tag = label, Margin = new Avalonia.Thickness(0, 0, Handlers.Theming.ButtonRow.Gap, 4) };
            ToolTip.SetTip(button, $"Open {label}");
            button.Click += (_, _) => this.RunGuarded(() => this.OpenComponentAsync(label), "Open component");
            return button;
        }

        // The groups under the Components panel - made the first time they are shown.
        private MissingGroup[] MissingControls()
        {
            if (this.thisMissingGroups is { } made)
            {
                return made;
            }

            StatusRow components = this.thisStatusRows[BoardStatusPart.Components];
            StackPanel host = components.Extra;
            host.Margin = new Avalonia.Thickness(0, 8, 0, 0);

            made = new MissingGroup[4];

            for (int index = 0; index < made.Length; index++)
            {
                var panel = new StackPanel { Spacing = 4 };

                // In the panel of the warnings it is a sentence, with the panel's own title over it (owner
                // request, 2026-11-04); the others are headings of the panel's lines.
                var heading = new TextBlock
                {
                    FontWeight = index == TabDrafts.UnmarkedGroup ? Avalonia.Media.FontWeight.Normal : Avalonia.Media.FontWeight.SemiBold,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                };
                var buttons = new WrapPanel();
                var more = new TextBlock { FontSize = 11, Opacity = 0.75 };

                panel.Children.Add(heading);
                panel.Children.Add(buttons);
                panel.Children.Add(more);

                if (index == TabDrafts.UnmarkedGroup)
                {
                    components.Problems.Inner.Children.Insert(components.Problems.Inner.Children.IndexOf(components.Table), panel);
                }
                else
                {
                    host.Children.Add(panel);
                }

                made[index] = new MissingGroup(panel, heading, buttons, more);
            }

            this.thisMissingGroups = made;

            return made;
        }
    }
}
