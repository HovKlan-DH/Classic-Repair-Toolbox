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
    // GETTING A BOARD READY, AND WHAT A BOARD STILL LACKS (CRT 4.0 step 3, slice 3a - the plan: "Getting
    // this board ready ... On an existing board, the same panel shows What's missing: components with no
    // description, unmarked components, pins with no baseline").
    //
    // On the right of the Contribute tab while nothing is open there (the start panel):
    //
    //   - A NEW BOARD (one only a draft has) shows its checklist (BoardReadiness): each step done by what
    //     its draft holds, the first with something to do next, each with a button to the tool that
    //     does it - the "Schematic images" and "KiCad data" windows, the Pictures part (Mark components
    //     and Align KiCad on each picture), Add component, Important signals, Submit.
    //   - AN EXISTING BOARD shows what it lacks, each component a click away (BoardReadiness.Missing).
    //   - "ADD N FROM KICAD" (slice 3b) puts every component of the board's KiCad data the draft lacks
    //     into it in one step - with their values, categories, packages, pins and datasheet links.
    //   - SUGGESTED MATCHES (slice 3d - the plan: "Match: suggested pairs ("...Top.png" with "PCB Top")
    //     to confirm") are listed under the Match step, each picture with the KiCad view its name and
    //     its file's name say it shows (SchematicCadNameSuggestions); "Use" takes one, "Use all N"
    //     every one - one step either way.
    //
    // Worked out ONCE for each board shown and each time the right side goes back to the start panel -
    // never after every refresh of the drafts list, which runs whenever CRT comes to the front, and not
    // at all while the tab is off screen: then when it is shown (code review, 2026-10-26 - twice per
    // board load, in any tab). EVERY ROW AND BUTTON IS MADE ONCE AND CHANGED IN PLACE - the checklist's
    // and "What's missing"'s alike: the KiCad data arrives a moment after the rest, Main reads the board
    // again after a change, and a control made again under a press threw that click away (the trap the
    // drafts list met, 2026-10-10). A board's KiCad facts are kept for it, so the start panel coming
    // back shows them while they are read again rather than flicking to "not known yet".
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // Which time the lists were asked for - a read answers only the latest ask.
        private int thisReadinessAsk;

        // Asked for while the tab was off screen - worked out when it is shown.
        private bool thisReadinessStale;

        private readonly Dictionary<ReadinessStepKind, (TextBlock Icon, TextBlock Title, TextBlock Detail, Button Action, StackPanel Text)> thisReadinessRows = new();

        // ###########################################################################################
        // The matches suggested under "Match pictures to KiCad views" (slice 3d), made once in its
        // row: a line saying where they come from, a row per pair with its "Use", and "Use all N" -
        // and the pairs shown, which "Use all" takes.
        // ###########################################################################################
        private sealed record MatchControls(StackPanel Panel, StackPanel Rows, Button UseAll);

        private MatchControls? thisMatchControls;
        private IReadOnlyList<SchematicCadNameSuggestion> thisShownMatches = [];

        // The steps shown - what a step's button acts on - and the board they are of.
        private IReadOnlyList<ReadinessStep> thisReadinessSteps = [];
        private string? thisReadinessStepsBoard;

        // The last KiCad facts read, and the board they are of.
        private (string Board, BoardKiCadFacts Facts)? thisReadinessFacts;

        // ###########################################################################################
        // What a step's button did - said only while the draft is as it left it (code review,
        // 2026-10-26: after an Undo it still said the components were added): its board, the draft's
        // content it was said of, and - for a save that failed while something else waited to be
        // saved - whether it goes once the draft is saved.
        // ###########################################################################################
        private sealed record ReadinessNote(string Board, string Text, BoardContent Content, bool GoesWhenSaved);

        private ReadinessNote? thisReadinessNote;

        // "What's missing"'s groups, made once: a heading, the first labels as buttons, "and N more".
        private sealed record MissingGroup(StackPanel Panel, TextBlock Heading, WrapPanel Buttons, TextBlock More);

        private MissingGroup[]? thisMissingGroups;
        private TextBlock? thisNothingMissingText;

        // The checklist's last work - for a test to wait for.
        private Task thisReadinessShown = Task.CompletedTask;

        // Stands in for reading a board's KiCad data - a test has no KiCad files.
        internal Func<string, Task<KiCadProjectRoot?>>? BoardKiCadOverrideForTests { get; set; }

        internal IReadOnlyList<ReadinessStep> ReadinessForTests => this.thisReadinessSteps;

        internal string? ReadinessNoteForTests => this.ReadinessNoteText.IsVisible ? this.ReadinessNoteText.Text : null;

        internal BoardMissing? MissingForTests { get; private set; }

        internal Task ReadinessShownForTests => this.thisReadinessShown;

        // How many times the lists were worked out - each is a session opened and KiCad data read.
        internal int ReadinessAsksForTests => this.thisReadinessAsk;

        // Stands in for whether the tab is on screen - a test's tab is in no window.
        internal Func<bool>? IsOnScreenOverrideForTests { get; set; }

        // ###########################################################################################
        // The start panel shown or hidden - and its lists worked out again each time it comes back
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
        // The checklist of a new board, or what an existing board lacks - for the board shown.
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
                this.HideSteps();
                this.MissingPanel.IsVisible = false;
                return;
            }

            if (!entry.IsDraftOnly)
            {
                // The draft held in memory when there is one - the pane's saves can leave the board
                // shown behind until Main reads it again.
                BoardData current = DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile)?.Content.Board ?? board;

                this.HideSteps();
                this.ShowMissing(BoardReadiness.Missing(current));
                return;
            }

            this.MissingPanel.IsVisible = false;

            // Another board's checklist goes at once - never left standing under this board's header,
            // its buttons acting on this board (code review, 2026-10-26).
            if (!string.Equals(this.thisReadinessStepsBoard, entry.ExcelDataFile, StringComparison.OrdinalIgnoreCase))
            {
                this.HideSteps();
            }

            this.thisReadinessShown = this.ShowStepsOfAsync(entry, ask);
        }

        private void HideSteps()
        {
            this.ReadinessPanel.IsVisible = false;
            this.thisReadinessSteps = [];
            this.thisReadinessStepsBoard = null;
        }

        // ###########################################################################################
        // The new board's steps: first from its draft and the KiCad facts last read for it, then again
        // once its KiCad data is read (the views and the components it has - KiCadProjectLoader keeps
        // the parse, so the board on screen has usually read it already). A newer ask wins; a draft
        // that cannot be opened shows no checklist.
        // ###########################################################################################
        private async Task ShowStepsOfAsync(HardwareBoardEntry entry, int ask)
        {
            try
            {
                string draftsRoot = DraftManager.DraftsRoot;
                string identity = entry.ExcelDataFile;
                string folder = DraftFolderLayout.GetBoardFolder(draftsRoot, identity);

                DraftSession? session = DraftManager.Sessions.Find(draftsRoot, identity)
                                        ?? await Task.Run(() => DraftManager.Sessions.Get(draftsRoot, identity));

                if (ask != this.thisReadinessAsk)
                {
                    return;
                }

                if (session is null)
                {
                    this.HideSteps();
                    return;
                }

                bool hasKiCad = session.ListedFiles.Keys.Any(SubmissionKiCadFiles.IsInKiCadFolder) ||
                                await Task.Run(() => DraftFiles.ArrivedKiCadNames(folder).Count > 0);

                if (ask != this.thisReadinessAsk)
                {
                    return;
                }

                BoardKiCadFacts? known = hasKiCad && this.thisReadinessFacts is { } last && string.Equals(last.Board, identity, StringComparison.OrdinalIgnoreCase)
                    ? last.Facts
                    : null;

                this.ShowSteps(identity, session, BoardReadiness.Steps(entry.ToString(), session.Content, hasKiCad, known));

                if (!hasKiCad)
                {
                    return;
                }

                KiCadProjectRoot? root = await this.ReadKiCadForReadinessAsync(identity, "getting the board ready");

                if (ask != this.thisReadinessAsk)
                {
                    return;
                }

                BoardKiCadFacts facts = await Task.Run(() => BoardKiCadFacts.Of(root));

                if (ask == this.thisReadinessAsk)
                {
                    this.thisReadinessFacts = (identity, facts);
                    this.ShowSteps(identity, session, BoardReadiness.Steps(entry.ToString(), session.Content, hasKiCad, facts));
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not work out how ready [{entry.ExcelDataFile}] is - [{ex.Message}]");

                if (ask == this.thisReadinessAsk)
                {
                    this.HideSteps();
                }
            }
        }

        // The board's KiCad project, as the Schematics tab draws it (ReadBoardKiCadAsync) - null without.
        private Task<KiCadProjectRoot?> ReadKiCadForReadinessAsync(string identity, string what) =>
            this.BoardKiCadOverrideForTests is { } read
                ? read(identity)
                : TabDrafts.ReadBoardKiCadAsync(DataManager.DataRoot, DraftFolderLayout.GetBoardFolder(DraftManager.DraftsRoot, identity), identity, what);

        // The steps shown - each row made once, its words and state changed in place.
        private void ShowSteps(string board, DraftSession session, IReadOnlyList<ReadinessStep> steps)
        {
            this.thisReadinessSteps = steps;
            this.thisReadinessStepsBoard = board;

            foreach (ReadinessStep step in steps)
            {
                (TextBlock icon, TextBlock title, TextBlock detail, Button action, StackPanel text) = this.ReadinessRowFor(step.Kind);

                icon.Text = step.State switch
                {
                    ReadinessState.Done => "",
                    ReadinessState.Next => "",
                    ReadinessState.Later => "",
                    _ => "",
                };

                icon.Classes.Set("done", step.State == ReadinessState.Done);
                icon.Classes.Set("next", step.State == ReadinessState.Next);
                title.Classes.Set("next", step.State == ReadinessState.Next);

                title.Text = string.Create(CultureInfo.InvariantCulture, $"{step.Number}  {step.Title}");
                detail.Text = step.Detail;

                string label = BoardReadiness.ActionOf(step);
                action.Content = label;
                action.IsVisible = label.Length > 0;

                if (step.Kind == ReadinessStepKind.Match)
                {
                    this.ShowMatches(text, step.Matches ?? []);
                }
            }

            bool noteShown = this.thisReadinessNote is { } note &&
                             string.Equals(note.Board, board, StringComparison.OrdinalIgnoreCase) &&
                             ReferenceEquals(note.Content, session.Content) &&
                             !(note.GoesWhenSaved && session.SaveState == DraftSaveState.Saved);

            this.ReadinessNoteText.Text = noteShown ? this.thisReadinessNote!.Text : null;
            this.ReadinessNoteText.IsVisible = noteShown;
            this.ReadinessPanel.IsVisible = true;
        }

        // A step's row - made the first time it is shown.
        private (TextBlock Icon, TextBlock Title, TextBlock Detail, Button Action, StackPanel Text) ReadinessRowFor(ReadinessStepKind kind)
        {
            if (this.thisReadinessRows.TryGetValue(kind, out var row))
            {
                return row;
            }

            var icon = new TextBlock { Classes = { "ReadinessIcon" } };
            var title = new TextBlock { Classes = { "ReadinessTitle" }, TextWrapping = Avalonia.Media.TextWrapping.Wrap };
            var detail = new TextBlock { Classes = { "ReadinessDetail" } };

            // A button that goes ahead and makes something is green (CLAUDE.md, button colours).
            var action = new Button { Tag = kind, VerticalAlignment = VerticalAlignment.Top, Margin = new Avalonia.Thickness(8, 0, 0, 0) };
            action.Classes.Set("Accept", kind is ReadinessStepKind.Components or ReadinessStepKind.Submit);
            action.Click += this.OnReadinessActionClick;

            var text = new StackPanel { Spacing = 1 };
            text.Children.Add(title);
            text.Children.Add(detail);

            var grid = new Grid { ColumnDefinitions = new ColumnDefinitions("26,*,Auto") };
            Grid.SetColumn(icon, 0);
            Grid.SetColumn(text, 1);
            Grid.SetColumn(action, 2);
            grid.Children.Add(icon);
            grid.Children.Add(text);
            grid.Children.Add(action);

            this.ReadinessStepsPanel.Children.Add(grid);

            row = (icon, title, detail, action, text);
            this.thisReadinessRows[kind] = row;
            return row;
        }

        // ###########################################################################################
        // The suggested matches under the Match step - a row per pair, changed in place (a row made
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
                ToolTip.SetTip(useAll, "Give every picture listed its suggested KiCad view - one step, so Undo takes them all back");
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
        // BusyOverlay. What came of it is said under the checklist; a change reads the board on screen
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

        private void OnReadinessActionClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ReadinessStepKind kind })
            {
                this.DoReadinessStep(kind);
            }
        }

        // ###########################################################################################
        // A step's button: the tool that does the step. Matching a picture to a KiCad view is its CAD
        // name in "Schematic images"; aligning and marking are on each picture's card in the Pictures
        // part ("Align KiCad", "Mark components").
        // ###########################################################################################
        internal void DoReadinessStep(ReadinessStepKind kind)
        {
            switch (kind)
            {
                case ReadinessStepKind.Pictures:
                case ReadinessStepKind.Match:
                    this.ManagePictures();
                    break;
                case ReadinessStepKind.KiCad:
                    this.RunGuarded(() => this.OnCurrentDraftAsync(entry => this.ManageFilesAsync(entry, BoardFilesSection.KiCadData)), "KiCad data");
                    break;
                case ReadinessStepKind.Align:
                case ReadinessStepKind.Mark:
                    this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.Pictures), "Pictures");
                    break;
                case ReadinessStepKind.Components when this.thisReadinessSteps.Any(step => step.Kind == ReadinessStepKind.Components && step.FromKiCad > 0):
                    this.RunGuarded(this.AddFromKiCadAsync, "Add from KiCad");
                    break;
                case ReadinessStepKind.Components:
                    this.RunGuarded(this.AddComponentAsync, "Add component");
                    break;
                case ReadinessStepKind.Signals:
                    this.RunGuarded(() => this.OpenBoardPartAsync(PaneBoardPart.ImportantSignals), "Important signals");
                    break;
                case ReadinessStepKind.Submit:
                    this.RunGuarded(() => this.OnCurrentDraftAsync(this.SubmitAsync), "Submit");
                    break;
            }
        }

        // ###########################################################################################
        // "ADD N FROM KICAD" (the plan's step 6: "one press adds every KiCad component that isn't on the
        // board yet"): the draft opened as every change opens it (DraftEdits.OpenForChangeAsync - what
        // the editors hold handed over, a file written outside read), the board's KiCad data read again
        // (KiCadProjectLoader keeps the parse), and every component the draft lacks AS IT IS THEN added
        // as ONE step named for that many (code review, 2026-10-26: counted on the draft before it was
        // read again, the step named more than it added) - with pins and datasheet links, saved at once,
        // all under the BusyOverlay. What came of it is said under the checklist; a draft that cannot
        // be opened, on the tab's problem line. Only a change reads the board on screen again.
        // ###########################################################################################
        internal Task AddFromKiCadAsync() =>
            this.OnCurrentDraftAsync(async entry =>
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
        // What an existing board lacks: a group for each kind with any, each component a button that
        // opens it - the first ones named, the rest counted. Every group and button made once, kept
        // while it shows the same (InPlaceList.Sync).
        // ###########################################################################################
        private void ShowMissing(BoardMissing missing)
        {
            this.MissingForTests = missing;

            (MissingGroup[] groups, TextBlock nothing) = this.MissingControls();

            nothing.IsVisible = missing.IsEmpty;
            this.ShowMissingGroup(groups[0], BoardReadiness.WithoutDescriptionHeading(missing.WithoutDescription.Count), missing.WithoutDescription);
            this.ShowMissingGroup(groups[1], BoardReadiness.UnmarkedHeading(missing.Unmarked.Count), missing.Unmarked);
            this.ShowMissingGroup(groups[2], BoardReadiness.PinsWithoutBaselineHeading(missing.PinsWithoutBaseline.Count), missing.PinsWithoutBaseline);

            this.MissingPanel.IsVisible = true;
        }

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

        // A component's button in "What's missing" - it opens the component in the pane.
        private Control MissingComponentButton(string label)
        {
            var button = new Button { Content = label, Tag = label, Margin = new Avalonia.Thickness(0, 0, 6, 4) };
            ToolTip.SetTip(button, $"Open {label}");
            button.Click += (_, _) => this.RunGuarded(() => this.OpenComponentAsync(label), "Open component");
            return button;
        }

        // "What's missing"'s groups and its "nothing missing" line - made the first time they are shown.
        private (MissingGroup[] Groups, TextBlock Nothing) MissingControls()
        {
            if (this.thisMissingGroups is { } made && this.thisNothingMissingText is { } nothing)
            {
                return (made, nothing);
            }

            nothing = new TextBlock { Text = BoardReadiness.NothingMissing, TextWrapping = Avalonia.Media.TextWrapping.Wrap };
            this.MissingGroupsPanel.Children.Add(nothing);

            made = new MissingGroup[3];

            for (int index = 0; index < made.Length; index++)
            {
                var panel = new StackPanel { Spacing = 4 };
                var heading = new TextBlock { FontWeight = Avalonia.Media.FontWeight.SemiBold };
                var buttons = new WrapPanel();
                var more = new TextBlock { FontSize = 11, Opacity = 0.75 };

                panel.Children.Add(heading);
                panel.Children.Add(buttons);
                panel.Children.Add(more);
                this.MissingGroupsPanel.Children.Add(panel);

                made[index] = new MissingGroup(panel, heading, buttons, more);
            }

            this.thisMissingGroups = made;
            this.thisNothingMissingText = nothing;

            return (made, nothing);
        }
    }
}
