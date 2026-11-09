using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE LEFT COLUMN'S ENTRIES AND THE TWO VIEWS OF THEM (owner requests, 2026-11-03).
    //
    //   - "Overview" first, then "Board" and "Components": "now we do not show the list of components
    //     directly in the 'Contribute' tab, then it leaves some room for that left-side menu ... This then
    //     also means that we should have no tabs in the 'component view'". A component's page is an entry,
    //     with how many it holds and marked while something on it stops Submit. They are about THE COMPONENT
    //     PICKED in CRT's own list (thisPickedComponent) - which stays picked whatever else is open, until
    //     another board is shown. None picked, the entry says where to pick one (NoComponentPanel).
    //   - "UI" OR "TABLE" (a joined switch in the header; "Table view" until then): "when I select e.g. 'Board:
    //     Schematic images' then I will see the table for 'Board schematics' in the content area, instead of
    //     the normal UI elements". An entry with a table shows it (ContributeEntries - every sheet has an
    //     entry, "Component images" two), the sheet tabs gone (BoardTableEditor.SheetParts.cs) and the left
    //     column staying. An entry that is no sheet shows its page, saying so (NoTableNotice). Switching keeps
    //     the entry. A component picked goes to its first row on the sheet shown.
    //
    // WHICH ENTRY IS OPEN IS WORKED OUT FROM WHAT IS OPEN, every time (EntryOpen) - never a flag beside the
    // pane. The only things kept are what no control can say: the entry a table was asked for, and the
    // component page asked for with none picked.
    //
    // Part of the TabDrafts partial class - see TabDrafts.axaml.cs for the file map.
    // ###########################################################################################
    public partial class TabDrafts
    {
        // "Table" picked: the entries with a table show it. For the run.
        private bool thisTableMode;

        // The entry the open table shows - it follows the table as a pill, a search or an undo moves it.
        private ContributeEntry thisEntryInTable = ContributeEntry.SchematicImages;

        // A component's page asked for with no component picked - NoComponentPanel says where to pick one.
        private ComponentPage? thisNoComponentPage;

        // The component picked in the list, and shown selected there whatever is open on the right.
        private string thisPickedComponent = string.Empty;

        // ###########################################################################################
        // *** SEVERAL PICKED, EDITED AS ONE (owner request, 2026-11-08; ComponentPane.Several.cs). *** Every component
        // picked while more than one is - thisPickedComponent the first of them - and empty while one, or none, is. The
        // pane shows them together, a component's page lists the problems of every one of them, and the table shows their
        // rows alone (BoardTableEditor.Components.cs).
        // ###########################################################################################
        private List<string> thisPickedSeveral = [];

        // The components picked: the several, the one, or none.
        private IReadOnlyList<string> PickedLabels =>
            this.thisPickedSeveral.Count > 1 ? this.thisPickedSeveral
            : this.thisPickedComponent.Length > 0 ? [this.thisPickedComponent]
            : [];

        internal IReadOnlyList<string> PickedComponentsForTests => this.PickedLabels;

        // Opening it right now: a click selects and is tapped too, and only one should open it.
        private string? thisOpeningComponent;

        internal bool IsTableViewPickedForTests => this.thisTableMode;

        internal string PickedComponentForTests => this.thisPickedComponent;

        internal bool IsNoComponentPageShownForTests => this.NoComponentPanel.IsVisible;

        internal string NoComponentTextForTests => this.NoComponentText.Text ?? string.Empty;

        internal bool IsNoTableNoticeShownForTests => this.NoTableNotice.IsVisible;

        private IEnumerable<(Button Button, ContributeEntry Entry)> NavEntries() =>
        [
            (this.StatusNavButton, ContributeEntry.Overview),
            (this.NameNavButton, ContributeEntry.Name),
            (this.RegionsNavButton, ContributeEntry.Regions),
            (this.KiCadNavButton, ContributeEntry.KiCadFiles),
            (this.SignalsNavButton, ContributeEntry.ImportantSignals),
            (this.PicturesNavButton, ContributeEntry.SchematicImages),
            (this.CategoriesNavButton, ContributeEntry.Categories),
            (this.FilesNavButton, ContributeEntry.Files),
            (this.LinksNavButton, ContributeEntry.Links),
            (this.CreditsNavButton, ContributeEntry.Credits),
            (this.BasicsNavButton, ContributeEntry.ComponentBasics),
            (this.ReplacementsNavButton, ContributeEntry.ComponentReplacements),
            (this.PinoutNavButton, ContributeEntry.ComponentPinout),
            (this.ComponentPicturesNavButton, ContributeEntry.ComponentPictures),
            (this.BaselinesNavButton, ContributeEntry.ComponentBaselines),
            (this.ComponentFilesNavButton, ContributeEntry.ComponentFiles),
            (this.ComponentLinksNavButton, ContributeEntry.ComponentLinks),
        ];

        private (TextBlock Label, TextBlock Count, Button Button) NavOf(ComponentPage page) => page switch
        {
            ComponentPage.Basics => (this.BasicsNavLabel, this.BasicsCountText, this.BasicsNavButton),
            ComponentPage.Replacements => (this.ReplacementsNavLabel, this.ReplacementsCountText, this.ReplacementsNavButton),
            ComponentPage.Pinout => (this.PinoutNavLabel, this.PinoutCountText, this.PinoutNavButton),
            ComponentPage.Pictures => (this.ComponentPicturesNavLabel, this.ComponentPicturesCountText, this.ComponentPicturesNavButton),
            ComponentPage.Baselines => (this.BaselinesNavLabel, this.BaselinesCountText, this.BaselinesNavButton),
            ComponentPage.Files => (this.ComponentFilesNavLabel, this.ComponentFilesCountText, this.ComponentFilesNavButton),
            _ => (this.ComponentLinksNavLabel, this.ComponentLinksCountText, this.ComponentLinksNavButton),
        };

        private void WireEntries()
        {
            this.NoTableNoticeText.Text = ContributeBoardParts.NoTableForEntry;
            this.TableEditor.SheetsPickedElsewhere = true;
            this.TableEditor.SheetShown += (_, _) => this.OnTableSheetShown();
            this.TableEditor.SheetsShownChanged += (_, _) => this.ShowTableEntriesEnabled();
            this.ComponentPane.PagesChanged += (_, _) => this.ShowComponentPagesInNav();
            this.WireProblems();
        }

        // ---- which entry is open ---------------------------------------------------------------------

        // The entry open on the right - null for the new component's form, which is none.
        private ContributeEntry? EntryOpen()
        {
            if (this.IsTableOpen)
            {
                return this.thisEntryInTable;
            }

            if (this.IsComponentOpen)
            {
                return this.ComponentPane.BoardPart is { } part
                    ? ContributeEntries.OfPart(part)
                    : this.ComponentPane.IsNewComponentForm ? null : ContributeEntries.OfPage(this.ComponentPane.Page);
            }

            if (this.IsKiCadOpen)
            {
                return ContributeEntry.KiCadFiles;
            }

            if (this.thisBoardNamePicked)
            {
                return ContributeEntry.Name;
            }

            return this.thisNoComponentPage is { } page ? ContributeEntries.OfPage(page) : ContributeEntry.Overview;
        }

        internal ContributeEntry? EntryOpenForTests => this.EntryOpen();

        // The entry of the left column shown selected.
        internal string? SelectedNavEntryForTests =>
            this.NavEntries().Select(entry => entry.Button).FirstOrDefault(button => button.Classes.Contains("selected"))?.Name;

        private void OnNavEntryClick(object? sender, RoutedEventArgs e)
        {
            foreach ((Button button, ContributeEntry entry) in this.NavEntries())
            {
                if (ReferenceEquals(sender, button))
                {
                    this.RunGuarded(() => this.ShowEntryAsync(entry), entry.ToString());
                    return;
                }
            }
        }

        // ###########################################################################################
        // An entry picked, as the view picked shows it: its table, else its page.
        // ###########################################################################################
        internal async Task ShowEntryAsync(ContributeEntry entry)
        {
            if (this.thisTableMode && ContributeEntries.TableOf(entry) is not null)
            {
                await this.ShowEntryInTableAsync(entry, takeFocus: true);
            }
            else if (ContributeEntries.PageOf(entry) is { } page)
            {
                await this.ShowComponentPageAsync(page);
            }
            else if (ContributeEntries.PartOf(entry) is { } part)
            {
                await this.OpenBoardPartAsync(part);
            }
            else if (entry == ContributeEntry.Name)
            {
                await this.OpenBoardNameAsync();
            }
            else if (entry == ContributeEntry.KiCadFiles)
            {
                await this.OpenKiCadAsync();
            }
            else
            {
                await this.ShowBoardStatusAsync();
            }
        }

        // ---- a component's page ----------------------------------------------------------------------

        // ###########################################################################################
        // The page of the component picked. With the pane on it, only the page changes; with none
        // picked, the page says where to pick one, and a component picked then opens on that page.
        // ###########################################################################################
        private async Task ShowComponentPageAsync(ComponentPage page)
        {
            // "Mark all" from the other tabs is shown on no page either (second code review, 2026-11-09: a page picked after
            // it opened every component of the board as one) - the page says to pick some (TakePickFromOtherTabs).
            if (this.thisPickedComponent.Length == 0 || this.IsEveryTakenOver)
            {
                await this.ShowNoComponentPageAsync(page);
                return;
            }

            if (this.IsPaneOnPicked())
            {
                this.ComponentPane.ShowPage(page);
                this.ShowNavSelection();
                return;
            }

            // Not while a picture's name holds the pane open: the page stays what it was.
            if (this.IsComponentOpen && this.ComponentPane.KeepsRefusedPictureName())
            {
                return;
            }

            this.ComponentPane.ShowPage(page);
            await this.OpenPickedAsync();
            this.ShowNavSelection();
        }

        private async Task ShowNoComponentPageAsync(ComponentPage page)
        {
            if (this.IsComponentOpen && this.ComponentPane.KeepsRefusedPictureName())
            {
                return;
            }

            // Said before anything closes: a close brings "Overview" back otherwise, worked out for nothing.
            ComponentPage? before = this.thisNoComponentPage;
            this.thisNoComponentPage = page;
            this.thisBoardNamePicked = false;

            if (this.IsComponentOpen)
            {
                this.CloseComponent();
            }

            await this.CloseKiCadAsync();

            if (this.IsTableOpen && !await this.CloseTableAsync())
            {
                this.thisNoComponentPage = before;
                return;
            }

            this.ComponentPane.ShowPage(page);
            this.ShowWorkspace();
        }

        private bool IsPaneOnComponent(string label) =>
            this.IsComponentOpen &&
            this.ComponentPane.BoardPart is null &&
            !this.ComponentPane.IsNewComponentForm &&
            !this.ComponentPane.IsSeveral &&
            string.Equals(this.ComponentPane.Label, label, StringComparison.OrdinalIgnoreCase);

        // The pane on exactly the components picked - one, or the several.
        private bool IsPaneOnPicked() =>
            this.IsComponentOpen &&
            this.ComponentPane.BoardPart is null &&
            !this.ComponentPane.IsNewComponentForm &&
            this.PickedLabels.Count > 0 &&
            this.ComponentPane.Labels.SequenceEqual(this.PickedLabels, StringComparer.OrdinalIgnoreCase);

        // The components picked opened in the pane - several together.
        private Task OpenPickedAsync() =>
            this.PickedLabels.Count > 1 ? this.OpenSeveralAsync(this.PickedLabels.ToList()) : this.OpenComponentAsync(this.thisPickedComponent);

        // ###########################################################################################
        // The component picked, as what is open says: the pane's while it is on one - it follows a rename
        // there - and gone once the board has no such component any more.
        //
        // NEITHER THE BOARD ON SCREEN NOR THE DRAFT HOLDS IT, THAT IS (code review, 2026-11-04). Once the pane
        // or the table closes on a published board the column reads the board on screen, which has not caught
        // up with their saves yet: a component just renamed or added is not on it, and the pick went - listed
        // again a moment later, but no longer selected, and its pages saying to pick one.
        // ###########################################################################################
        private void FollowPickedComponent()
        {
            // Not while a pick is still opening: the pane shows the pick before it, and following it would take the pick back.
            if (this.thisPicksOpening == 0 &&
                this.IsComponentOpen && this.ComponentPane.BoardPart is null && !this.ComponentPane.IsNewComponentForm && this.ComponentPane.Label.Length > 0)
            {
                this.thisPickedComponent = this.ComponentPane.Label;
                this.thisPickedSeveral = this.ComponentPane.IsSeveral ? this.ComponentPane.Labels.ToList() : [];
            }
            else if (this.PickedLabels.Count > 0 && this.BoardForList() is { } board)
            {
                BoardData? draft = this.thisCurrentEntry is { } entry &&
                                   DraftManager.Sessions.Find(DraftManager.DraftsRoot, entry.ExcelDataFile) is { } session
                    ? session.Content.Board
                    : null;

                bool Held(string label) => TabDrafts.Holds(board, label) || (draft is not null && TabDrafts.Holds(draft, label));

                List<string> held = this.PickedLabels.Where(Held).ToList();

                if (held.Count != this.PickedLabels.Count)
                {
                    bool every = this.IsEveryTakenOver;

                    this.thisPickedComponent = held.FirstOrDefault() ?? string.Empty;
                    this.thisPickedSeveral = held.Count > 1 ? held : [];

                    // "Mark all" from the other tabs stays what it is with a component gone from it.
                    if (every)
                    {
                        this.thisEveryTakenOver = this.thisPickedSeveral;
                    }
                }
            }
        }

        private static bool Holds(BoardData board, string label) =>
            board.Components.Any(row => ComponentPaneEdits.IsOfComponent(row.BoardLabel, label));

        // ###########################################################################################
        // "Clear" over the lists: no component picked. Its page stays what the right shows, as the entry
        // asked for with none picked.
        // ###########################################################################################
        internal void LetGoOfPickedComponent()
        {
            this.thisPickedComponent = string.Empty;
            this.thisPickedSeveral = [];
            this.thisPickGeneration++;

            if (this.IsComponentOpen && this.ComponentPane.BoardPart is null && !this.ComponentPane.IsNewComponentForm)
            {
                this.thisNoComponentPage = this.ComponentPane.Page;
                this.CloseComponent();
            }

            this.ShowWorkspace();
        }

        // What each of the component's pages holds, beside its entry - nothing with none picked.
        private void ShowComponentPagesInNav()
        {
            IReadOnlyDictionary<ComponentPage, ComponentPageFact>? facts =
                this.thisPickedComponent.Length == 0 ? null
                : this.IsPaneOnPicked() && this.ComponentPane.PagesShown is { } shown ? shown
                : this.PickedLabels.Count > 1 ? null
                : this.BoardForList() is { } board ? ContributeEntries.FactsOf(board, this.thisPickedComponent, this.thisCurrentRegion)
                : null;

            foreach (ComponentPage page in Enum.GetValues<ComponentPage>())
            {
                ComponentPageFact fact = facts?.GetValueOrDefault(page) ?? default;
                (TextBlock label, TextBlock count, Button button) = this.NavOf(page);

                count.Text = facts is null ? string.Empty : ContributeBoardParts.CountText(fact.Count);
                label.Classes.Set("error", fact.HasError);
                ToolTip.SetTip(button, fact.HasError ? ComponentPaneWording.PageHasErrorTip : null);
            }
        }

        // ---- "UI" or "Table" -------------------------------------------------------------------------

        private void OnViewClick(object? sender, RoutedEventArgs e) =>
            this.RunGuarded(() => this.ShowViewAsync(table: ReferenceEquals(sender, this.TableViewButton)), "UI or Table");

        // ###########################################################################################
        // The view picked, the same entry staying open. One that cannot be followed - no draft could be
        // made for the table, or a table that cannot be saved and stays - goes back to what is so.
        //
        // The new component's form is no entry (EntryOpen), so it has no view to switch (code review, 2026-11-04):
        // "Table" there took the mode with nothing to show, and once added the component's page stood under
        // "Table", where a further press did nothing. The switch is off while the form is open (ShowViewSwitch).
        // ###########################################################################################
        internal async Task ShowViewAsync(bool table)
        {
            ContributeEntry? shown = this.EntryOpen();

            if (table == this.thisTableMode || shown is null)
            {
                this.ShowViewSwitch();
                return;
            }

            this.thisTableMode = table;

            if (shown is { } entry && ContributeEntries.TableOf(entry) is not null)
            {
                await this.ShowEntryAsync(entry);

                this.thisTableMode = this.IsTableOpen;
            }

            this.ShowWorkspace();
        }

        private void ShowViewSwitch()
        {
            bool hasEntry = this.EntryOpen() is not null;

            this.UiViewButton.IsEnabled = hasEntry;
            this.TableViewButton.IsEnabled = hasEntry;

            this.UiViewButton.Classes.Set("Selected", !this.thisTableMode);
            this.TableViewButton.Classes.Set("Selected", this.thisTableMode);

            this.NoTableNotice.IsVisible = this.thisTableMode &&
                                           !this.IsTableOpen &&
                                           this.EntryOpen() is { } entry &&
                                           ContributeEntries.TableOf(entry) is null;
        }

        // ###########################################################################################
        // An entry's table: on a draft, made first when the board has none, as every action on the board.
        // A component's sheet goes to the rows of the component picked.
        // ###########################################################################################
        private Task ShowEntryInTableAsync(ContributeEntry entry, bool takeFocus) =>
            this.OnCurrentDraftAsync(async draft =>
            {
                if (!this.IsTableOpenFor(draft))
                {
                    await this.OpenTableAsync(draft, takeFocus);
                }

                if (!this.IsTableOpenFor(draft) || ContributeEntries.TableOf(entry) is not { } table)
                {
                    return;
                }

                this.TableEditor.ShowSheetPart(table.Sheet, table.Part);

                // Their rows alone, and the cursor to the first one's (owner request, 2026-11-08) - every row for "Mark all"'s.
                this.TableEditor.ShowOnlyComponents(this.TableLabels);

                if (ContributeEntries.PageOf(entry) is not null && this.thisPickedComponent.Length > 0 && !this.IsEveryTakenOver)
                {
                    this.TableEditor.GoToRowsOfComponent(this.thisPickedComponent);
                }

                this.ShowNavSelection();

                if (takeFocus)
                {
                    this.FocusTableIfOpen();
                }
            });

        // The table moved - by a pick here, a pill, a search or an undo: the column follows it.
        private void OnTableSheetShown()
        {
            if (ContributeEntries.EntryOfTable(this.TableEditor.CurrentSheet?.Name, this.TableEditor.SheetPart) is { } entry)
            {
                this.thisEntryInTable = entry;
            }

            this.ShowNavSelection();
        }

        // ###########################################################################################
        // While a pill or a search narrows the table, the entries it would show nothing of are off - as
        // the tab of such a sheet went (BoardTableDocument.SheetsShown). The entry open always stays.
        //
        // So are the entries that are no sheet ("Overview", "Name", "Regions", "KiCad: Files", "Categories" - owner request,
        // 2026-11-07: "it should also disable the non-table menu entries, as it does not search in those"): neither a pill nor
        // the search looks in them, so they hold nothing it found.
        // ###########################################################################################
        private void ShowTableEntriesEnabled()
        {
            Func<string, BoardTableSheetPart, bool>? shows = this.IsTableOpen ? this.TableEditor.WhatShows() : null;
            bool filtered = shows is not null && this.TableEditor.IsFilteredByPillsOrSearch;

            foreach ((Button button, ContributeEntry entry) in this.NavEntries())
            {
                button.IsEnabled = shows is null ||
                                   (ContributeEntries.TableOf(entry) is { } table ? shows(table.Sheet, table.Part) : !filtered);
            }
        }

        // ---- a click on the component already picked -------------------------------------------------

        // ###########################################################################################
        // No selection changes, so the list says nothing - and with a board part open, the component
        // stays selected beside it: a click on it has to open it (owner request, 2026-11-03). On the
        // tunnel route nothing, as only a press that ended on an item is a click.
        // ###########################################################################################
        private void OnComponentsListTapped(object? sender, TappedEventArgs e)
        {
            // A click with Ctrl, Shift or Cmd adds to the selection, and several selected are open already: neither is a click
            // on the one picked (owner report, 2026-11-08: "I cannot mark multiple components when I am in "Contribute" tab" -
            // this put the list back to the one clicked after each Ctrl+click).
            if ((e.KeyModifiers & (KeyModifiers.Control | KeyModifiers.Shift | KeyModifiers.Meta)) != 0 || this.SelectedInList().Count > 1)
            {
                return;
            }

            if (e.Source is Avalonia.Visual source &&
                source.FindAncestorOfType<ListBoxItem>(includeSelf: true)?.DataContext is ComponentListItem item)
            {
                this.OpenPickedComponent(item);
            }
        }
    }
}
