using Avalonia.Controls;
using Avalonia.Controls.Templates;
using Avalonia.Layout;
using Handlers.DataHandling;
using System;

namespace CRT
{
    // ###########################################################################################
    // The amber "Draft" chip on the Hardware and Board drop-down entries (owner request,
    // 2026-09-24) - the same chip a drafted component row carries in the component list, so a
    // system with local, unpublished work is recognisable before it is even opened. See
    // Main.axaml.cs for the file map of the whole partial class.
    //
    // WHICH ENTRIES are decided by DraftBadgeSet (pure, unit tested), built from the systems the
    // Drafts tab lists. WHEN is ApplyDraftsTabVisibility, which already runs whenever the set of
    // drafts can change - startup, a save that creates a draft, a new system, a discard, a
    // retirement - so the chips follow the Drafts tab in the same pass rather than on a schedule
    // of their own.
    //
    // *** THE DROP-DOWN ITEMS STAY PLAIN STRINGS. *** Roughly twenty places read
    // "SelectedItem as string" or assign a name to SelectedItem, and the board-load cascade hangs
    // off SelectionChanged. Turning the items into view models to carry a flag would touch all of
    // them for a purely visual addition. Instead each drop-down gets a code-built ItemTemplate that
    // asks the current DraftBadgeSet about the name it is drawing.
    //
    // *** A CHANGED SET IS APPLIED BY REPLACING THE TEMPLATE, not the items. *** Re-assigning
    // ItemsSource would fire SelectionChanged and reload the board; a new ItemTemplate instance
    // makes the drop-down re-template its entries and the closed box's selected item alike, with
    // no selection change at all. MainDraftBadgeTests pins exactly that: a chip appearing on an
    // ALREADY RENDERED drop-down.
    // ###########################################################################################
    public partial class Main
    {
        private DraftBadgeSet thisDraftBadges = DraftBadgeSet.Empty;

        // ###########################################################################################
        // Rebuilds the chip set from the Drafts tab's current list and re-templates both drop-downs.
        // Called from ApplyDraftsTabVisibility, AFTER TabDrafts.RefreshDrafts has produced that list.
        // ###########################################################################################
        internal void ApplyDraftBadges()
        {
            this.thisDraftBadges = DraftBadgeSet.From(this.TabDrafts.DraftedSystems);

            this.HardwareComboBox.ItemTemplate = Main.BuildNameWithDraftChipTemplate(
                name => this.thisDraftBadges.HardwareHasDraft(name),
                "One or more boards of this hardware have a local, unpublished draft");

            // The board question is always asked about a board OF the selected hardware - a board
            // name alone is not unique. Read at render time, so switching hardware (which gives this
            // drop-down new items, and so new renders) always asks about the right one.
            this.BoardComboBox.ItemTemplate = Main.BuildNameWithDraftChipTemplate(
                name => this.thisDraftBadges.BoardHasDraft(this.HardwareComboBox.SelectedItem as string, name),
                "This board has a local, unpublished draft");
        }

        // ###########################################################################################
        // One drop-down entry: the name, then the chip when hasDraft says so. The chip's look is
        // App.axaml's Border.DraftChip, shared with the component list, so only its STRUCTURE is
        // built here.
        // ###########################################################################################
        private static FuncDataTemplate<string> BuildNameWithDraftChipTemplate(Func<string, bool> hasDraft, string chipTooltip) =>
            new((name, _) =>
            {
                var panel = new StackPanel
                {
                    Orientation = Orientation.Horizontal,
                    Spacing = 6,
                };

                panel.Children.Add(new TextBlock
                {
                    Text = name,
                    VerticalAlignment = VerticalAlignment.Center,
                });

                if (hasDraft(name))
                {
                    var chip = new Border
                    {
                        Child = new TextBlock { Text = "Draft" },
                    };

                    chip.Classes.Add("DraftChip");
                    ToolTip.SetTip(chip, chipTooltip);

                    panel.Children.Add(chip);
                }

                return panel;
            });
    }
}
