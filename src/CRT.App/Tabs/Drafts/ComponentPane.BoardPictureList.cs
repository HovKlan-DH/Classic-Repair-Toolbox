using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE BOARD'S PICTURES, ADDED, NAMED, MATCHED, ORDERED AND REMOVED ON THEIR CARDS (CRT 4.0 step 3 -
    // the plan: "Pictures: today's Schematic images content, moved inline"; owner, 2026-10-28: "do
    // what is best for having the best UX for the user. It should be as easy as possible"). The
    // "Schematic images" window is gone - its rules came along, unchanged:
    //
    //   - "ADD PICTURES" (or pictures dropped on the tab, TabDrafts.Drop.cs): each copied into the draft
    //     under a name of the board's own, at its top (ComponentPaneFileNames.ForBoardPicture), named
    //     after its file and never a name a picture has (BoardPictureEdits.NamesFor) - ONE step for
    //     all, the draft saved at once. Never a hidden file (ContributionPackaging.IsAddablePictureFile).
    //   - THE NAME BOX renames on Enter or leaving it, Escape puts the name back. A blank name or one
    //     another picture has is refused with the reason; a rename carries the picture's highlights and
    //     KiCad calibration to the new name in the same step (SchematicRename.Plan, ApplyWithFollow).
    //   - "SHOWS KICAD VIEW" picks the picture's CAD name from the board's views - read by the host off
    //     the UI thread (UseKiCadViews); a view another picture has is left out of the other lists
    //     (SchematicCadNameChoices, owner request 2026-10-10). Every list changed IN PLACE.
    //   - UP AND DOWN move the picture one place - the board's own order (BoardPictureEdits.Moved).
    //   - REMOVE asks first when components are marked on the picture or KiCad is aligned on it - both
    //     go with it (SchematicRemoval, RemoveSchematicWindow); Undo brings all of it back.
    //
    // *** A RENAME OR A VIEW PICKED CHANGES ITS CARD IN PLACE. *** Typing a name and then pressing
    // another card's button commits the name as the box loses focus - on the PRESS - and cards made
    // again then would throw that click away (the trap the drafts list met, 2026-10-10). Only adding,
    // moving and removing - each its own button's doing, its click done - show the cards again.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // The board's KiCad view names - null until the host has read them for the part (UseKiCadViews).
        private IReadOnlyList<string>? thisCadViewNames;

        // ###########################################################################################
        // The host reads the board's KiCad views for the Pictures part - asked when the part opens and
        // again whenever the draft's KiCad files change while it is open (FollowKiCadFiles; code
        // review, 2026-10-28). The ask is numbered: a read answers only the ask it was started for, so
        // one from before an import ending last never puts the old views back.
        // ###########################################################################################
        internal event EventHandler<int>? KiCadViewsWanted;

        private int thisKiCadViewsAsk;

        internal int KiCadViewsAsk => this.thisKiCadViewsAsk;

        private void AskForKiCadViews()
        {
            this.thisKiCadFilesAskedFor = this.thisSession is { } session ? ComponentPane.KiCadFilesOf(session) : string.Empty;
            this.thisKiCadViewsAsk++;
            this.KiCadViewsWanted?.Invoke(this, this.thisKiCadViewsAsk);
        }

        // Stands in for the picker of the pictures to add - null, cancelled.
        internal Func<Task<IReadOnlyList<string>>>? PickPicturesOverrideForTests { get; set; }

        // Stands in for the question before a picture with marks or an alignment is removed.
        internal Func<string, SchematicRemovalLoss, Task<bool>>? ConfirmPictureRemovalOverrideForTests { get; set; }

        // ###########################################################################################
        // The board's KiCad view names, once the host has read them - every card's drop-down follows.
        // `ask`: the KiCadViewsAsk the read was started for - an older one is never taken; null takes
        // them whatever was asked. Only while the Pictures part is shown.
        // ###########################################################################################
        internal void UseKiCadViews(IReadOnlyList<string> views, int? ask = null)
        {
            if (this.thisBoardPart != PaneBoardPart.Pictures || (ask is int asked && asked != this.thisKiCadViewsAsk))
            {
                return;
            }

            this.thisCadViewNames = views ?? [];
            this.ShowCadNameChoices();
        }

        // Each card's views, judged against every other card's - one picture's view is no other's choice.
        private void ShowCadNameChoices()
        {
            List<(string SchematicName, string CadName)> pictures = this.thisBoardPictures
                .Select(row => (row.Name, row.CadName))
                .ToList();

            for (int index = 0; index < this.thisBoardPictures.Count; index++)
            {
                this.thisBoardPictures[index].ApplyCadNameChoices(SchematicCadNameChoices.For(pictures, index, this.thisCadViewNames ?? []));
            }
        }

        // Up and down offered where they lead somewhere.
        private void ShowPictureOrder()
        {
            for (int index = 0; index < this.thisBoardPictures.Count; index++)
            {
                this.thisBoardPictures[index].CanMoveUp = index > 0;
                this.thisBoardPictures[index].CanMoveDown = index < this.thisBoardPictures.Count - 1;
            }
        }

        // Which row of the draft a card is: its name, its image file, its place among cards sharing both.
        private SchematicRowRef RefOf(ComponentPaneBoardPictureRow row) =>
            SchematicRows.RefOf(
                this.thisBoardPictures.Select(other => (other.Name, other.File)).ToList(),
                this.thisBoardPictures.IndexOf(row));

        // The cards shown again from the draft as it is now - after adding, moving or removing.
        private void ShowPicturesAgain()
        {
            if (this.thisSession is { } session)
            {
                this.ShowBoardPictures(session.Content.Board);
            }
        }

        // ---- adding -----------------------------------------------------------------------------

        private async void OnAddPicturesClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                this.CommitHeld();

                IReadOnlyList<string> picked = await this.PickPicturesAsync();

                if (picked.Count > 0)
                {
                    await this.AddPicturesAsync(picked);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Adding pictures of the board in the component pane failed - [{ex}]");
            }
        }

        private async Task<IReadOnlyList<string>> PickPicturesAsync()
        {
            if (this.PickPicturesOverrideForTests is { } pick)
            {
                return await pick();
            }

            if (TopLevel.GetTopLevel(this)?.StorageProvider is not { } storage)
            {
                return [];
            }

            IReadOnlyList<IStorageFile> files = await storage.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = "Pick pictures of the board",
                AllowMultiple = true,
                FileTypeFilter = [ComponentPane.PictureFileType],
            });

            return files.Select(file => file.TryGetLocalPath()).OfType<string>().ToList();
        }

        // ###########################################################################################
        // Pictures into the board - picked, or dropped on the tab: each copied in under a name of the
        // board's own and named after its file, ONE step for all, the draft saved at once. False when
        // nothing was added - said under the heading.
        // ###########################################################################################
        internal async Task<bool> AddPicturesAsync(IReadOnlyList<string> paths)
        {
            if (this.thisSession is not { } session || this.thisBoardPart != PaneBoardPart.Pictures)
            {
                return false;
            }

            this.thisTypingTimer.Stop();
            this.CommitHeld();

            List<string> accepted = paths.Where(ContributionPackaging.IsAddablePictureFile).ToList();

            if (accepted.Count == 0)
            {
                this.ShowPicturesNote(ComponentPaneWording.NoPicturesToAdd);
                return false;
            }

            List<string> taken = this.TakenNames();
            var files = new List<(string Source, string Name)>();

            foreach (string source in accepted)
            {
                string name = ComponentPaneFileNames.ForBoardPicture(source, taken);
                taken.Add(name);
                files.Add((source, name));
            }

            List<string> names = BoardPictureEdits.NamesFor(accepted, session.Content.Board.Schematics.Select(row => row.SchematicName));
            List<(string Name, string File)> pictures = names.Zip(files, (name, file) => (name, file.Name)).ToList();

            return await this.AddToDraftAsync(files, copied =>
            {
                // Only the pictures copied get a row - one locked by another program, or gone since it
                // was picked, is said instead (code review, 2026-10-28: its row named a file the draft
                // lacked).
                var landed = new HashSet<string>(copied, StringComparer.OrdinalIgnoreCase);
                List<(string Name, string File)> added = pictures.Where(picture => landed.Contains(picture.File)).ToList();

                this.ApplyBoardPart(DraftStepWording.SchematicsAdded(added.Count), board => BoardPictureEdits.WithAdded(board, added));
                this.ShowPicturesAgain();
                this.ShowPicturesNote(ComponentPaneWording.PicturesAdded(added.Count, notCopied: pictures.Count - added.Count));
            });
        }

        // ---- the name ---------------------------------------------------------------------------

        private void OnPictureNameKeyDown(object? sender, KeyEventArgs e)
        {
            if (sender is not Control { DataContext: ComponentPaneBoardPictureRow row })
            {
                return;
            }

            if (e.Key == Key.Enter)
            {
                this.CommitPictureName(row);
                e.Handled = true;
            }
            else if (e.Key == Key.Escape && row.HasNameEdit)
            {
                row.ForgetNameEdit();
                e.Handled = true;
            }
        }

        private void OnPictureNameLostFocus(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { DataContext: ComponentPaneBoardPictureRow row })
            {
                this.CommitPictureName(row);
            }
        }

        // ###########################################################################################
        // Whether the Pictures part must stay open because a picture's name in it cannot be taken -
        // asked by the pane's Close and by the host before another part, a component or the table
        // replaces it (code review, 2026-10-28: those threw the typed name and its reason away). The
        // reason is said under the part's heading.
        // ###########################################################################################
        internal bool KeepsRefusedPictureName() =>
            this.thisBoardPart == PaneBoardPart.Pictures && !this.CommitPictureNames();

        // Every name typed and not saved yet - what the pane holds handed over (CommitBoardPart). False
        // when one of them was refused.
        private bool CommitPictureNames()
        {
            bool allTaken = true;

            foreach (ComponentPaneBoardPictureRow row in this.thisBoardPictures.Where(row => row.HasNameEdit).ToList())
            {
                allTaken &= this.CommitPictureName(row);
            }

            return allTaken;
        }

        // ###########################################################################################
        // What the card's name box holds, as the picture's name - ONE step, its highlights and KiCad
        // calibration carried to the new name (SchematicRename.Plan, decided on the draft as it is).
        // True when renamed or nothing was to be saved.
        //
        // *** A NAME REFUSED STAYS IN ITS BOX, with the reason (code review, 2026-10-28). *** A blank
        // name or one another picture has: the typed name was put back to the old one, and closing the
        // pane threw it away with the reason no one saw. Now it waits there to be fixed - Escape puts
        // the old name back - and the pane's Close stays open on it (OnCloseClick), as the
        // "Schematic images" window did. A card the draft no longer has is shown again as the draft is.
        // ###########################################################################################
        internal bool CommitPictureName(ComponentPaneBoardPictureRow row)
        {
            if (this.thisSession is not { } session || !this.thisBoardPictures.Contains(row) || !row.HasNameEdit)
            {
                return true;
            }

            string oldName = row.Name;
            string wanted = row.NameDraft?.Trim() ?? string.Empty;

            if (string.Equals(wanted, oldName.Trim(), StringComparison.Ordinal))
            {
                row.ForgetNameEdit();
                return true;
            }

            string? problem = SchematicRename.Problem(wanted, this.thisBoardPictures.Where(other => !ReferenceEquals(other, row)).Select(other => other.Name));

            if (problem is not null)
            {
                this.ShowPicturesNote(problem);
                return false;
            }

            SchematicRowRef target = this.RefOf(row);
            SchematicRenameOutcome? decided = null;

            bool renamed = session.ApplyWithFollow(
                DraftStepWording.SchematicRenamed(oldName, wanted),
                content =>
                {
                    SchematicRenamePlan plan = SchematicRename.Plan(content.Board, target, wanted);
                    decided = plan.Outcome;
                    return plan.Board is null ? (null, null) : (content with { Board = plan.Board }, plan.Follow);
                },
                this.thisOrigin);

            if (!renamed)
            {
                switch (decided)
                {
                    case SchematicRenameOutcome.NameTaken:
                        this.ShowPicturesNote(SchematicRename.NameTaken(wanted));
                        return false;
                    case SchematicRenameOutcome.Blank:
                        this.ShowPicturesNote(SchematicRename.NeedsAName);
                        return false;
                    case SchematicRenameOutcome.Unchanged:
                        row.ForgetNameEdit();
                        return true;
                    default:
                        // The draft moved under the card - shown as it now is, nothing said wrongly.
                        this.ShowPicturesAgain();
                        return true;
                }
            }

            row.ForgetNameEdit();
            row.Name = wanted;
            this.ShowCadNameChoices();
            this.ShowBoardPartChanged();
            this.ContentChanged?.Invoke(this, EventArgs.Empty);
            return true;
        }

        // ---- the KiCad view ---------------------------------------------------------------------

        // Queued AFTER the binding that reported it has finished - never under a ComboBox still setting
        // its selection (the window's lesson).
        private void OnCadNameChosen(ComponentPaneBoardPictureRow row, string cadName) =>
            Dispatcher.UIThread.Post(() => this.CommitCadName(row, cadName));

        // ###########################################################################################
        // The view a card's drop-down picked, as the picture's CAD name - ONE step - then every card's
        // list follows (a view one picture has is no other's choice) and Align KiCad with it.
        // ###########################################################################################
        internal void CommitCadName(ComponentPaneBoardPictureRow row, string cadName)
        {
            if (this.thisSession is not { } session || !this.thisBoardPictures.Contains(row))
            {
                return;
            }

            string wanted = cadName?.Trim() ?? string.Empty;
            SchematicRowRef target = this.RefOf(row);
            bool found = false;

            this.ApplyBoardPart(
                DraftStepWording.CadName(row.Name),
                board =>
                {
                    int index = SchematicRows.IndexOf(board.Schematics, target);
                    found = index >= 0;

                    return found && !string.Equals(board.Schematics[index].CadName?.Trim(), wanted, StringComparison.OrdinalIgnoreCase)
                        ? SchematicRows.WithCadName(board, target, wanted).Board
                        : board;
                });

            if (!found)
            {
                // The draft moved under the card - show it as it is.
                this.ShowPicturesAgain();
                return;
            }

            row.CadName = wanted;
            this.ShowCadNameChoices();
        }

        // ---- order ------------------------------------------------------------------------------

        private void OnMovePictureUpClick(object? sender, RoutedEventArgs e) => this.MovePicture(sender, -1);

        private void OnMovePictureDownClick(object? sender, RoutedEventArgs e) => this.MovePicture(sender, +1);

        private void MovePicture(object? sender, int by)
        {
            if (sender is Control { Tag: ComponentPaneBoardPictureRow row })
            {
                this.MovePicture(row, by);
            }
        }

        // One place up (-1) or down (+1), ONE step - what the pane holds handed over first.
        internal void MovePicture(ComponentPaneBoardPictureRow row, int by)
        {
            this.thisTypingTimer.Stop();
            this.CommitHeld();

            if (this.thisSession is null || !this.thisBoardPictures.Contains(row))
            {
                return;
            }

            SchematicRowRef target = this.RefOf(row);
            this.ApplyBoardPart(DraftStepWording.SchematicOrder, board => BoardPictureEdits.Moved(board, target, by));
            this.ShowPicturesAgain();
        }

        // ---- removing ---------------------------------------------------------------------------

        private async void OnRemoveBoardPictureClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Control { Tag: ComponentPaneBoardPictureRow row })
                {
                    await this.RemovePictureAsync(row);
                }
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Removing a picture of the board in the component pane failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The picture taken off the board - ONE step, only the card pressed on (by name AND image file,
        // SchematicRows). Asked first when components are marked on it or KiCad is aligned on it: both
        // go with it, and adding the picture again brings neither back (SchematicRemoval). Undo does.
        // ###########################################################################################
        internal async Task RemovePictureAsync(ComponentPaneBoardPictureRow row)
        {
            this.thisTypingTimer.Stop();
            this.CommitHeld();

            if (this.thisSession is not { } session || !this.thisBoardPictures.Contains(row))
            {
                return;
            }

            string name = row.Name;
            SchematicRowRef target = this.RefOf(row);

            // Another picture keeping the name keeps what hangs on it, too.
            bool nameStays = this.thisBoardPictures.Any(other => !ReferenceEquals(other, row) && SchematicRename.SameName(other.Name, name));

            SchematicRemovalLoss loss = nameStays
                ? SchematicRemovalLoss.Nothing
                : SchematicRemoval.LossOf(
                    name,
                    session.Content.Board.ComponentHighlights,
                    session.Content.Calibrations.Any(calibration => SchematicRename.SameName(calibration.SchematicName, name)));

            if (!loss.IsNothing && !await this.ConfirmPictureRemovalAsync(name, loss))
            {
                return;
            }

            bool removed = false;

            this.ApplyBoardPart(
                DraftStepWording.SchematicRemoved(name),
                board =>
                {
                    if (SchematicRows.Without(board, target) is (BoardData without, true))
                    {
                        removed = true;
                        return without;
                    }

                    return board;
                });

            // Said only when it went - a card the draft no longer had is just shown again.
            this.ShowPicturesAgain();
            this.ShowPicturesNote(removed ? ComponentPaneWording.PictureRemoved(name) : null);
        }

        private async Task<bool> ConfirmPictureRemovalAsync(string name, SchematicRemovalLoss loss)
        {
            if (this.ConfirmPictureRemovalOverrideForTests is { } confirm)
            {
                return await confirm(name, loss);
            }

            if (TopLevel.GetTopLevel(this) is not Window owner)
            {
                return false;
            }

            var dialog = new RemoveSchematicWindow();
            dialog.Initialize(name, loss);

            return await dialog.ShowDialog<bool?>(owner) == true;
        }
    }
}
