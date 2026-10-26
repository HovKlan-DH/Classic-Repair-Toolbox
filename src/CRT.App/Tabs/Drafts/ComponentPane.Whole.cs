using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // THE COMPONENT AS A WHOLE (CRT 4.0 step 2, slice 3c, 2026-10-26 - the plan: "[Show on board]
    // [Preview popup] [Undo changes] ... The 2,200-line component editor is replaced, not adapted").
    //
    // Owns:
    //   - A NEW COMPONENT: "Add component" opens the pane on a short form - the board label, the region
    //     on a board with regional components, the category. "Add component" there puts it into the
    //     draft as ONE step (ComponentPaneEdits.WithNewComponent, placed in its category) and the pane
    //     shows it like any other. A label the board has in any region is refused with the reason: a
    //     label names one component (ContributionPackaging.ValidateNewComponent). No category is fine -
    //     unfinished work is - and said under the box.
    //   - DELETE: the component off the board in the draft - every row of it, every region
    //     (ComponentBoardWriter.ApplyComponentDelete, as the editor window did) - as ONE step. The pane
    //     then says it is gone and offers "Bring [U8] back" while that step is the latest.
    //   - "SHOW ON BOARD" and "PREVIEW POPUP": asked of the host (ShowOnBoardRequested,
    //     PreviewRequested), what the boxes hold handed to the draft first.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // True while the pane shows the new-component form.
        private bool thisIsNew;

        // Asked to show the component on the schematics - its label named.
        public event EventHandler<string>? ShowOnBoardRequested;

        // Asked to open the component's popup as the draft has it - its label named.
        public event EventHandler<string>? PreviewRequested;

        internal bool IsNewComponentForm => this.thisIsNew;

        private void InitializeWhole()
        {
            this.NewRegionBox.ItemsSource = ComponentPaneWording.RegionChoices;
            this.NewRegionBox.SelectedIndex = 0;
        }

        // ###########################################################################################
        // Opens the form for a new component of the draft held in `session`, in `region` (the one CRT
        // shows - a new component in every region is shown there too). `published` and its file as for
        // Show. What the pane held for another component goes in first.
        // ###########################################################################################
        internal void ShowNew(DraftSession session, string? region, BoardData? published, string publishedBoardFile = "")
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

            this.thisLabel = string.Empty;
            this.thisRegion = region?.Trim() ?? string.Empty;
            this.thisPublished = published;
            this.thisPublishedBoardFile = publishedBoardFile ?? string.Empty;
            this.thisIsNew = true;
            this.thisBoardPart = null;
            this.CloseCopyPanel();

            this.thisShowing = true;

            try
            {
                BoardData board = session.Content.Board;

                this.NewLabelBox.Text = string.Empty;
                this.NewCategoryBox.Text = string.Empty;
                this.NewCategoryBox.ItemsSource = ComponentListBuilder.BuildDistinctCategories(board)
                    .OrderBy(category => category, StringComparer.OrdinalIgnoreCase)
                    .ToList();
                this.NewRegionBox.SelectedIndex = 0;
                this.NewRegionPanel.IsVisible = ComponentListBuilder.HasExplicitRegionComponents(board);
            }
            finally
            {
                this.thisShowing = false;
            }

            this.ShowNewComponentForm();
            Dispatcher.UIThread.Post(() => this.NewLabelBox.Focus(), DispatcherPriority.Background);
        }

        // The form on screen, and whether what it holds can be added.
        private void ShowNewComponentForm()
        {
            this.TitleText.Text = ComponentPaneWording.NewComponentTitle;
            this.RegionText.IsVisible = false;
            this.ChangedMark.IsVisible = false;
            this.ActionsPanel.IsVisible = false;
            this.GonePanel.IsVisible = false;
            this.SectionsScrollViewer.IsVisible = false;
            this.BoardPartPanel.IsVisible = false;
            this.NewComponentPanel.IsVisible = true;

            this.ShowNewComponentState();
        }

        private ContributionPackaging.NewComponentProblem NewComponentProblem() =>
            this.thisSession is { } session
                ? ContributionPackaging.ValidateNewComponent(
                    this.NewLabelBox.Text,
                    this.NewCategoryBox.Text,
                    session.Content.Board.Components.Select(row => row.BoardLabel))
                : ContributionPackaging.NewComponentProblem.BoardLabelMissing;

        // A label taken says so; no category is said and still added; no label yet says nothing.
        private void ShowNewComponentState()
        {
            ContributionPackaging.NewComponentProblem problem = this.NewComponentProblem();
            string? labelProblem = ComponentPaneWording.NewLabelProblem(problem, this.NewLabelBox.Text);

            this.NewLabelProblemText.Text = labelProblem ?? string.Empty;
            this.NewLabelProblemText.IsVisible = labelProblem is not null;
            this.NewCategoryNoteText.Text = ComponentPaneWording.NoCategoryYet;
            this.NewCategoryNoteText.IsVisible = problem == ContributionPackaging.NewComponentProblem.CategoryMissing;
            this.AddNewComponentButton.IsEnabled = ComponentPane.CanAdd(problem);
        }

        private static bool CanAdd(ContributionPackaging.NewComponentProblem problem) =>
            problem is ContributionPackaging.NewComponentProblem.None or ContributionPackaging.NewComponentProblem.CategoryMissing;

        // The form's boxes typed in - judged again (the generic handler leaves them to this).
        private void OnNewComponentTextChanged()
        {
            if (this.thisIsNew && !this.thisShowing)
            {
                this.ShowNewComponentState();
            }
        }

        private void OnAddNewComponentClick(object? sender, RoutedEventArgs e) => this.AddNewComponent();

        // ###########################################################################################
        // The form's component into the draft, as one step - then shown like any other, the friendly
        // name's box focused. False when nothing was added (no label, a label taken).
        // ###########################################################################################
        internal bool AddNewComponent()
        {
            if (!this.thisIsNew || this.thisSession is not { } session || !ComponentPane.CanAdd(this.NewComponentProblem()))
            {
                return false;
            }

            string label = ComponentPaneEdits.Normalised(ComponentField.FriendlyName, this.NewLabelBox.Text);
            string region = this.NewRegionPanel.IsVisible ? ComponentPaneWording.RegionOfChoice(this.NewRegionBox.SelectedItem as string) : string.Empty;
            string category = this.NewCategoryBox.Text ?? string.Empty;

            bool added = session.Apply(
                DraftStepWording.ComponentAdded(label),
                content => ComponentPaneEdits.WithNewComponent(content.Board, label, region, category) is BoardData board && !ReferenceEquals(board, content.Board)
                    ? content with { Board = board }
                    : null,
                this.thisOrigin);

            if (!added)
            {
                return false;
            }

            this.thisIsNew = false;
            this.thisLabel = label;

            // A component of one region is shown in its own.
            if (region.Length > 0)
            {
                this.thisRegion = region;
            }

            this.ShowComponent();
            this.ContentChanged?.Invoke(this, EventArgs.Empty);
            Dispatcher.UIThread.Post(() => this.FriendlyNameBox.Focus(), DispatcherPriority.Background);

            return true;
        }

        // Enter in the form's label or category adds what it holds.
        private bool AddOnEnter(KeyEventArgs e)
        {
            if (!this.thisIsNew || e.Key != Key.Enter)
            {
                return false;
            }

            e.Handled = this.AddNewComponent();
            return true;
        }

        // ---- the buttons above the sections ----------------------------------------------------

        // What the buttons offer for the component shown: "Show on board" only for one marked on a
        // picture, "Delete component" named.
        private void ShowActions(BoardData board)
        {
            bool marked = ComponentPaneEdits.IsMarkedOnAPicture(board, this.thisLabel);

            this.ActionsPanel.IsVisible = true;
            this.ShowOnBoardButton.IsEnabled = marked;
            ToolTip.SetTip(this.ShowOnBoardButton, ComponentPaneWording.ShowOnBoardTip(this.thisLabel, marked));
            this.DeleteComponentButton.Content = "Delete component";
            ToolTip.SetTip(this.DeleteComponentButton, ComponentPaneWording.DeleteTip(this.thisLabel));
        }

        // Gone from the draft: "Bring [U8] back" while the step that deleted it is the latest.
        private void ShowGoneActions()
        {
            this.ActionsPanel.IsVisible = false;

            bool deletedLast = this.thisSession is { } session &&
                string.Equals(session.UndoDescription, DraftStepWording.ComponentDeleted(this.thisLabel), StringComparison.Ordinal);

            this.BringBackButton.IsVisible = deletedLast;
            this.BringBackButton.Content = ComponentPaneWording.BringBack(this.thisLabel);
        }

        private void OnDeleteComponentClick(object? sender, RoutedEventArgs e) => this.DeleteComponent();

        // ###########################################################################################
        // The component off the board in the draft - every row, every region - as one step. False when
        // the draft has no such component.
        // ###########################################################################################
        internal bool DeleteComponent()
        {
            if (this.thisSession is not { } session || this.thisLabel.Length == 0)
            {
                return false;
            }

            this.CommitHeld();

            string label = this.thisLabel;

            bool deleted = session.Apply(
                DraftStepWording.ComponentDeleted(label),
                content => ComponentPaneEdits.RowShown(content.Board, label, region: null) is null
                    ? null
                    : content with { Board = ComponentBoardWriter.ApplyComponentDelete(content.Board, label) },
                this.thisOrigin);

            if (deleted)
            {
                this.ShowComponent();
                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }

            return deleted;
        }

        // The deletion undone - only while it is the latest step, so nothing else is taken back.
        private void OnBringBackClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisSession is { } session &&
                string.Equals(session.UndoDescription, DraftStepWording.ComponentDeleted(this.thisLabel), StringComparison.Ordinal))
            {
                session.Undo();
                this.ContentChanged?.Invoke(this, EventArgs.Empty);
            }
        }

        private void OnShowOnBoardClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisLabel.Length > 0)
            {
                this.CommitHeld();
                this.ShowOnBoardRequested?.Invoke(this, this.thisLabel);
            }
        }

        private void OnPreviewPopupClick(object? sender, RoutedEventArgs e)
        {
            if (this.thisLabel.Length > 0)
            {
                this.CommitHeld();
                this.PreviewRequested?.Invoke(this, this.thisLabel);
            }
        }
    }
}
