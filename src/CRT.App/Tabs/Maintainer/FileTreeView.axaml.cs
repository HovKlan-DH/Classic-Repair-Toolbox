using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Avalonia;
using Avalonia.Collections;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.MaintainerHandling;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // Draws a system's file tree (owner request, 2026-09-28) - see the markup. What each file
    // becomes and which rows show are Handlers/FileTree; this keeps the rows on screen in step with
    // the folders opened and closed, and hands a file to be read or opened to whoever hosts it
    // (BetaView, FileTreeWindow) through Files.
    //
    //   FileTreeView.axaml.cs     - the rows, opening and closing folders, opening a file
    //   FileTreeView.FilePreview.cs - the hover card on a file
    //
    // *** OPENING OR CLOSING A FOLDER CHANGES ONLY ITS OWN ROWS. *** The list is updated in place -
    // the rows that differ are replaced, the rest stay - so it does not jump back to the top every
    // time a folder is clicked halfway down a board.
    // ###########################################################################################
    public partial class FileTreeView : UserControl
    {
        private readonly AvaloniaList<FileTreeRowView> thisRows = [];
        private readonly HashSet<string> thisExpanded = new(StringComparer.Ordinal);

        private FileTreeNode thisRoot = FileTree.Build([]);

        public FileTreeView()
        {
            this.InitializeComponent();

            if (this.FindControl<ListBox>("RowsList") is ListBox list)
                list.ItemsSource = this.thisRows;

            this.WireFilePreview();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // ###########################################################################################
        // Where the files' bytes come from - reading one for the hover card, and fetching and
        // opening one. Null (the default) opens nothing and shows no card.
        // ###########################################################################################
        public IFileTreeFiles? Files { get; set; }

        // Off where the host already says the counts (BetaView's own line).
        public bool ShowSummary
        {
            get => this.FindControl<TextBlock>("SummaryText")?.IsVisible ?? false;
            set
            {
                if (this.FindControl<TextBlock>("SummaryText") is TextBlock summary)
                    summary.IsVisible = value;
            }
        }

        // ###########################################################################################
        // Shows one system's files. Folders on the way to a change start open
        // (FileTree.DefaultExpanded).
        // ###########################################################################################
        public void Show(IReadOnlyList<SystemFileEntry>? files)
        {
            this.HideFilePreview();

            this.thisRoot = FileTree.Build(files);

            this.thisExpanded.Clear();
            this.thisExpanded.UnionWith(FileTree.DefaultExpanded(this.thisRoot));

            if (this.FindControl<TextBlock>("SummaryText") is TextBlock summary)
                summary.Text = files is null ? string.Empty : FileTreeWording.Summary(this.thisRoot);

            this.ShowMessage(null, isError: false);
            this.Refresh(replaceAll: true);
        }

        // Nothing on screen - a system deselected, or a sign-out.
        public void Clear() => this.Show(null);

        public void ShowMessage(string? message, bool isError) =>
            WindowMessage.Show(this.FindControl<TextBlock>("MessageText"), message, isError);

        // The rows on screen, for tests.
        internal IReadOnlyList<FileTreeRowView> RowsForTests => this.thisRows;

        // Opens or closes a folder the way a click does, for tests.
        internal void ToggleForTests(string folderPath)
        {
            if (this.thisRows.FirstOrDefault(row => row.Node.IsFolder && row.Node.Path == folderPath) is FileTreeRowView row)
                this.Toggle(row);
        }

        // ###########################################################################################
        // Ticking the box opens every folder with a change in it again, so the changes are what is
        // on screen; unticking keeps the folders as they are, around the rest of the system.
        // ###########################################################################################
        private void OnOnlyChangedChanged(object? sender, RoutedEventArgs e)
        {
            if (this.OnlyChanged)
                this.thisExpanded.UnionWith(FileTree.DefaultExpanded(this.thisRoot));

            this.Refresh(replaceAll: true);
        }

        private bool OnlyChanged => this.FindControl<CheckBox>("OnlyChangedCheckBox")?.IsChecked == true;

        private void OnExpandAllClick(object? sender, RoutedEventArgs e)
        {
            this.thisExpanded.UnionWith(FileTree.AllFolders(this.thisRoot));
            this.Refresh(replaceAll: true);
        }

        private void OnCollapseAllClick(object? sender, RoutedEventArgs e)
        {
            this.thisExpanded.Clear();
            this.Refresh(replaceAll: true);
        }

        // ###########################################################################################
        // Brings the rows on screen in step with the tree. After a folder is opened or closed only
        // the rows between the unchanged start and the unchanged end are replaced - one change to
        // the list, which keeps its scroll position.
        // ###########################################################################################
        private void Refresh(bool replaceAll)
        {
            this.HideFilePreview();

            IReadOnlyList<FileTreeRowView> rows = FileTree
                .Rows(this.thisRoot, this.thisExpanded, this.OnlyChanged)
                .Select(row => new FileTreeRowView(row))
                .ToList();

            if (replaceAll)
            {
                this.thisRows.Clear();
                this.thisRows.AddRange(rows);
                return;
            }

            int start = 0;

            while (start < this.thisRows.Count && start < rows.Count && this.thisRows[start].SameAs(rows[start]))
                start++;

            int end = 0;

            while (end < this.thisRows.Count - start && end < rows.Count - start &&
                   this.thisRows[this.thisRows.Count - 1 - end].SameAs(rows[rows.Count - 1 - end]))
            {
                end++;
            }

            this.thisRows.RemoveRange(start, this.thisRows.Count - start - end);
            this.thisRows.InsertRange(start, rows.Skip(start).Take(rows.Count - start - end));
        }

        private void Toggle(FileTreeRowView row)
        {
            if (!row.Node.IsFolder)
                return;

            if (!this.thisExpanded.Remove(row.Node.Path))
                this.thisExpanded.Add(row.Node.Path);

            this.Refresh(replaceAll: false);

            // The folder stays the chosen row, so the keyboard carries on from where it was.
            if (this.FindControl<ListBox>("RowsList") is ListBox list)
                list.SelectedItem = this.thisRows.FirstOrDefault(candidate => candidate.Node == row.Node);
        }

        private static FileTreeRowView? RowFrom(object? source) =>
            (source as StyledElement)?.DataContext as FileTreeRowView;

        private void OnRowTapped(object? sender, TappedEventArgs e)
        {
            if (FileTreeView.RowFrom(e.Source) is FileTreeRowView { Node.IsFolder: true } row)
                this.Toggle(row);
        }

        private async void OnRowDoubleTapped(object? sender, TappedEventArgs e)
        {
            if (FileTreeView.RowFrom(e.Source) is FileTreeRowView { Node.IsFolder: false } row)
                await this.OpenAsync(row.Node.File!);
        }

        private async void OnRowsKeyDown(object? sender, KeyEventArgs e)
        {
            if ((sender as ListBox)?.SelectedItem is not FileTreeRowView row)
                return;

            bool open = row.Node.IsFolder && this.thisExpanded.Contains(row.Node.Path);

            switch (e.Key)
            {
                case Key.Enter when row.Node.IsFolder:
                case Key.Right when row.Node.IsFolder && !open:
                case Key.Left when row.Node.IsFolder && open:
                    this.Toggle(row);
                    e.Handled = true;
                    break;

                case Key.Enter:
                    e.Handled = true;
                    await this.OpenAsync(row.Node.File!);
                    break;
            }
        }

        private async Task OpenAsync(SystemFileEntry file)
        {
            if (this.Files is null)
                return;

            this.HideFilePreview();

            if (file.OpenFrom == SystemFileSource.NotWrittenYet)
            {
                this.ShowMessage(FileTreeWording.Note(file), isError: false);
                return;
            }

            this.ShowMessage(null, isError: false);

            string? problem = await this.Files.OpenAsync(file);
            this.ShowMessage(problem, isError: problem is not null);
        }

        private FileTreeRowView? SelectedRow => this.FindControl<ListBox>("RowsList")?.SelectedItem as FileTreeRowView;

        // "Open" is for a file only.
        private void OnRowMenuOpening(object? sender, System.ComponentModel.CancelEventArgs e)
        {
            if (this.SelectedRow is not FileTreeRowView row)
            {
                e.Cancel = true;
                return;
            }

            this.HideFilePreview();

            if (this.FindControl<MenuItem>("OpenRowItem") is MenuItem open)
                open.IsVisible = !row.Node.IsFolder;
        }

        private async void OnOpenRowClick(object? sender, RoutedEventArgs e)
        {
            if (this.SelectedRow is FileTreeRowView { Node.IsFolder: false } row)
                await this.OpenAsync(row.Node.File!);
        }

        private async void OnCopyPathClick(object? sender, RoutedEventArgs e)
        {
            if (this.SelectedRow is FileTreeRowView row && TopLevel.GetTopLevel(this)?.Clipboard is { } clipboard)
                await clipboard.SetTextAsync(row.Node.Path);
        }
    }

    // ###########################################################################################
    // One row as the template draws it. Everything is worked out once, when the row is made - a row
    // is replaced, never changed, when its folder opens or closes.
    // ###########################################################################################
    public sealed class FileTreeRowView
    {
        // Font Awesome: square-minus and square-plus (the Regular face - an outlined box), file.
        internal const string OpenFolderGlyph = "";
        internal const string ClosedFolderGlyph = "";
        internal const string FileGlyph = "";

        internal FileTreeRowView(FileTreeRow row)
        {
            this.Row = row;

            FileTreeNode node = row.Node;
            SystemFileChange change = node.File?.Change ?? SystemFileChange.Unchanged;

            this.Indent = new Thickness(row.Depth * 18, 0, 0, 0);

            this.IsFolder = node.IsFolder;
            this.Glyph = node.IsFolder ? (row.IsExpanded ? OpenFolderGlyph : ClosedFolderGlyph) : FileGlyph;
            this.Name = node.Name;
            this.Weight = node.IsFolder ? FontWeight.SemiBold : FontWeight.Normal;
            this.Decorations = change == SystemFileChange.Removed ? TextDecorations.Strikethrough : null;
            this.HasPill = !node.IsFolder && change != SystemFileChange.Unchanged;
            this.IsAdded = !node.IsFolder && change == SystemFileChange.Added;
            this.IsChanged = !node.IsFolder && change == SystemFileChange.Changed;
            this.IsRemoved = !node.IsFolder && change == SystemFileChange.Removed;
            this.PillText = FileTreeWording.ChangeWord(change);
            this.Note = node.IsFolder ? FileTreeWording.FolderNote(node) : string.Empty;

            // A file has its hover card instead - except one with nothing to show yet, whose path
            // and the reason are all there is to say.
            this.ToolTip = node.IsFolder
                ? node.Path
                : node.File!.OpenFrom == SystemFileSource.NotWrittenYet ? $"{node.Path}\n{FileTreeWording.Note(node.File)}" : null;
        }

        internal FileTreeRow Row { get; }

        internal FileTreeNode Node => this.Row.Node;

        public Thickness Indent { get; }

        public bool IsFolder { get; }

        public string Glyph { get; }

        public string Name { get; }

        public FontWeight Weight { get; }

        public TextDecorationCollection? Decorations { get; }

        public bool HasPill { get; }

        public string PillText { get; }

        // Which of the table's colours the pill takes - see the styles in the markup.
        public bool IsAdded { get; }

        public bool IsChanged { get; }

        public bool IsRemoved { get; }

        public string Note { get; }

        public string? ToolTip { get; }

        // The same line of the tree drawn the same way - an opened folder is a different row.
        internal bool SameAs(FileTreeRowView other) =>
            ReferenceEquals(this.Node, other.Node) && this.Row.IsExpanded == other.Row.IsExpanded;
    }
}
