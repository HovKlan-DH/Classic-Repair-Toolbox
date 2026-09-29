using System;
using System.Threading.Tasks;
using Avalonia.Controls;
using Avalonia.Markup.Xaml;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // The window around one submission's file tree (owner request, 2026-09-28) - see the markup.
    // The main window fills it and keeps one open at a time: "Files..." on another submission shows
    // that one in the same window.
    // ###########################################################################################
    public partial class FileTreeWindow : Window
    {
        public FileTreeWindow()
        {
            this.InitializeComponent();
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        internal FileTreeView TreeView => this.FindControl<FileTreeView>("Tree")!;

        // ###########################################################################################
        // Shows `files` for the submission named in the heading. `readFiles` reads and opens them
        // (FileTreeFiles) - built against this window, so its waits are shown here; null opens none.
        // ###########################################################################################
        public void ShowFiles(
            string systemId,
            string submissionName,
            System.Collections.Generic.IReadOnlyList<SystemFileEntry> files,
            Func<FileTreeWindow, Handlers.MaintainerHandling.IFileTreeFiles>? readFiles)
        {
            this.Title = $"Files - {systemId}";

            if (this.FindControl<TextBlock>("HeadingText") is TextBlock heading)
                heading.Text = systemId.Replace("/", " / ", StringComparison.Ordinal);

            if (this.FindControl<TextBlock>("SubheadingText") is TextBlock subheading)
                subheading.Text = $"The BETA data after approving {submissionName}, compared with BETA now.";

            FileTreeView tree = this.TreeView;
            tree.Files = readFiles?.Invoke(this);
            tree.Show(files);
        }
    }
}
