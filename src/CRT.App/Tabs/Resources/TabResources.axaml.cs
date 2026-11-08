using Avalonia;
using Avalonia.Controls;
using System;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Windows.Input;

namespace CRT
{
    public partial class TabResources : UserControl
    {
        public ObservableCollection<ResourceItem> LocalFiles { get; set; } = new ObservableCollection<ResourceItem>();
        public ObservableCollection<ResourceItem> WebLinks { get; set; } = new ObservableCollection<ResourceItem>();

        public TabResources()
        {
            this.InitializeComponent();

            this.LocalFilesItemsControl.ItemsSource = this.LocalFiles;
            this.WebLinksItemsControl.ItemsSource = this.WebLinks;
        }

        // ###########################################################################################
        // Populates the resources tab tables with items retrieved from the board data source
        // ###########################################################################################
        public void LoadData(System.Collections.Generic.IEnumerable<ResourceItem> localFiles, System.Collections.Generic.IEnumerable<ResourceItem> webLinks)
        {
            this.LocalFiles.Clear();
            foreach (var file in localFiles)
            {
                this.LocalFiles.Add(file);
            }

            this.WebLinks.Clear();
            foreach (var link in webLinks)
            {
                this.WebLinks.Add(link);
            }
        }
    }

    public class ResourceItem
    {
        public string Category { get; set; }
        public string Name { get; set; }
        public string Target { get; set; }

        // The root a local file was found under (DraftFileResolution.Root) - a draft's file sits in
        // the draft folder, not the data root. Null for a link, and for a file of the data root.
        public string? Root { get; }

        // What the file or link is about, several lines (owner request, 2026-11-06) - shown under its name,
        // its web addresses clickable as in any text a user typed (TextLinkRenderer). Empty for none.
        public string Note { get; }

        public bool HasNote => this.Note.Trim().Length > 0;

        public ICommand OpenCommand { get; }

        public ResourceItem(string category, string name, string target, string? root = null, string? note = null)
        {
            this.Category = category;
            this.Name = name;
            this.Target = target;
            this.Root = root;
            this.Note = note?.Trim() ?? string.Empty;
            this.OpenCommand = new ActionCommand(this.OpenTarget);
        }

        // ###########################################################################################
        // Opens the specified target only after strict validation. URLs are limited to HTTP/HTTPS
        // and local files must remain inside the root they were found under.
        // ###########################################################################################
        private void OpenTarget()
        {
            ExternalTargetLauncher.OpenOrSay(this.Target, this.Root);
        }
    }

    // Standard concise ICommand implementation wrapper
    public class ActionCommand : ICommand
    {
        private readonly Action _action;

        public ActionCommand(Action action)
        {
            this._action = action;
        }

        public event EventHandler? CanExecuteChanged
        {
            add { }
            remove { }
        }

        public bool CanExecute(object? parameter) => true;

        public void Execute(object? parameter) => this._action();
    }

}