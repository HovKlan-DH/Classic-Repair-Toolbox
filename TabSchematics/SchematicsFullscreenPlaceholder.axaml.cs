using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using System.Collections.ObjectModel;
using TabSchematics;

namespace CRT
{
    public partial class SchematicsFullscreenPlaceholder : UserControl
    {
        private ListBox? thisHostedThumbnailList;
        private bool thisSuppressSelectionSync;

        public SchematicsFullscreenPlaceholder()
        {
            this.InitializeComponent();
        }

        // ###########################################################################################
        // Initializes the placeholder with the shared thumbnails source and selection synchronization.
        // ###########################################################################################
        public void Initialize(ObservableCollection<SchematicThumbnail> thumbnails, ListBox? hostedThumbnailList, double ratio)
        {
            this.thisHostedThumbnailList = hostedThumbnailList;

            this.RootGrid.ColumnDefinitions[0].Width = new GridLength(ratio * 100.0, GridUnitType.Star);
            this.RootGrid.ColumnDefinitions[2].Width = new GridLength((1.0 - ratio) * 100.0, GridUnitType.Star);

            this.SchematicsThumbnailListPlaceholder.ItemsSource = thumbnails;

            if (this.thisHostedThumbnailList != null)
            {
                this.SchematicsThumbnailListPlaceholder.SelectedItem = this.thisHostedThumbnailList.SelectedItem;
                this.SchematicsThumbnailListPlaceholder.SelectionChanged += this.OnPlaceholderSelectionChanged;
                this.thisHostedThumbnailList.SelectionChanged += this.OnHostedThumbnailSelectionChanged;
            }
        }

        // ###########################################################################################
        // Pushes placeholder thumbnail selection changes over to the hosted fullscreen list.
        // ###########################################################################################
        private void OnPlaceholderSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisSuppressSelectionSync || this.thisHostedThumbnailList == null)
                return;

            this.thisSuppressSelectionSync = true;
            this.thisHostedThumbnailList.SelectedItem = this.SchematicsThumbnailListPlaceholder.SelectedItem;

            if (this.SchematicsThumbnailListPlaceholder.SelectedItem != null)
            {
                this.thisHostedThumbnailList.ScrollIntoView(this.SchematicsThumbnailListPlaceholder.SelectedItem);
            }

            this.thisSuppressSelectionSync = false;
        }

        // ###########################################################################################
        // Mirrors hosted fullscreen thumbnail selection changes back into the placeholder list.
        // ###########################################################################################
        private void OnHostedThumbnailSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            if (this.thisSuppressSelectionSync || this.thisHostedThumbnailList == null)
                return;

            this.thisSuppressSelectionSync = true;
            this.SchematicsThumbnailListPlaceholder.SelectedItem = this.thisHostedThumbnailList.SelectedItem;

            if (this.thisHostedThumbnailList.SelectedItem != null)
            {
                this.SchematicsThumbnailListPlaceholder.ScrollIntoView(this.thisHostedThumbnailList.SelectedItem);
            }

            this.thisSuppressSelectionSync = false;
        }
    }
}