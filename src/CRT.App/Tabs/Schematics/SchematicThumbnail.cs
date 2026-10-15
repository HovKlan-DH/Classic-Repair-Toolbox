using Avalonia;
using Avalonia.Media;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace Tabs.TabSchematics
{
    // ###########################################################################################
    // Represents a single schematic thumbnail item for display in the Schematics tab gallery.
    // ###########################################################################################
    public class SchematicThumbnail : INotifyPropertyChanged
    {
        private IImage? imageSource;
        private double visualOpacity = 1.0;
        private bool isMatchForSelection;
        private bool isDropPlaceholder;
        private IReadOnlyList<ThumbnailWorklogPillsOverlay.Pill> worklogPills = Array.Empty<ThumbnailWorklogPillsOverlay.Pill>();

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Name { get; init; } = string.Empty;
        public string ImageFilePath { get; init; } = string.Empty;
        public IImage? BaseThumbnail { get; init; }
        public PixelSize OriginalPixelSize { get; init; }

        // ###########################################################################################
        // State pills for the "Show worklogs" overlay - see ThumbnailWorklogPillsOverlay. Populated
        // and cleared by TabSchematics.Worklog.cs's RefreshWorklogEntriesListOverlay so visibility
        // follows the same "Show worklogs" checkbox as the full-size overlay.
        // ###########################################################################################
        public IReadOnlyList<ThumbnailWorklogPillsOverlay.Pill> WorklogPills
        {
            get => this.worklogPills;
            set
            {
                this.worklogPills = value ?? Array.Empty<ThumbnailWorklogPillsOverlay.Pill>();
                this.OnPropertyChanged();
            }
        }

        public IImage? ImageSource
        {
            get => this.imageSource;
            set
            {
                if (ReferenceEquals(this.imageSource, value))
                    return;
                this.imageSource = value;
                this.OnPropertyChanged();
            }
        }

        // ###########################################################################################
        // The selected components' highlights for this thumbnail, as a transparent image drawn OVER
        // ImageSource - and how strongly it shows (code review, 2026-10-13).
        //
        // The highlights used to be baked into ImageSource, so blinking ("Blink selected", twice a
        // second) rendered every highlighted thumbnail again on every tick. Now the highlights are
        // rendered once per selection (TabSchematics.ApplyHighlightVisuals) and the blink only sets
        // HighlightLayerOpacity: 1 shows them, 0 hides them - what the old bitmaps looked like at
        // each phase. Every thumbnail template stacks this Image on the base one, the same size.
        //
        // TabSchematics owns the bitmap: it disposes the one replaced, and Main's board load the
        // last one. BaseThumbnail belongs to Main's thumbnail cache and is never disposed here.
        // ###########################################################################################
        private IImage? highlightImage;
        private double highlightLayerOpacity = 1.0;

        public IImage? HighlightImage
        {
            get => this.highlightImage;
            set
            {
                if (ReferenceEquals(this.highlightImage, value))
                    return;
                this.highlightImage = value;
                this.OnPropertyChanged();
            }
        }

        public double HighlightLayerOpacity
        {
            get => this.highlightLayerOpacity;
            set
            {
                if (this.highlightLayerOpacity == value)
                    return;
                this.highlightLayerOpacity = value;
                this.OnPropertyChanged();
            }
        }

        public double VisualOpacity
        {
            get => this.visualOpacity;
            set
            {
                if (this.visualOpacity == value)
                    return;
                this.visualOpacity = value;
                this.OnPropertyChanged();
            }
        }

        public bool IsMatchForSelection
        {
            get => this.isMatchForSelection;
            set
            {
                if (this.isMatchForSelection == value)
                    return;
                this.isMatchForSelection = value;
                this.OnPropertyChanged();
            }
        }

        public bool IsDropPlaceholder
        {
            get => this.isDropPlaceholder;
            set
            {
                if (this.isDropPlaceholder == value)
                    return;
                this.isDropPlaceholder = value;
                this.OnPropertyChanged();
            }
        }

        // ###########################################################################################
        // Raises PropertyChanged for the given property name.
        // ###########################################################################################
        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
            => this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
    }
}