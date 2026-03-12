using Avalonia;
using Avalonia.Media;
using System;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace TabSchematics
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
        private double placeholderHeight = 120.0;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Name { get; init; } = string.Empty;
        public string ImageFilePath { get; init; } = string.Empty;
        public IImage? BaseThumbnail { get; init; }
        public PixelSize OriginalPixelSize { get; init; }

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

        public double PlaceholderHeight
        {
            get => this.placeholderHeight;
            set
            {
                if (Math.Abs(this.placeholderHeight - value) < 0.01)
                    return;
                this.placeholderHeight = value;
                this.OnPropertyChanged();
            }
        }

        private double placeholderWidth = 160.0;

        public double PlaceholderWidth
        {
            get => this.placeholderWidth;
            set
            {
                if (Math.Abs(this.placeholderWidth - value) < 0.01)
                    return;
                this.placeholderWidth = value;
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