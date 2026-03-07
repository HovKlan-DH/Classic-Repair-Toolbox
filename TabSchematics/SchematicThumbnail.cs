using Avalonia;
using Avalonia.Media;
using System.ComponentModel;
using System.Runtime.CompilerServices;

namespace TabSchematics
{
    // ###########################################################################################
    // Represents a single schematic thumbnail item for display in the Schematics tab gallery.
    // ###########################################################################################
    public class SchematicThumbnail : INotifyPropertyChanged
    {
        private IImage? _imageSource;
        private double _visualOpacity = 1.0;
        private bool _isMatchForSelection;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Name { get; init; } = string.Empty;
        public string ImageFilePath { get; init; } = string.Empty;
        public IImage? BaseThumbnail { get; init; }
        public PixelSize OriginalPixelSize { get; init; }

        public IImage? ImageSource
        {
            get => this._imageSource;
            set
            {
                if (ReferenceEquals(this._imageSource, value))
                    return;
                this._imageSource = value;
                this.OnPropertyChanged();
            }
        }

        public double VisualOpacity
        {
            get => this._visualOpacity;
            set
            {
                if (this._visualOpacity == value)
                    return;
                this._visualOpacity = value;
                this.OnPropertyChanged();
            }
        }

        public bool IsMatchForSelection
        {
            get => this._isMatchForSelection;
            set
            {
                if (this._isMatchForSelection == value)
                    return;
                this._isMatchForSelection = value;
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