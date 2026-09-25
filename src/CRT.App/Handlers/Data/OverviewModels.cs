using System.Collections.Generic;
using System.ComponentModel;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Row and link models behind the Overview tab's component table and its printable exports.
    //
    // These are plain data models - INotifyPropertyChanged comes from System.ComponentModel, not
    // from Avalonia - so they live here with the logic that builds and groups them rather than
    // inside the tab. The Overview AXAML binds to them through the "data" xmlns mapping.
    // ###########################################################################################
    public class OverviewRow : INotifyPropertyChanged
    {
        private bool _isSelectedForPrint = true;

        public event PropertyChangedEventHandler? PropertyChanged;

        public bool IsSelectedForPrint
        {
            get => this._isSelectedForPrint;
            set
            {
                if (this._isSelectedForPrint == value)
                    return;

                this._isSelectedForPrint = value;
                this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.IsSelectedForPrint)));
            }
        }

        public string Component { get; init; } = string.Empty;
        public string Category { get; init; } = string.Empty;
        public string TechnicalName { get; init; } = string.Empty;
        public string FriendlyName { get; init; } = string.Empty;
        public string PartNumber { get; init; } = string.Empty;
        public string ShortDescription { get; init; } = string.Empty;
        public string Notes { get; init; } = string.Empty;
        public List<OverviewLink> Links { get; init; } = new();
    }

    // ###########################################################################################
    // One aggregated bill-of-materials line: identical components collapsed into a single row.
    // ###########################################################################################
    public class OverviewQuantityGroup
    {
        public string Type { get; init; } = string.Empty;
        public string Components { get; init; } = string.Empty;
        public string TechnicalName { get; init; } = string.Empty;
        public string FriendlyName { get; init; } = string.Empty;
        public int Quantity { get; init; }
    }

    public enum OverviewLinkType
    {
        LocalFile,
        WebLink
    }

    public class OverviewLink
    {
        public string Name { get; }
        public string Target { get; }
        public OverviewLinkType Type { get; }

        public bool IsLocalFile => this.Type == OverviewLinkType.LocalFile;
        public bool IsWebLink => this.Type == OverviewLinkType.WebLink;

        public OverviewLink(string name, string target, OverviewLinkType type)
        {
            this.Name = name;
            this.Target = target;
            this.Type = type;
        }
    }
}
