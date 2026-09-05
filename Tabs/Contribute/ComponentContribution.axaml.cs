using Avalonia;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.Input;
using Avalonia.Platform.Storage;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.IO;
using System.IO.Compression;
using System.ComponentModel;
using System.Runtime.CompilerServices;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Handlers.DataHandling;

namespace CRT
{
    public sealed class ContributionComponentRow : INotifyPropertyChanged
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public string UuidV4 { get; set; } = string.Empty;

        private string thisBoardLabel = string.Empty;
        public string BoardLabel
        {
            get => this.thisBoardLabel;
            set
            {
                if (this.thisBoardLabel != value)
                {
                    this.thisBoardLabel = value;
                    this.OnPropertyChanged();

                    // Typing into the box answers the complaint, so the mark goes at once rather
                    // than surviving until the next attempt to send.
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        this.HasBoardLabelError = false;
                        this.BoardLabelErrorText = string.Empty;
                    }
                }
            }
        }

        public string FriendlyName { get; set; } = string.Empty;
        public string TechnicalNameOrValue { get; set; } = string.Empty;
        public string PartNumber { get; set; } = string.Empty;

        private string thisCategory = string.Empty;
        public string Category
        {
            get => this.thisCategory;
            set
            {
                if (this.thisCategory != value)
                {
                    this.thisCategory = value;
                    this.OnPropertyChanged();

                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        this.HasCategoryError = false;
                        this.CategoryErrorText = string.Empty;
                    }
                }
            }
        }

        public string Region { get; set; } = string.Empty;
        public string Description { get; set; } = string.Empty;

        // Set by the pre-submit validation of a NEW component, whose board label is the one field
        // that must be filled in and must not already be taken. The board label box turns red and
        // BoardLabelErrorText appears beneath it. Never part of the uploaded payload.
        private bool thisHasBoardLabelError;
        [JsonIgnore]
        public bool HasBoardLabelError
        {
            get => this.thisHasBoardLabelError;
            set
            {
                if (this.thisHasBoardLabelError != value)
                {
                    this.thisHasBoardLabelError = value;
                    this.OnPropertyChanged();
                }
            }
        }

        private string thisBoardLabelErrorText = string.Empty;
        [JsonIgnore]
        public string BoardLabelErrorText
        {
            get => this.thisBoardLabelErrorText;
            set
            {
                if (this.thisBoardLabelErrorText != value)
                {
                    this.thisBoardLabelErrorText = value;
                    this.OnPropertyChanged();
                }
            }
        }

        // The categories this board already uses, offered as you type. Suggestions only - a
        // category the board has never used is still accepted, it just has to be typed in full.
        // Never part of the uploaded payload.
        [JsonIgnore]
        public ObservableCollection<string> AvailableCategories { get; } = new();

        // The same marking for the category, which a new component is equally unusable without:
        // the main window builds its category filter from the categories in the data and skips
        // blank ones, so a component with none is invisible there however complete it otherwise is.
        private bool thisHasCategoryError;
        [JsonIgnore]
        public bool HasCategoryError
        {
            get => this.thisHasCategoryError;
            set
            {
                if (this.thisHasCategoryError != value)
                {
                    this.thisHasCategoryError = value;
                    this.OnPropertyChanged();
                }
            }
        }

        private string thisCategoryErrorText = string.Empty;
        [JsonIgnore]
        public string CategoryErrorText
        {
            get => this.thisCategoryErrorText;
            set
            {
                if (this.thisCategoryErrorText != value)
                {
                    this.thisCategoryErrorText = value;
                    this.OnPropertyChanged();
                }
            }
        }
    }

    public interface IContributionFileRow
    {
        string FileLocation { get; set; }
        string File { get; set; }
        string? OriginalFilePath { get; set; }
        ObservableCollection<string> AvailableFileLocations { get; }
    }

    public sealed class ContributionComponentImageRow : INotifyPropertyChanged, IContributionFileRow
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public string UuidV4 { get; set; } = string.Empty;
        public string BoardLabel { get; set; } = string.Empty;
        public string Region { get; set; } = string.Empty;
        public string Pin { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string ExpectedOscilloscopeReading { get; set; } = string.Empty;
        public string VoltsDiv { get; set; } = string.Empty;
        public string TimeDiv { get; set; } = string.Empty;
        public string TriggerLevelVolts { get; set; } = string.Empty;

        private string thisFileLocation = string.Empty;
        public string FileLocation
        {
            get => this.thisFileLocation;
            set
            {
                if (this.thisFileLocation != value)
                {
                    this.thisFileLocation = value;
                    this.OnPropertyChanged();
                }
            }
        }

        private string thisFile = string.Empty;
        public string File
        {
            get => this.thisFile;
            set
            {
                if (this.thisFile != value)
                {
                    this.thisFile = value;
                    this.OnPropertyChanged();
                }
            }
        }

        [JsonIgnore]
        public ObservableCollection<string> AvailableFileLocations { get; } = new();

        public string Note { get; set; } = string.Empty;

        // Zip entry name of the attached file for this row, so the server can locate it exactly.
        // Empty when the row's file could not be resolved and therefore was not attached.
        public string ZipEntry { get; set; } = string.Empty;

        [JsonIgnore]
        public string? OriginalFilePath { get; set; }

        private Bitmap? thisPreviewImage;
        [JsonIgnore]
        public Bitmap? PreviewImage
        {
            get => this.thisPreviewImage;
            set
            {
                if (this.thisPreviewImage != value)
                {
                    this.thisPreviewImage = value;
                    this.OnPropertyChanged();
                }
            }
        }

        [JsonIgnore]
        public string PreviewStatusText { get; set; } = "No preview available";

        // Set by the pre-submit validation so the row can show where the problem is: the row
        // border turns red and FileErrorText appears inside it. Cleared as soon as the row is
        // given a usable file. Never part of the uploaded payload.
        private bool thisHasFileError;
        [JsonIgnore]
        public bool HasFileError
        {
            get => this.thisHasFileError;
            set
            {
                if (this.thisHasFileError != value)
                {
                    this.thisHasFileError = value;
                    this.OnPropertyChanged();
                }
            }
        }

        private string thisFileErrorText = string.Empty;
        [JsonIgnore]
        public string FileErrorText
        {
            get => this.thisFileErrorText;
            set
            {
                if (this.thisFileErrorText != value)
                {
                    this.thisFileErrorText = value;
                    this.OnPropertyChanged();
                }
            }
        }
    }

    public sealed class ContributionComponentLocalFileRow : INotifyPropertyChanged, IContributionFileRow
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public string UuidV4 { get; set; } = string.Empty;
        public string BoardLabel { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;

        private string thisFileLocation = string.Empty;
        public string FileLocation
        {
            get => this.thisFileLocation;
            set
            {
                if (this.thisFileLocation != value)
                {
                    this.thisFileLocation = value;
                    this.OnPropertyChanged();
                }
            }
        }

        private string thisFile = string.Empty;
        public string File
        {
            get => this.thisFile;
            set
            {
                if (this.thisFile != value)
                {
                    this.thisFile = value;
                    this.OnPropertyChanged();
                }
            }
        }

        [JsonIgnore]
        public ObservableCollection<string> AvailableFileLocations { get; } = new();

        // Zip entry name of the attached file for this row, so the server can locate it exactly.
        // Empty when the row's file could not be resolved and therefore was not attached.
        public string ZipEntry { get; set; } = string.Empty;

        [JsonIgnore]
        public string? OriginalFilePath { get; set; }
    }

    public sealed class ContributionComponentLinkRow
    {
        public string UuidV4 { get; set; } = string.Empty;
        public string BoardLabel { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
    }

    public sealed class ContributionBoardLocalFileRow : INotifyPropertyChanged, IContributionFileRow
    {
        public event PropertyChangedEventHandler? PropertyChanged;

        private void OnPropertyChanged([CallerMemberName] string? propertyName = null)
        {
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(propertyName));
        }

        public string UuidV4 { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;

        private string thisFileLocation = string.Empty;
        public string FileLocation
        {
            get => this.thisFileLocation;
            set
            {
                if (this.thisFileLocation != value)
                {
                    this.thisFileLocation = value;
                    this.OnPropertyChanged();
                }
            }
        }

        private string thisFile = string.Empty;
        public string File
        {
            get => this.thisFile;
            set
            {
                if (this.thisFile != value)
                {
                    this.thisFile = value;
                    this.OnPropertyChanged();
                }
            }
        }

        [JsonIgnore]
        public ObservableCollection<string> AvailableFileLocations { get; } = new();

        // Zip entry name of the attached file for this row, so the server can locate it exactly.
        // Empty when the row's file could not be resolved and therefore was not attached.
        public string ZipEntry { get; set; } = string.Empty;

        [JsonIgnore]
        public string? OriginalFilePath { get; set; }
    }

    public sealed class ContributionBoardLinkRow
    {
        public string UuidV4 { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
    }

    public sealed class ComponentContributionPayload
    {
        // Bumped whenever the payload schema changes, so the server review page can tell which
        // contract a queued submission was produced with. Version 2 added BoardExcelFile,
        // BoardRevisionDate and the per-row ZipEntry pointers.
        public int PayloadFormat { get; set; } = 2;
        public string ApplicationVersion { get; set; } = string.Empty;
        public string HardwareName { get; set; } = string.Empty;
        public string BoardName { get; set; } = string.Empty;
        public string BoardExcelFile { get; set; } = string.Empty;
        public string BoardRevisionDate { get; set; } = string.Empty;
        public string Region { get; set; } = string.Empty;
        public string ComponentBoardLabel { get; set; } = string.Empty;
        public string ComponentDisplayText { get; set; } = string.Empty;
        public string ComponentUuidV4 { get; set; } = string.Empty;
        public string Email { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;
        public DateTimeOffset SubmittedUtc { get; set; }

        // True when this submission asks for the component to be REMOVED rather than changed. The
        // component-scoped sections below are then sent present-but-empty, which the server's
        // buildDiffSections already reads as "remove every server row in this section" - so the
        // flag adds no new merge behaviour. What it adds is that the review page can say what it
        // is looking at, instead of the reviewer inferring a deletion from five empty lists.
        public bool DeleteComponent { get; set; }

        public List<ContributionComponentRow> Components { get; set; } = new();
        public List<ContributionComponentImageRow> ComponentImages { get; set; } = new();
        public List<ComponentHighlightEntry> ComponentHighlights { get; set; } = new();
        public List<ContributionComponentLocalFileRow> ComponentLocalFiles { get; set; } = new();
        public List<ContributionComponentLinkRow> ComponentLinks { get; set; } = new();
        public List<ContributionBoardLocalFileRow> BoardLocalFiles { get; set; } = new();
        public List<ContributionBoardLinkRow> BoardLinks { get; set; } = new();
    }

    public partial class ComponentContributionWindow : Window
    {
        public ObservableCollection<string> AvailableEndFolders { get; } = new();

        private readonly ObservableCollection<ContributionComponentRow> thisComponentRows = new();
        private readonly ObservableCollection<ContributionComponentImageRow> thisComponentImageRows = new();
        private readonly ObservableCollection<ContributionComponentLocalFileRow> thisComponentLocalFileRows = new();
        private readonly ObservableCollection<ContributionComponentLinkRow> thisComponentLinkRows = new();
        private readonly ObservableCollection<ContributionBoardLocalFileRow> thisBoardLocalFileRows = new();
        private readonly ObservableCollection<ContributionBoardLinkRow> thisBoardLinkRows = new();
        private readonly List<ComponentHighlightEntry> thisComponentHighlightRows = new();

        private string thisHardwareName = string.Empty;
        private string thisBoardName = string.Empty;
        private string thisBoardExcelFile = string.Empty;
        private string thisBoardRevisionDate = string.Empty;
        private string thisLocalRegion = string.Empty;
        private string thisBoardLabel = string.Empty;
        private string thisComponentDisplayText = string.Empty;
        private string thisDataRoot = string.Empty;
        private string thisComponentUuidV4 = string.Empty;

        // True when the window was opened on a component that is not in the board data at all. The
        // board label then comes from the contributor rather than from the board, which is what the
        // extra validation guards - see LoadNewComponent and ValidateNewComponentRow.
        private bool thisIsNewComponent;

        // True once the contributor has switched the window into "remove this component" mode. The
        // data sections are then disabled and dimmed rather than hidden - a deletion should be made
        // with the thing being deleted still in front of you - and the payload sends the
        // component-scoped sections empty, which is what the server reads as "remove these rows".
        private bool thisIsDeleteComponent;

        // Every board label already on the board, so a new component cannot reuse one of them.
        private readonly HashSet<string> thisExistingBoardLabels = new(StringComparer.OrdinalIgnoreCase);

        // The categories this board already uses, offered as suggestions on every component row.
        private readonly List<string> thisAvailableCategories = new();

        // Shown in place of the component summary while a new component has no label yet.
        private const string NewComponentTitleText = "New component - not yet part of the board data";

        // How far the data sections fade once delete mode is on. Faint enough to read as inactive,
        // solid enough that the rows can still be read - which is the whole point of dimming them
        // rather than hiding them.
        private const double DimmedSectionOpacity = 0.45;

        private static readonly JsonSerializerOptions thisContributionPayloadJsonOptions = new()
        {
            WriteIndented = true
        };

        public ComponentContributionWindow()
        {
            this.InitializeComponent();

            this.ComponentRowsItemsControl.ItemsSource = this.thisComponentRows;
            this.ComponentImageRowsItemsControl.ItemsSource = this.thisComponentImageRows;
            this.ComponentLocalFileRowsItemsControl.ItemsSource = this.thisComponentLocalFileRows;
            this.ComponentLinkRowsItemsControl.ItemsSource = this.thisComponentLinkRows;
            this.BoardLocalFileRowsItemsControl.ItemsSource = this.thisBoardLocalFileRows;
            this.BoardLinkRowsItemsControl.ItemsSource = this.thisBoardLinkRows;

            this.EmailTextBox.Text = UserSettings.ContactEmail;
            this.Closed += this.OnWindowClosed;

            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Loads the selected component context and all editable rows into the window.
        // ###########################################################################################
        public void LoadComponent(BoardData boardData, string dataRoot, string hardwareName, string boardName, string region, string boardLabel, string boardExcelFile)
        {
            this.ApplyBoardContext(boardData, dataRoot, hardwareName, boardName, region, boardExcelFile);

            this.thisIsNewComponent = false;
            this.thisIsDeleteComponent = false;
            this.thisExistingBoardLabels.Clear();
            this.thisBoardLabel = boardLabel;
            this.thisComponentUuidV4 = string.Empty;

            // Only a component the board actually has can be deleted.
            this.DeleteComponentButton.IsVisible = true;
            this.DeleteComponentHintTextBlock.IsVisible = true;

            var primaryComponent = boardData.Components.FirstOrDefault(c =>
                string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase) &&
                (string.IsNullOrWhiteSpace(c.Region) ||
                 string.Equals(c.Region.Trim(), region, StringComparison.OrdinalIgnoreCase)))
                ?? boardData.Components.FirstOrDefault(c =>
                    string.Equals(c.BoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase));

            this.thisComponentUuidV4 = primaryComponent?.UuidV4?.Trim() ?? string.Empty;
            this.thisComponentDisplayText = this.BuildComponentDisplayText(primaryComponent, boardLabel);

            this.PopulateHeader();
            this.LoadRows(boardData, boardLabel);
        }

        // ###########################################################################################
        // Opens the editor on a component this board does not have yet. Nothing is preloaded for the
        // component itself - the single blank row is where the contributor names it - but the
        // board-wide sections are loaded exactly as for an existing component: those are diffed
        // against the server as a whole, so sending them empty would read as a request to delete
        // every board local file and board link the board has.
        // ###########################################################################################
        public void LoadNewComponent(BoardData boardData, string dataRoot, string hardwareName, string boardName, string region, string boardExcelFile)
        {
            this.ApplyBoardContext(boardData, dataRoot, hardwareName, boardName, region, boardExcelFile);

            this.thisIsNewComponent = true;
            this.thisIsDeleteComponent = false;
            this.thisBoardLabel = string.Empty;
            this.thisComponentUuidV4 = string.Empty;
            this.thisComponentDisplayText = NewComponentTitleText;

            // Nothing to delete - the component does not exist anywhere yet.
            this.DeleteComponentButton.IsVisible = false;
            this.DeleteComponentHintTextBlock.IsVisible = false;

            // Every label on the board, whatever its region: the server resolves a contribution by
            // board label alone, so a label another region's component holds is taken here too.
            this.thisExistingBoardLabels.Clear();
            foreach (var component in boardData.Components)
            {
                string existingLabel = component.BoardLabel?.Trim() ?? string.Empty;
                if (!string.IsNullOrWhiteSpace(existingLabel))
                {
                    this.thisExistingBoardLabels.Add(existingLabel);
                }
            }

            this.PopulateHeader();
            this.LoadRows(boardData, string.Empty);

            var newRow = new ContributionComponentRow
            {
                Region = this.thisLocalRegion
            };

            this.SetAvailableCategories(newRow);
            this.thisComponentRows.Add(newRow);

            // The one section that has to be filled in, so it does not start folded away.
            this.ComponentExpander.IsExpanded = true;
        }

        // ###########################################################################################
        // Applies the board-level context shared by both ways of opening the window.
        // ###########################################################################################
        private void ApplyBoardContext(BoardData boardData, string dataRoot, string hardwareName, string boardName, string region, string boardExcelFile)
        {
            this.thisDataRoot = dataRoot;
            this.thisHardwareName = hardwareName;
            this.thisBoardName = boardName;
            this.thisBoardExcelFile = boardExcelFile?.Trim().Replace('\\', '/') ?? string.Empty;
            this.thisBoardRevisionDate = boardData.RevisionDate?.Trim() ?? string.Empty;
            this.thisLocalRegion = region;

            this.PopulateEndFolders(dataRoot);
            this.PopulateAvailableCategories(boardData);
        }

        // ###########################################################################################
        // Collects the categories the board already uses, so a component row can suggest them while
        // the category is being typed. Matching the existing spelling matters: the main window groups
        // and filters components by this exact string, so "Capacitors" beside "Capacitor" splits one
        // group into two rather than joining the one that is there.
        // ###########################################################################################
        private void PopulateAvailableCategories(BoardData boardData)
        {
            this.thisAvailableCategories.Clear();

            var categories = boardData.Components
                .Select(component => component.Category?.Trim() ?? string.Empty)
                .Where(category => !string.IsNullOrWhiteSpace(category))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(category => category, StringComparer.OrdinalIgnoreCase);

            this.thisAvailableCategories.AddRange(categories);
        }

        // ###########################################################################################
        // Fills one component row's category suggestion list from the board's categories.
        // ###########################################################################################
        private void SetAvailableCategories(ContributionComponentRow row)
        {
            row.AvailableCategories.Clear();

            foreach (var category in this.thisAvailableCategories)
            {
                row.AvailableCategories.Add(category);
            }
        }

        // ###########################################################################################
        // Discovers and populates all end folders within the given data root directory.
        // ###########################################################################################
        private void PopulateEndFolders(string dataRoot)
        {
            this.AvailableEndFolders.Clear();

            if (string.IsNullOrWhiteSpace(dataRoot) || !Directory.Exists(dataRoot))
            {
                return;
            }

            try
            {
                var endFolders = new List<string>();
                this.FindEndFoldersRecursive(dataRoot, dataRoot, endFolders);

                foreach (var folder in endFolders.OrderBy(f => f))
                {
                    this.AvailableEndFolders.Add(folder);
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Failed to populate end folders: {ex.Message}");
            }
        }

        // ###########################################################################################
        // Helper to recursively find all directories containing no sub-directories.
        // ###########################################################################################
        private void FindEndFoldersRecursive(string rootPath, string currentPath, List<string> endFolders)
        {
            try
            {
                var subDirs = Directory.GetDirectories(currentPath);
                if (subDirs.Length == 0)
                {
                    string relativePath = Path.GetRelativePath(rootPath, currentPath);
                    if (!string.IsNullOrWhiteSpace(relativePath) && relativePath != ".")
                    {
                        endFolders.Add(relativePath.Replace('\\', '/'));
                    }
                }
                else
                {
                    foreach (var subDir in subDirs)
                    {
                        this.FindEndFoldersRecursive(rootPath, subDir, endFolders);
                    }
                }
            }
            catch
            {
                // Unreadable directories are skipped safely
            }
        }

        // ###########################################################################################
        // Populates the window header with the selected component and board context.
        // ###########################################################################################
        private void PopulateHeader()
        {
            if (this.thisIsDeleteComponent)
            {
                this.Title = string.IsNullOrWhiteSpace(this.thisBoardLabel)
                    ? "Component contribution - delete component"
                    : $"Component contribution - delete {this.thisBoardLabel}";
            }
            else if (this.thisIsNewComponent)
            {
                this.Title = "Component contribution - new component";
            }
            else
            {
                this.Title = string.IsNullOrWhiteSpace(this.thisBoardLabel)
                    ? "Component contribution"
                    : $"Component contribution - {this.thisBoardLabel}";
            }

            // Every way of opening the window carries the same warning - what is sent is a
            // suggestion for the online data, not an edit of anything on this machine - so only the
            // wording changes with the mode.
            this.ContributionNoticeHeadingTextBlock.Text = this.thisIsDeleteComponent
                ? "You are proposing to DELETE this component"
                : this.thisIsNewComponent
                    ? "You are adding a component this board does not have yet"
                    : "You are modifying an existing component";

            // The delete notice replaces both of the others rather than joining them: in delete
            // mode nothing is being added or modified, so a line about the edits reaching the
            // online source describes a submission that is not being made.
            this.NewComponentNoticeTextBlock.IsVisible = !this.thisIsDeleteComponent && this.thisIsNewComponent;
            this.ExistingComponentNoticeTextBlock.IsVisible = !this.thisIsDeleteComponent && !this.thisIsNewComponent;
            this.DeleteComponentHintTextBlock.IsVisible = !this.thisIsDeleteComponent && !this.thisIsNewComponent;
            this.DeleteComponentNoticeTextBlock.IsVisible = this.thisIsDeleteComponent;

            if (this.thisIsDeleteComponent)
            {
                this.DeleteComponentNoticeTextBlock.Text = ContributionPackaging.BuildDeleteComponentSummary(
                    this.thisBoardLabel,
                    this.thisComponentImageRows.Count,
                    this.thisComponentHighlightRows.Count,
                    this.thisComponentLocalFileRows.Count,
                    this.thisComponentLinkRows.Count);
            }

            this.ComponentTitleTextBlock.Text = this.thisComponentDisplayText;
            this.HardwareContextTextBlock.Text = $"Hardware: {this.thisHardwareName}";
            this.BoardContextTextBlock.Text = $"Board...: {this.thisBoardName}";
            this.RegionContextTextBlock.Text = $"Region..: {this.thisLocalRegion}";
            this.ComponentImagesRegionTextBlock.Text = $"Component images relevant for the {this.thisLocalRegion} region";
        }

        // ###########################################################################################
        // Loads editable row collections from the selected board and component.
        // ###########################################################################################
        private void LoadRows(BoardData boardData, string boardLabel)
        {
            // A new component owns nothing that is already on the board - not even a stray data row
            // carrying a blank board label, which comparing against a blank label would drag in.
            bool BelongsToComponent(string? rowBoardLabel) =>
                !this.thisIsNewComponent &&
                string.Equals(rowBoardLabel, boardLabel, StringComparison.OrdinalIgnoreCase);

            foreach (var row in this.thisComponentImageRows)
            {
                DisposeComponentImagePreview(row);
            }

            this.thisComponentRows.Clear();
            this.thisComponentImageRows.Clear();
            this.thisComponentLocalFileRows.Clear();
            this.thisComponentLinkRows.Clear();
            this.thisBoardLocalFileRows.Clear();
            this.thisBoardLinkRows.Clear();
            this.thisComponentHighlightRows.Clear();

            foreach (var row in boardData.Components.Where(c => BelongsToComponent(c.BoardLabel)))
            {
                var componentRow = new ContributionComponentRow
                {
                    UuidV4 = row.UuidV4,
                    BoardLabel = row.BoardLabel,
                    FriendlyName = row.FriendlyName,
                    TechnicalNameOrValue = row.TechnicalNameOrValue,
                    PartNumber = row.PartNumber,
                    Category = row.Category,
                    Region = row.Region,
                    Description = row.Description
                };

                this.SetAvailableCategories(componentRow);
                this.thisComponentRows.Add(componentRow);
            }

            foreach (var row in boardData.ComponentImages.Where(c =>
                BelongsToComponent(c.BoardLabel) &&
                (string.IsNullOrWhiteSpace(c.Region) ||
                 string.Equals(c.Region.Trim(), this.thisLocalRegion, StringComparison.OrdinalIgnoreCase))))
            {
                string fileLocation = GetExistingFileLocation(row, row.File);

                var imageRow = new ContributionComponentImageRow
                {
                    UuidV4 = row.UuidV4,
                    BoardLabel = row.BoardLabel,
                    Region = row.Region,
                    Pin = row.Pin,
                    Name = row.Name,
                    ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading,
                    VoltsDiv = row.VoltsDiv,
                    TimeDiv = row.TimeDiv,
                    TriggerLevelVolts = row.TriggerLevelVolts,
                    FileLocation = fileLocation,
                    File = Path.GetFileName(row.File ?? string.Empty),
                    OriginalFilePath = row.File,
                    Note = row.Note
                };

                this.SetAvailableFileLocations(imageRow);
                this.thisComponentImageRows.Add(imageRow);
            }

            foreach (var row in boardData.ComponentHighlights.Where(c => BelongsToComponent(c.BoardLabel)))
            {
                this.thisComponentHighlightRows.Add(new ComponentHighlightEntry
                {
                    SchematicName = row.SchematicName,
                    BoardLabel = row.BoardLabel,
                    X = row.X,
                    Y = row.Y,
                    Width = row.Width,
                    Height = row.Height
                });
            }

            foreach (var row in boardData.ComponentLocalFiles.Where(c => BelongsToComponent(c.BoardLabel)))
            {
                string fileLocation = GetExistingFileLocation(row, row.File);

                var localFileRow = new ContributionComponentLocalFileRow
                {
                    UuidV4 = row.UuidV4,
                    BoardLabel = row.BoardLabel,
                    Name = row.Name,
                    FileLocation = fileLocation,
                    File = Path.GetFileName(row.File ?? string.Empty),
                    OriginalFilePath = row.File
                };

                this.SetAvailableFileLocations(localFileRow);
                this.thisComponentLocalFileRows.Add(localFileRow);
            }

            foreach (var row in boardData.ComponentLinks.Where(c => BelongsToComponent(c.BoardLabel)))
            {
                this.thisComponentLinkRows.Add(new ContributionComponentLinkRow
                {
                    UuidV4 = row.UuidV4,
                    BoardLabel = row.BoardLabel,
                    Name = row.Name,
                    Url = row.Url
                });
            }

            foreach (var row in boardData.BoardLocalFiles)
            {
                string fileLocation = GetExistingFileLocation(row, row.File);

                var boardLocalFileRow = new ContributionBoardLocalFileRow
                {
                    UuidV4 = row.UuidV4,
                    Category = row.Category,
                    Name = row.Name,
                    FileLocation = fileLocation,
                    File = Path.GetFileName(row.File ?? string.Empty),
                    OriginalFilePath = row.File
                };

                this.SetAvailableFileLocations(boardLocalFileRow);
                this.thisBoardLocalFileRows.Add(boardLocalFileRow);
            }

            foreach (var row in boardData.BoardLinks)
            {
                this.thisBoardLinkRows.Add(new ContributionBoardLinkRow
                {
                    UuidV4 = row.UuidV4,
                    Category = row.Category,
                    Name = row.Name,
                    Url = row.Url
                });
            }

            this.RefreshAllComponentImagePreviews();
            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Builds a compact display label for the selected component.
        // ###########################################################################################
        private string BuildComponentDisplayText(ComponentEntry? component, string boardLabel)
        {
            if (component == null)
            {
                return boardLabel;
            }

            return BuildComponentDisplayText(
                component.BoardLabel,
                component.FriendlyName,
                component.TechnicalNameOrValue,
                boardLabel);
        }

        // ###########################################################################################
        // Builds the same compact label from loose values, for a component that exists only as an
        // edited row and therefore has no board data entry to read it from.
        // ###########################################################################################
        private static string BuildComponentDisplayText(
            string? boardLabel,
            string? friendlyName,
            string? technicalNameOrValue,
            string fallbackText)
        {
            var parts = new List<string>();
            if (!string.IsNullOrWhiteSpace(boardLabel))
                parts.Add(boardLabel.Trim());
            if (!string.IsNullOrWhiteSpace(friendlyName))
                parts.Add(friendlyName.Trim());
            if (!string.IsNullOrWhiteSpace(technicalNameOrValue))
                parts.Add(technicalNameOrValue.Trim());

            return parts.Count == 0 ? fallbackText : string.Join(" | ", parts);
        }

        // ###########################################################################################
        // The board label this whole contribution belongs to. For an existing component that is the
        // label the window was opened on; for a new one it is whatever was typed into the single
        // component row, which is the only place it exists.
        // ###########################################################################################
        private string ResolveEffectiveBoardLabel()
        {
            if (!this.thisIsNewComponent)
            {
                return this.thisBoardLabel?.Trim() ?? string.Empty;
            }

            return this.thisComponentRows.FirstOrDefault()?.BoardLabel?.Trim() ?? string.Empty;
        }

        // ###########################################################################################
        // The component summary carried by the payload and the notification email. A new component
        // is described by what has just been typed rather than by the header text, which was written
        // before it had a name.
        // ###########################################################################################
        private string ResolveComponentDisplayText()
        {
            var row = this.thisIsNewComponent ? this.thisComponentRows.FirstOrDefault() : null;
            if (row == null)
            {
                return this.thisComponentDisplayText;
            }

            return BuildComponentDisplayText(
                row.BoardLabel,
                row.FriendlyName,
                row.TechnicalNameOrValue,
                this.thisComponentDisplayText);
        }

        // ###########################################################################################
        // The board label a component-scoped row belongs to. Rows added in this window are stamped
        // with it as they are created, but a new component has no label at that point - so a row
        // still blank at send time inherits the label finally entered.
        // ###########################################################################################
        private static string ResolveRowBoardLabel(string? rowBoardLabel, string effectiveBoardLabel)
        {
            string trimmed = rowBoardLabel?.Trim() ?? string.Empty;

            return string.IsNullOrWhiteSpace(trimmed) ? effectiveBoardLabel : trimmed;
        }

        // ###########################################################################################
        // Switches the window into - or back out of - "delete this component" mode.
        //
        // It is a toggle rather than a one-way door because the only thing standing between a
        // mistaken click and a submitted deletion is the mandatory comment; being able to back out
        // is what makes the button safe to press in order to see what it would remove.
        // ###########################################################################################
        private void OnToggleDeleteComponentClick(object? sender, RoutedEventArgs e)
        {
            this.thisIsDeleteComponent = !this.thisIsDeleteComponent;
            this.ApplyDeleteComponentMode();
        }

        // ###########################################################################################
        // Applies the current mode to the whole window.
        //
        // The data sections are DISABLED AND DIMMED rather than hidden. Hiding them would leave the
        // contributor confirming a deletion with no way to look at what is about to go, and the
        // counts in the section headers are exactly what makes the decision an informed one; a
        // disabled expander can still be opened and read. Dimming is what says the rows are no
        // longer part of what gets sent, since a deletion submits none of them.
        // ###########################################################################################
        private void ApplyDeleteComponentMode()
        {
            bool deleting = this.thisIsDeleteComponent;

            foreach (var section in this.GetDataSectionControls())
            {
                section.IsEnabled = !deleting;
                section.Opacity = deleting ? DimmedSectionOpacity : 1.0;
            }

            this.DeleteComponentButton.Content = deleting ? "Cancel deletion" : "Delete this component";
            this.SubmitButton.Content = deleting ? "Send deletion request" : "Send contribution update";

            this.PopulateHeader();

            if (deleting)
            {
                // The section holding the button is where the mode is left again, so it must not be
                // folded away behind the contributor while they are in it.
                //
                // Nothing else moves. Scrolling down to the comment box would take the contributor
                // away from the very rows the notice has just told them are about to be deleted,
                // at the moment they most need to look at them - and it moves the button they may
                // have clicked by mistake off screen. The comment is checked on send, which is
                // where being sent to it belongs (RevealMandatoryComment).
                this.ComponentExpander.IsExpanded = true;
            }
        }

        // ###########################################################################################
        // The six containers holding contributable data. Note that DeleteComponentButton is
        // deliberately NOT inside ComponentSectionFields: Avalonia gives no way for a child to
        // re-enable itself under a disabled parent, so a button inside one of these could not be
        // clicked again to leave delete mode.
        // ###########################################################################################
        private IEnumerable<Control> GetDataSectionControls()
        {
            yield return this.ComponentSectionFields;
            yield return this.ComponentImagesSection;
            yield return this.ComponentLocalFilesSection;
            yield return this.ComponentLinksSection;
            yield return this.BoardLocalFilesSection;
            yield return this.BoardLinksSection;
        }

        // ###########################################################################################
        // Removes an editable row from the Components section.
        // ###########################################################################################
        private void OnRemoveComponentRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionComponentRow row })
            {
                this.thisComponentRows.Remove(row);
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Component images section.
        // ###########################################################################################
        private void OnAddComponentImageRowClick(object? sender, RoutedEventArgs e)
        {
            var row = new ContributionComponentImageRow
            {
                BoardLabel = this.thisBoardLabel,
                Region = this.thisLocalRegion
            };

            this.SetAvailableFileLocations(row);

            InsertRowAtTop(this.thisComponentImageRows, row);
            this.RefreshComponentImagePreview(row);
            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Component images section.
        // ###########################################################################################
        private void OnRemoveComponentImageRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionComponentImageRow row })
            {
                DisposeComponentImagePreview(row);
                this.thisComponentImageRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Component local files section.
        // ###########################################################################################
        private void OnAddComponentLocalFileRowClick(object? sender, RoutedEventArgs e)
        {
            var row = new ContributionComponentLocalFileRow
            {
                BoardLabel = this.thisBoardLabel
            };

            this.SetAvailableFileLocations(row);

            InsertRowAtTop(this.thisComponentLocalFileRows, row);
            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Component local files section.
        // ###########################################################################################
        private void OnRemoveComponentLocalFileRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionComponentLocalFileRow row })
            {
                this.thisComponentLocalFileRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Component links section.
        // ###########################################################################################
        private void OnAddComponentLinkRowClick(object? sender, RoutedEventArgs e)
        {
            InsertRowAtTop(this.thisComponentLinkRows, new ContributionComponentLinkRow
            {
                BoardLabel = this.thisBoardLabel
            });

            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Component links section.
        // ###########################################################################################
        private void OnRemoveComponentLinkRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionComponentLinkRow row })
            {
                this.thisComponentLinkRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Board local files section.
        // ###########################################################################################
        private void OnAddBoardLocalFileRowClick(object? sender, RoutedEventArgs e)
        {
            var row = new ContributionBoardLocalFileRow();

            this.SetAvailableFileLocations(row);

            InsertRowAtTop(this.thisBoardLocalFileRows, row);
            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Board local files section.
        // ###########################################################################################
        private void OnRemoveBoardLocalFileRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionBoardLocalFileRow row })
            {
                this.thisBoardLocalFileRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Adds a new editable row to the Board links section.
        // ###########################################################################################
        private void OnAddBoardLinkRowClick(object? sender, RoutedEventArgs e)
        {
            InsertRowAtTop(this.thisBoardLinkRows, new ContributionBoardLinkRow());
            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Removes an editable row from the Board links section.
        // ###########################################################################################
        private void OnRemoveBoardLinkRowClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: ContributionBoardLinkRow row })
            {
                this.thisBoardLinkRows.Remove(row);
                this.UpdateSectionCounters();
            }
        }

        // ###########################################################################################
        // Closes the window without sending anything.
        // ###########################################################################################
        private void OnCancelClick(object? sender, RoutedEventArgs e)
        {
            this.Close();
        }

        // ###########################################################################################
        // Validates the edited payload and submits it to the contribution backend.
        // ###########################################################################################
        private async void OnSubmitClick(object? sender, RoutedEventArgs e)
        {
            string email = this.EmailTextBox.Text?.Trim() ?? string.Empty;
            string comment = this.MandatoryCommentTextBox.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(email))
            {
                this.ShowStatus("Please provide your email address before sending", true);
                return;
            }

            if (!Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                this.ShowStatus("Please enter a valid email address", true);
                return;
            }

            if (string.IsNullOrWhiteSpace(comment))
            {
                this.RevealMandatoryComment();
                this.ShowStatus("Please provide a mandatory change comment before sending", true);
                return;
            }

            if (this.ShouldValidateContributedRows())
            {
                var newComponentProblem = this.ValidateNewComponentRow();
                if (newComponentProblem != null)
                {
                    this.RevealComponentRow(newComponentProblem.Value.Row);
                    this.ShowStatus(newComponentProblem.Value.Message, true);
                    return;
                }

                var componentImageProblem = this.ValidateComponentImageRows();
                if (componentImageProblem != null)
                {
                    this.RevealComponentImageRow(componentImageProblem.Value.Row);
                    this.ShowStatus(componentImageProblem.Value.Message, true);
                    return;
                }
            }

            UserSettings.ContactEmail = email;
            this.SubmitButton.IsEnabled = false;

            bool submissionAccepted = false;

            try
            {
                IProgress<string> progress = new Progress<string>(statusMessage =>
                {
                    this.ShowStatus(statusMessage, false);
                });

                var result = await this.ProcessAndSendContributionAsync(email, comment, progress);

                if (result.Success)
                {
                    submissionAccepted = true;
                    this.ShowStatus(this.BuildSubmissionSuccessText(), false);
                }
                else
                {
                    Logger.Warning($"Component contribution submission failed. HTTP {result.StatusCode}. Server responded with: {result.ResponseBody}");

                    if (ContributionPackaging.TryParseOutdatedVersionResponse(result.ResponseBody, out string minimumVersion))
                    {
                        // The server names the MINIMUM version it accepts, not necessarily the
                        // newest release - any version at or above it is fine.
                        string updateTargetText = string.IsNullOrWhiteSpace(minimumVersion)
                            ? "a newer version"
                            : $"version [{minimumVersion}] or newer";

                        this.ShowStatus($"This application version [{AppConfig.AppDisplayVersionString}] is too old to contribute data - please update to {updateTargetText}", true);
                    }
                    else if (result.StatusCode == 404)
                    {
                        this.ShowStatus("Failed to send contribution: Server endpoint not found (HTTP 404)", true);
                    }
                    else
                    {
                        this.ShowStatus($"Failed to send contribution (HTTP {result.StatusCode}) - please check the logfile for details", true);
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Exception while sending component contribution: {ex}");
                this.ShowStatus("Network or system error while sending contribution - please try again later...", true);
            }
            finally
            {
                this.ApplySubmissionOutcome(submissionAccepted);
            }
        }

        // ###########################################################################################
        // Whether the edited rows are checked before sending. A deletion submits none of them, so
        // neither row check applies to it. The component image check in particular MUST be skipped:
        // a component whose image file has gone missing, or was never in a format the application
        // can display, is a prime candidate for deletion - and refusing the deletion over the very
        // row being deleted would make broken data impossible to remove.
        // ###########################################################################################
        private bool ShouldValidateContributedRows()
        {
            return !this.thisIsDeleteComponent;
        }

        // ###########################################################################################
        // Marks the mandatory comment box and brings it on screen. It is the last thing in the
        // scrolling area, so on a contribution of any size it sits well below the fold - naming the
        // problem in the status line alone leaves the user looking at the wrong part of the window.
        // ###########################################################################################
        private void RevealMandatoryComment()
        {
            SetErrorMark(this.MandatoryCommentTextBox, true);
            this.MandatoryCommentTextBox.BringIntoView();
            this.MandatoryCommentTextBox.Focus();
        }

        // ###########################################################################################
        // Drops the mark the moment the user starts writing, so the box stops claiming to be empty
        // while it is being filled in.
        // ###########################################################################################
        private void OnMandatoryCommentTextChanged(object? sender, TextChangedEventArgs e)
        {
            if (!string.IsNullOrWhiteSpace(this.MandatoryCommentTextBox.Text))
            {
                SetErrorMark(this.MandatoryCommentTextBox, false);
            }
        }

        // ###########################################################################################
        // Adds or removes the styling class that turns a required box red. Guarded because Classes
        // is a plain list - adding twice would leave a duplicate that one removal cannot undo.
        // ###########################################################################################
        private static void SetErrorMark(Control control, bool hasError)
        {
            bool isMarked = control.Classes.Contains("HasError");

            if (hasError && !isMarked)
            {
                control.Classes.Add("HasError");
            }
            else if (!hasError && isMarked)
            {
                control.Classes.Remove("HasError");
            }
        }

        // ###########################################################################################
        // Settles what the Send button does after an attempt. A contribution the server accepted
        // cannot be sent from this window again: the same suggestion would be queued for review a
        // second time, and for a new component that means the component itself proposed twice. The
        // button says why it is disabled, since a greyed-out button on its own only looks broken.
        //
        // A failed attempt is the opposite case - the button comes straight back, because the whole
        // point of a failure message is that the user can fix it and try again.
        // ###########################################################################################
        private void ApplySubmissionOutcome(bool submissionAccepted)
        {
            this.SubmitButton.IsEnabled = !submissionAccepted;

            ToolTip.SetTip(
                this.SubmitButton,
                submissionAccepted
                    ? "Already sent - close this window and open it again to send another contribution"
                    : null);
        }

        // ###########################################################################################
        // The line shown once the server has accepted the contribution. A new component is worth its
        // own wording: a contribution is a suggestion for the online data and changes nothing on this
        // machine, so the component the contributor has just described stays absent from every list
        // here until the change has been reviewed and synced back down. Saying so is what stops it
        // reading as a submission that quietly did nothing.
        // ###########################################################################################
        private string BuildSubmissionSuccessText()
        {
            const string SuccessText = "Contribution submitted successfully - thank you :-)";

            if (this.thisIsDeleteComponent)
            {
                string deletedLabel = this.ResolveEffectiveBoardLabel();
                string deletedText = string.IsNullOrWhiteSpace(deletedLabel)
                    ? "The component"
                    : $"The component [{deletedLabel}]";

                return $"{SuccessText} {deletedText} stays in your local data until the deletion has been reviewed and accepted, after which you can download the updated data.";
            }

            if (!this.thisIsNewComponent)
            {
                return SuccessText;
            }

            string boardLabel = this.ResolveEffectiveBoardLabel();
            string componentText = string.IsNullOrWhiteSpace(boardLabel)
                ? "The new component"
                : $"The new component [{boardLabel}]";

            return $"{SuccessText} {componentText} will get added to the online source once the contribution has been reviewed and accepted.";
        }

        // ###########################################################################################
        // Checks the two fields a NEW component cannot be submitted without, marks the one at fault
        // and returns its row together with the message for the status line - or null when there is
        // nothing to complain about. An existing component is never checked here: it was resolved
        // from the board data and already carries what the board agrees with.
        // ###########################################################################################
        private (ContributionComponentRow Row, string Message)? ValidateNewComponentRow()
        {
            if (!this.thisIsNewComponent)
            {
                return null;
            }

            var row = this.thisComponentRows.FirstOrDefault();
            if (row == null)
            {
                return null;
            }

            var problem = ContributionPackaging.ValidateNewComponent(
                row.BoardLabel,
                row.Category,
                this.thisExistingBoardLabels);

            row.BoardLabelErrorText = problem switch
            {
                ContributionPackaging.NewComponentProblem.BoardLabelMissing => "A board label is required",
                ContributionPackaging.NewComponentProblem.BoardLabelAlreadyExists => "This board label is already taken",
                _ => string.Empty
            };

            row.CategoryErrorText = problem == ContributionPackaging.NewComponentProblem.CategoryMissing
                ? "A category is required"
                : string.Empty;

            row.HasBoardLabelError = !string.IsNullOrEmpty(row.BoardLabelErrorText);
            row.HasCategoryError = !string.IsNullOrEmpty(row.CategoryErrorText);

            string message = problem switch
            {
                ContributionPackaging.NewComponentProblem.BoardLabelMissing =>
                    "The new component needs a board label - it is what names the component on the board",
                ContributionPackaging.NewComponentProblem.BoardLabelAlreadyExists =>
                    $"This board already has a component labelled [{row.BoardLabel?.Trim()}] - close this window and pick it from the component list to change it",
                ContributionPackaging.NewComponentProblem.CategoryMissing =>
                    "The new component needs a category - without one it never appears in the component list, whatever else it carries",
                _ => string.Empty
            };

            return string.IsNullOrEmpty(message) ? null : (row, message);
        }

        // ###########################################################################################
        // Brings a component row on screen so its red mark is actually seen - the same reasoning as
        // RevealComponentImageRow below: the section can be collapsed, and the row can sit below the
        // fold, so the scroll is posted once the expander has laid its content out.
        // ###########################################################################################
        private void RevealComponentRow(ContributionComponentRow row)
        {
            this.ComponentExpander.IsExpanded = true;

            Dispatcher.UIThread.Post(() =>
            {
                int index = this.thisComponentRows.IndexOf(row);
                if (index < 0)
                {
                    return;
                }

                if (this.ComponentRowsItemsControl.ContainerFromIndex(index) is Control container)
                {
                    container.BringIntoView();
                }
            }, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // Marks every component image row that cannot be submitted and returns the first of them
        // together with the message for the status line, or null when all rows are fine. Marking
        // happens on every row, not just the first, so one pass shows the reviewer every problem;
        // rows that are fine get their mark cleared here too.
        // ###########################################################################################
        private (ContributionComponentImageRow Row, string Message)? ValidateComponentImageRows()
        {
            (ContributionComponentImageRow Row, string Message)? firstProblem = null;

            for (int index = 0; index < this.thisComponentImageRows.Count; index++)
            {
                var row = this.thisComponentImageRows[index];
                var problem = ContributionPackaging.ValidateComponentImageFile(GetStoredFilePath(row));

                row.FileErrorText = problem switch
                {
                    ContributionPackaging.ComponentImageFileProblem.NoFileSelected => "No image file selected",
                    ContributionPackaging.ComponentImageFileProblem.NotDisplayable => "Not an image the application can display",
                    _ => string.Empty
                };

                row.HasFileError = problem != ContributionPackaging.ComponentImageFileProblem.None;

                if (row.HasFileError && firstProblem == null)
                {
                    string rowLabel = $"Component image #{index + 1}";

                    string message = problem == ContributionPackaging.ComponentImageFileProblem.NoFileSelected
                        ? $"{rowLabel} has no file selected"
                        : $"{rowLabel} is not a format the application can display - use one of: " +
                          string.Join(", ", ContributionPackaging.DisplayableImageExtensions);

                    firstProblem = (row, message);
                }
            }

            return firstProblem;
        }

        // ###########################################################################################
        // Brings a component image row on screen so its red mark is actually seen: the section is
        // collapsed by default, and the row can sit far below the fold. The scroll is posted because
        // the row's container does not exist until the expander has laid its content out.
        // ###########################################################################################
        private void RevealComponentImageRow(ContributionComponentImageRow row)
        {
            this.ComponentImagesExpander.IsExpanded = true;

            Dispatcher.UIThread.Post(() =>
            {
                int index = this.thisComponentImageRows.IndexOf(row);
                if (index < 0)
                {
                    return;
                }

                if (this.ComponentImageRowsItemsControl.ContainerFromIndex(index) is Control container)
                {
                    container.BringIntoView();
                }
            }, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // Builds the contribution payload, zips it together with any referenced files, and posts it.
        // ###########################################################################################
        private async Task<(bool Success, int StatusCode, string ResponseBody)> ProcessAndSendContributionAsync(string email, string comment, IProgress<string> progress)
        {
            progress.Report("Preparing contribution payload...");

            var payload = this.BuildPayload(email, comment);

            // A deletion attaches no files. It also MUST NOT go through AssignZipEntriesToPayload:
            // that walks the source row collections and the payload lists in lockstep by index, and
            // a deletion has emptied the payload side while the source rows are all still there.
            var attachments = this.thisIsDeleteComponent
                ? Array.Empty<ContributionAttachment>()
                : this.AssignZipEntriesToPayload(payload);
            string payloadJson = JsonSerializer.Serialize(payload, thisContributionPayloadJsonOptions);

            using var memoryStream = new MemoryStream();

            using (var archive = new ZipArchive(memoryStream, ZipArchiveMode.Create, true))
            {
                AddTextEntryToZip(archive, "ComponentContribution.json", payloadJson);

                for (int i = 0; i < attachments.Count; i++)
                {
                    progress.Report($"Packaging referenced files... {i + 1}/{attachments.Count}");
                    AddFileToZipSafe(archive, attachments[i].SourcePath, attachments[i].ZipEntryName);
                }
            }

            memoryStream.Position = 0;

            using var httpClient = new HttpClient
            {
                Timeout = AppConfig.UploadTimeout
            };

            using var formContent = new MultipartFormDataContent();
            formContent.Add(new StringContent(email), "email");
            formContent.Add(new StringContent(this.BuildContributionFeedbackText(comment)), "feedback");
            formContent.Add(new StringContent(AppConfig.AppDisplayVersionString), "version");

            var fileContent = new ByteArrayContent(memoryStream.ToArray());
            fileContent.Headers.ContentType = new MediaTypeHeaderValue("application/zip");
            formContent.Add(fileContent, "attachmentFile", "ComponentContributionPayload.zip");

            using var progressContent = new ProgressableStreamContent(formContent, percent =>
                progress.Report($"Sending to server... {percent}%"));

            httpClient.DefaultRequestHeaders.UserAgent.ParseAdd(AppConfig.AppShortName + " " + AppConfig.AppDisplayVersionString);

            var response = await httpClient.PostAsync(AppConfig.ContributionUploadUrl, progressContent);
            string responseBody = await response.Content.ReadAsStringAsync();
            bool isSuccess = response.IsSuccessStatusCode &&
                             responseBody.Trim().StartsWith("Success", StringComparison.OrdinalIgnoreCase);

            return (isSuccess, (int)response.StatusCode, responseBody);
        }

        // ###########################################################################################
        // Builds a structured payload representing the current edited state of the window.
        //
        // A deletion is expressed in the same shape rather than as a separate message: the five
        // component-scoped sections go out present-but-EMPTY, which the server already reads as
        // "remove every server row in this section", while the two board-wide sections go out
        // unchanged so nothing outside the component is proposed for removal. The identity fields
        // keep naming the component being removed - they must not be re-derived from the emptied
        // rows, which is why ResolveEffectiveBoardLabel reads the loaded label for an existing
        // component rather than the row.
        // ###########################################################################################
        private ComponentContributionPayload BuildPayload(string email, string comment)
        {
            string effectiveBoardLabel = this.ResolveEffectiveBoardLabel();
            bool deleting = this.thisIsDeleteComponent;

            return new ComponentContributionPayload
            {
                DeleteComponent = deleting,
                ApplicationVersion = AppConfig.AppDisplayVersionString,
                HardwareName = this.thisHardwareName,
                BoardName = this.thisBoardName,
                BoardExcelFile = this.thisBoardExcelFile,
                BoardRevisionDate = this.thisBoardRevisionDate,
                Region = this.thisLocalRegion,
                ComponentBoardLabel = effectiveBoardLabel,
                ComponentDisplayText = this.ResolveComponentDisplayText(),
                ComponentUuidV4 = this.thisComponentUuidV4?.Trim() ?? string.Empty,
                Email = email,
                Comment = comment,
                SubmittedUtc = DateTimeOffset.UtcNow,

                Components = deleting ? new() : this.thisComponentRows.Select(row => new ContributionComponentRow
                {
                    UuidV4 = row.UuidV4?.Trim() ?? string.Empty,
                    BoardLabel = row.BoardLabel?.Trim() ?? string.Empty,
                    FriendlyName = row.FriendlyName?.Trim() ?? string.Empty,
                    TechnicalNameOrValue = row.TechnicalNameOrValue?.Trim() ?? string.Empty,
                    PartNumber = row.PartNumber?.Trim() ?? string.Empty,
                    Category = row.Category?.Trim() ?? string.Empty,
                    Region = row.Region?.Trim() ?? string.Empty,
                    Description = row.Description?.Trim() ?? string.Empty
                }).ToList(),

                ComponentImages = deleting ? new() : this.thisComponentImageRows.Select(row => new ContributionComponentImageRow
                {
                    UuidV4 = row.UuidV4?.Trim() ?? string.Empty,
                    BoardLabel = ResolveRowBoardLabel(row.BoardLabel, effectiveBoardLabel),
                    Region = row.Region?.Trim() ?? string.Empty,
                    Pin = row.Pin?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading?.Trim() ?? string.Empty,
                    VoltsDiv = row.VoltsDiv?.Trim() ?? string.Empty,
                    TimeDiv = row.TimeDiv?.Trim() ?? string.Empty,
                    TriggerLevelVolts = row.TriggerLevelVolts?.Trim() ?? string.Empty,
                    FileLocation = row.FileLocation?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty,
                    Note = row.Note?.Trim() ?? string.Empty
                }).ToList(),

                ComponentHighlights = deleting ? new() : this.thisComponentHighlightRows.Select(row => new ComponentHighlightEntry
                {
                    SchematicName = row.SchematicName?.Trim() ?? string.Empty,
                    BoardLabel = ResolveRowBoardLabel(row.BoardLabel, effectiveBoardLabel),
                    X = row.X?.Trim() ?? string.Empty,
                    Y = row.Y?.Trim() ?? string.Empty,
                    Width = row.Width?.Trim() ?? string.Empty,
                    Height = row.Height?.Trim() ?? string.Empty
                }).ToList(),

                ComponentLocalFiles = deleting ? new() : this.thisComponentLocalFileRows.Select(row => new ContributionComponentLocalFileRow
                {
                    UuidV4 = row.UuidV4?.Trim() ?? string.Empty,
                    BoardLabel = ResolveRowBoardLabel(row.BoardLabel, effectiveBoardLabel),
                    Name = row.Name?.Trim() ?? string.Empty,
                    FileLocation = row.FileLocation?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty
                }).ToList(),

                ComponentLinks = deleting ? new() : this.thisComponentLinkRows.Select(row => new ContributionComponentLinkRow
                {
                    UuidV4 = row.UuidV4?.Trim() ?? string.Empty,
                    BoardLabel = ResolveRowBoardLabel(row.BoardLabel, effectiveBoardLabel),
                    Name = row.Name?.Trim() ?? string.Empty,
                    Url = row.Url?.Trim() ?? string.Empty
                }).ToList(),

                BoardLocalFiles = this.thisBoardLocalFileRows.Select(row => new ContributionBoardLocalFileRow
                {
                    UuidV4 = row.UuidV4?.Trim() ?? string.Empty,
                    Category = row.Category?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    FileLocation = row.FileLocation?.Trim() ?? string.Empty,
                    File = row.File?.Trim() ?? string.Empty
                }).ToList(),

                BoardLinks = this.thisBoardLinkRows.Select(row => new ContributionBoardLinkRow
                {
                    UuidV4 = row.UuidV4?.Trim() ?? string.Empty,
                    Category = row.Category?.Trim() ?? string.Empty,
                    Name = row.Name?.Trim() ?? string.Empty,
                    Url = row.Url?.Trim() ?? string.Empty
                }).ToList()
            };
        }

        // ###########################################################################################
        // Builds the plain-text feedback summary sent alongside the zipped JSON payload.
        // ###########################################################################################
        private string BuildContributionFeedbackText(string comment)
        {
            return ContributionPackaging.BuildFeedbackText(
                this.thisHardwareName,
                this.thisBoardName,
                this.ResolveComponentDisplayText(),
                this.thisComponentUuidV4,
                this.thisLocalRegion,
                comment,
                this.thisIsDeleteComponent);
        }

        // ###########################################################################################
        // Resolves the files referenced by the edited rows and stamps each payload row with the
        // zip entry its file will be packed under, so the server can locate every submitted file
        // exactly. Returns the distinct attachments to write into the zip.
        // ###########################################################################################
        private IReadOnlyList<ContributionAttachment> AssignZipEntriesToPayload(ComponentContributionPayload payload)
        {
            var references = new List<ContributionFileReference>();

            // The payload row lists were produced from these collections via 1:1 Select calls in
            // BuildPayload, so index alignment between source rows and payload rows is guaranteed.
            references.AddRange(this.thisComponentImageRows.Select(row => new ContributionFileReference
            {
                SectionFolder = "ComponentImages",
                ResolvedSourcePath = this.ResolveExistingFilePath(GetStoredFilePath(row))
            }));

            references.AddRange(this.thisComponentLocalFileRows.Select(row => new ContributionFileReference
            {
                SectionFolder = "ComponentLocalFiles",
                ResolvedSourcePath = this.ResolveExistingFilePath(GetStoredFilePath(row))
            }));

            references.AddRange(this.thisBoardLocalFileRows.Select(row => new ContributionFileReference
            {
                SectionFolder = "BoardLocalFiles",
                ResolvedSourcePath = this.ResolveExistingFilePath(GetStoredFilePath(row))
            }));

            var plan = ContributionPackaging.AssignZipEntries(references);

            int entryIndex = 0;

            foreach (var row in payload.ComponentImages)
            {
                row.ZipEntry = plan.EntryNames[entryIndex++];
            }

            foreach (var row in payload.ComponentLocalFiles)
            {
                row.ZipEntry = plan.EntryNames[entryIndex++];
            }

            foreach (var row in payload.BoardLocalFiles)
            {
                row.ZipEntry = plan.EntryNames[entryIndex++];
            }

            return plan.Attachments;
        }

        // ###########################################################################################
        // Resolves an edited file path so it can be verified for existence and attached.
        // Accepts both relative paths (resolved against data-root) and external absolute paths.
        // ###########################################################################################
        private string? ResolveExistingFilePath(string pathValue)
        {
            return ContributionPackaging.ResolveExistingFilePath(this.thisDataRoot, pathValue);
        }

        // ###########################################################################################
        // Adds a UTF-8 text file entry to the output zip archive.
        // ###########################################################################################
        private static void AddTextEntryToZip(ZipArchive archive, string entryName, string content)
        {
            var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
            using var writer = new StreamWriter(entry.Open(), Encoding.UTF8);
            writer.Write(content);
        }

        // ###########################################################################################
        // Adds a file to the provided zip archive, skipping unreadable files safely.
        // ###########################################################################################
        private static void AddFileToZipSafe(ZipArchive archive, string sourcePath, string entryName)
        {
            if (!File.Exists(sourcePath))
            {
                return;
            }

            try
            {
                using var fs = new FileStream(sourcePath, FileMode.Open, FileAccess.Read, FileShare.ReadWrite);
                var entry = archive.CreateEntry(entryName, CompressionLevel.Optimal);
                using var entryStream = entry.Open();
                fs.CopyTo(entryStream);
            }
            catch
            {
                // Ignore unreadable files.
            }
        }

        // ###########################################################################################
        // Updates the status text on the UI thread using success or error styling.
        // ###########################################################################################
        private void ShowStatus(string message, bool isError)
        {
            Dispatcher.UIThread.Post(() =>
            {
                this.StatusTextBlock.Text = message;

                // The panel is coloured as well as the text, so which kind of message this is can be
                // told at a glance from the whole box rather than only from the wording.
                SetStateClass(this.StatusTextBlock, isError);
                SetStateClass(this.StatusPanel, isError);

                this.StatusPanel.IsVisible = true;
            });
        }

        // ###########################################################################################
        // Puts a control into the success or error state, swapping the one class for the other.
        // Guarded because Classes is a plain list - adding twice would leave a duplicate that one
        // removal cannot undo.
        // ###########################################################################################
        private static void SetStateClass(Control control, bool isError)
        {
            string wanted = isError ? "error" : "success";
            string unwanted = isError ? "success" : "error";

            if (!control.Classes.Contains(wanted))
            {
                control.Classes.Add(wanted);
            }

            control.Classes.Remove(unwanted);
        }

        // ###########################################################################################
        // Refreshes all component image previews from the currently edited file paths.
        // ###########################################################################################
        private void RefreshAllComponentImagePreviews()
        {
            foreach (var row in this.thisComponentImageRows)
            {
                this.RefreshComponentImagePreview(row);
            }
        }

        // ###########################################################################################
        // Refreshes a single component image preview from its current file path.
        // ###########################################################################################
        private void RefreshComponentImagePreview(ContributionComponentImageRow row)
        {
            DisposeComponentImagePreview(row);

            string fullPath = !string.IsNullOrWhiteSpace(row.OriginalFilePath)
                ? row.OriginalFilePath
                : (string.IsNullOrWhiteSpace(row.FileLocation)
                    ? row.File
                    : Path.Combine(row.FileLocation, row.File ?? string.Empty));

            string? resolvedPath = this.ResolveExistingFilePath(fullPath);
            if (string.IsNullOrWhiteSpace(resolvedPath))
            {
                row.PreviewStatusText = string.Empty;
                return;
            }

            try
            {
                row.PreviewImage = new Bitmap(resolvedPath);
                row.PreviewStatusText = string.Empty;
            }
            catch (Exception ex)
            {
                row.PreviewStatusText = string.Empty;
                Logger.Warning($"Failed to load contribution image preview [{resolvedPath}] - [{ex.Message}]");
            }
        }

        // ###########################################################################################
        // Disposes the currently loaded preview image for a component image row.
        // ###########################################################################################
        private static void DisposeComponentImagePreview(ContributionComponentImageRow row)
        {
            row.PreviewImage?.Dispose();
            row.PreviewImage = null;
        }

        // ###########################################################################################
        // Disposes loaded preview images when the contribution window closes.
        // ###########################################################################################
        private void OnWindowClosed(object? sender, EventArgs e)
        {
            foreach (var row in this.thisComponentImageRows)
            {
                DisposeComponentImagePreview(row);
            }
        }

        // ###########################################################################################
        // Applies a newly selected file path to the corresponding row model.
        // ###########################################################################################
        private void ApplySelectedFilePath(object? tag, string selectedPath)
        {
            switch (tag)
            {
                case ContributionComponentImageRow componentImageRow:
                    this.ApplySelectedFilePathToRow(componentImageRow, selectedPath);
                    this.RefreshComponentImagePreview(componentImageRow);

                    // Only displayable images reach this point, so the row is fixed by definition.
                    componentImageRow.HasFileError = false;
                    componentImageRow.FileErrorText = string.Empty;
                    break;

                case ContributionComponentLocalFileRow componentLocalFileRow:
                    this.ApplySelectedFilePathToRow(componentLocalFileRow, selectedPath);
                    break;

                case ContributionBoardLocalFileRow boardLocalFileRow:
                    this.ApplySelectedFilePathToRow(boardLocalFileRow, selectedPath);
                    break;
            }
        }

        // ###########################################################################################
        // Opens the file picker when a row's read-only file box is clicked.
        // ###########################################################################################
        private async void OnFileTextBoxPointerReleased(object? sender, PointerReleasedEventArgs e)
        {
            e.Handled = true;

            if (sender is TextBox { Tag: IContributionFileRow row })
            {
                await this.PickFileForRowAsync(row);
            }
        }

        // ###########################################################################################
        // Opens the file picker from the browse button beside a row's file box.
        // ###########################################################################################
        private async void OnBrowseFileClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Button { Tag: IContributionFileRow row })
            {
                await this.PickFileForRowAsync(row);
            }
        }

        // The component image picker offers only the formats the application can draw, so an
        // unviewable file cannot be picked by accident. See ContributionPackaging for the set.
        private static readonly FilePickerFileType DisplayableImageFileType = new("Image files")
        {
            Patterns = ContributionPackaging.DisplayableImageExtensions.Select(extension => "*" + extension).ToArray(),
            MimeTypes = new[] { "image/*" },
            AppleUniformTypeIdentifiers = new[] { "public.image" }
        };

        // ###########################################################################################
        // Opens a file picker for any file-backed row and applies the selected path. Component
        // image rows are restricted to displayable image formats; the other sections take any file.
        // ###########################################################################################
        private async Task PickFileForRowAsync(IContributionFileRow row)
        {
            var topLevel = TopLevel.GetTopLevel(this);
            if (topLevel == null)
            {
                return;
            }

            bool imagesOnly = row is ContributionComponentImageRow;

            string? currentPath = this.GetCurrentFilePath(row);
            string? suggestedStartLocation = this.GetSuggestedStartLocation(currentPath);

            var options = new FilePickerOpenOptions
            {
                Title = imagesOnly ? "Select image file" : "Select file",
                AllowMultiple = false,
                FileTypeFilter = imagesOnly ? new[] { DisplayableImageFileType } : null
            };

            if (!string.IsNullOrWhiteSpace(suggestedStartLocation) && Directory.Exists(suggestedStartLocation))
            {
                try
                {
                    options.SuggestedStartLocation = await topLevel.StorageProvider.TryGetFolderFromPathAsync(suggestedStartLocation);
                }
                catch
                {
                }
            }

            var files = await topLevel.StorageProvider.OpenFilePickerAsync(options);
            if (files == null || files.Count == 0)
            {
                return;
            }

            string selectedPath = files[0].Path.LocalPath;

            // The dialog filter is only a suggestion - a name typed into the file box gets past it -
            // so the format is verified here rather than trusted, and the row is left untouched.
            if (imagesOnly && !ContributionPackaging.IsDisplayableImageFile(selectedPath))
            {
                this.ShowStatus(
                    $"[{Path.GetFileName(selectedPath)}] is not an image the application can display - use one of: " +
                    string.Join(", ", ContributionPackaging.DisplayableImageExtensions),
                    true);
                return;
            }

            // ApplySelectedFilePath writes row.File, and the two-way bound text box follows it,
            // so the box shows the same value whether the picker was opened from it or the button.
            this.ApplySelectedFilePath(row, selectedPath);
        }

        // ###########################################################################################
        // Returns the current file path for the given tagged row object.
        // ###########################################################################################
        private string? GetCurrentFilePath(object? tag)
        {
            return tag switch
            {
                IContributionFileRow fileRow => GetStoredFilePath(fileRow),
                _ => null
            };
        }

        // ###########################################################################################
        // Computes the best starting directory for the file picker based on the current file value.
        // ###########################################################################################
        private string? GetSuggestedStartLocation(string? currentPath)
        {
            string trimmed = currentPath?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(trimmed))
            {
                return Directory.Exists(this.thisDataRoot) ? this.thisDataRoot : null;
            }

            if (Path.IsPathRooted(trimmed))
            {
                string? rootedDirectory = Path.GetDirectoryName(trimmed);
                if (!string.IsNullOrWhiteSpace(rootedDirectory) && Directory.Exists(rootedDirectory))
                {
                    return rootedDirectory;
                }
            }

            string combinedPath = Path.Combine(this.thisDataRoot, trimmed.Replace('/', Path.DirectorySeparatorChar));
            string? combinedDirectory = Path.GetDirectoryName(combinedPath);
            if (!string.IsNullOrWhiteSpace(combinedDirectory) && Directory.Exists(combinedDirectory))
            {
                return combinedDirectory;
            }

            return Directory.Exists(this.thisDataRoot) ? this.thisDataRoot : null;
        }

        // ###########################################################################################
        // Persists the shared email address when the field loses focus and the value is valid.
        // ###########################################################################################
        private void OnEmailTextBoxLostFocus(object? sender, RoutedEventArgs e)
        {
            string email = this.EmailTextBox.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrEmpty(email) || Regex.IsMatch(email, @"^[^@\s]+@[^@\s]+\.[^@\s]+$"))
            {
                UserSettings.ContactEmail = email;
            }
        }

        // ###########################################################################################
        // Inserts a new row at the top of a collection so it becomes visible immediately.
        // ###########################################################################################
        private static void InsertRowAtTop<T>(ObservableCollection<T> collection, T row)
        {
            collection.Insert(0, row);
        }

        // ###########################################################################################
        // Scrolls the contribution editor to the top of the main content area.
        // ###########################################################################################
        private void OnScrollToTopClick(object? sender, RoutedEventArgs e)
        {
            this.MainScrollViewer.Offset = new Vector(this.MainScrollViewer.Offset.X, 0);
        }

        // ###########################################################################################
        // Scrolls the contribution editor to the bottom of the main content area.
        // ###########################################################################################
        private void OnScrollToBottomClick(object? sender, RoutedEventArgs e)
        {
            double bottomOffset = Math.Max(0, this.MainScrollViewer.Extent.Height - this.MainScrollViewer.Viewport.Height);
            this.MainScrollViewer.Offset = new Vector(this.MainScrollViewer.Offset.X, bottomOffset);
        }

        // ###########################################################################################
        // Updates the visible row counters for the editable contribution sections.
        // ###########################################################################################
        private void UpdateSectionCounters()
        {
            this.ComponentImagesCountTextBlock.Text = $"({this.thisComponentImageRows.Count})";
            this.ComponentLocalFilesCountTextBlock.Text = $"({this.thisComponentLocalFileRows.Count})";
            this.ComponentLinksCountTextBlock.Text = $"({this.thisComponentLinkRows.Count})";
            this.BoardLocalFilesCountTextBlock.Text = $"({this.thisBoardLocalFileRows.Count})";
            this.BoardLinksCountTextBlock.Text = $"({this.thisBoardLinkRows.Count})";
        }

        // ###########################################################################################
        // Extracts a file-location value from a row, with legacy fallback from the stored file path.
        // ###########################################################################################
        private static string GetExistingFileLocation(object row, string? filePath)
        {
            var propertyInfo = row.GetType().GetProperty("FileLocation");
            if (propertyInfo != null && propertyInfo.GetValue(row) is string fileLocation && !string.IsNullOrWhiteSpace(fileLocation))
            {
                return fileLocation.Trim();
            }

            if (string.IsNullOrWhiteSpace(filePath))
            {
                return string.Empty;
            }

            try
            {
                string? directory = Path.GetDirectoryName(filePath);
                return string.IsNullOrWhiteSpace(directory) ? string.Empty : directory.Replace('\\', '/');
            }
            catch
            {
                return string.Empty;
            }
        }

        // ###########################################################################################
        // Builds the effective source path for a file row from original path or location + filename.
        // ###########################################################################################
        private static string GetStoredFilePath(IContributionFileRow row)
        {
            if (!string.IsNullOrWhiteSpace(row.OriginalFilePath))
            {
                return row.OriginalFilePath;
            }

            return string.IsNullOrWhiteSpace(row.FileLocation)
                ? row.File
                : Path.Combine(row.FileLocation, row.File ?? string.Empty);
        }

        // ###########################################################################################
        // Applies the selected file to a row while keeping the source path separate from the filename.
        // ###########################################################################################
        private void ApplySelectedFilePathToRow(IContributionFileRow row, string selectedPath)
        {
            row.File = Path.GetFileName(selectedPath);
            row.OriginalFilePath = selectedPath;

            try
            {
                string? dir = Path.GetDirectoryName(selectedPath);
                if (!string.IsNullOrWhiteSpace(dir) &&
                    !string.IsNullOrWhiteSpace(this.thisDataRoot) &&
                    dir.StartsWith(this.thisDataRoot, StringComparison.OrdinalIgnoreCase))
                {
                    string rel = Path.GetRelativePath(this.thisDataRoot, dir);
                    row.FileLocation = (rel != "." && rel != "") ? rel.Replace('\\', '/') : string.Empty;
                }
                // We intentionally do NOT overwrite FileLocation if the user selects an external file.
                // The file should retain whatever drop-down folder the user selected.
            }
            catch
            {
            }

            this.SetAvailableFileLocations(row);

            string updatedLocation = row.FileLocation;
            row.FileLocation = string.Empty;
            row.FileLocation = updatedLocation;
        }

        // ###########################################################################################
        // Populates a row-specific file-location list and injects the current folder if missing.
        // The resulting list is kept sorted, including any injected non-end-folder path.
        // ###########################################################################################
        private void SetAvailableFileLocations(IContributionFileRow row)
        {
            row.AvailableFileLocations.Clear();

            string currentFileLocation = row.FileLocation?.Trim() ?? string.Empty;

            var folders = this.AvailableEndFolders
                .Concat(string.IsNullOrWhiteSpace(currentFileLocation) ? Enumerable.Empty<string>() : new[] { currentFileLocation })
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .OrderBy(folder => folder, StringComparer.OrdinalIgnoreCase);

            foreach (var folder in folders)
            {
                row.AvailableFileLocations.Add(folder);
            }
        }

    }
}