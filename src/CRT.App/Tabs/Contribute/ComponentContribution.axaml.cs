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

        // The "File location" box is marked red - see ContributionComponentImageRow.HasLocationError.
        bool HasLocationError { get; set; }

        // ###########################################################################################
        // Tells the bound drop-down to read FileLocation again, WITHOUT changing it (code review,
        // 2026-09-26). The list of available folders is rebuilt underneath it, and a ComboBox whose
        // ItemsSource has just been repopulated drops a selection it no longer recognises.
        //
        // This used to be done by assigning string.Empty and the value back. That goes through the
        // setter, whose non-blank arm answers the row's location mark - so browsing for a new file
        // silently cleared a still-unfixed "this folder cannot be written to" mark and its message,
        // and the red only came back after another refused save.
        // ###########################################################################################
        void NotifyFileLocationChanged();
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

                    // Choosing a folder answers the location mark, and the row's own mark and text
                    // were that same problem - see HasLocationError.
                    if (this.HasLocationError && !string.IsNullOrWhiteSpace(value))
                    {
                        this.HasLocationError = false;
                        this.HasFileError = false;
                        this.FileErrorText = string.Empty;
                    }
                }
            }
        }

        // Re-raises FileLocation's change notification without touching the value - see the
        // interface. Named for the property rather than using [CallerMemberName], which would
        // report this method's own name.
        public void NotifyFileLocationChanged() => this.OnPropertyChanged(nameof(this.FileLocation));

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

        // ###########################################################################################
        // Set by "Save to draft" when a NEWLY PICKED file has no folder, or one outside this board
        // and the shared folders (ContributionFileLocations.CheckNewFile): the "File location" box
        // turns red. Cleared the moment a folder is chosen. Never part of the uploaded payload.
        // ###########################################################################################
        private bool thisHasLocationError;
        [JsonIgnore]
        public bool HasLocationError
        {
            get => this.thisHasLocationError;
            set
            {
                if (this.thisHasLocationError != value)
                {
                    this.thisHasLocationError = value;
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

                    // Choosing a folder answers the location mark - see HasLocationError.
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        this.HasLocationError = false;
                    }
                }
            }
        }

        // Re-raises FileLocation's change notification without touching the value - see the
        // interface. Named for the property rather than using [CallerMemberName], which would
        // report this method's own name.
        public void NotifyFileLocationChanged() => this.OnPropertyChanged(nameof(this.FileLocation));

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

        // ###########################################################################################
        // Set by "Save to draft" when a NEWLY PICKED file has no folder, or one outside this board
        // and the shared folders (ContributionFileLocations.CheckNewFile): the "File location" box
        // turns red. Cleared the moment a folder is chosen. Never part of the uploaded payload.
        // ###########################################################################################
        private bool thisHasLocationError;
        [JsonIgnore]
        public bool HasLocationError
        {
            get => this.thisHasLocationError;
            set
            {
                if (this.thisHasLocationError != value)
                {
                    this.thisHasLocationError = value;
                    this.OnPropertyChanged();
                }
            }
        }
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

                    // Choosing a folder answers the location mark - see HasLocationError.
                    if (!string.IsNullOrWhiteSpace(value))
                    {
                        this.HasLocationError = false;
                    }
                }
            }
        }

        // Re-raises FileLocation's change notification without touching the value - see the
        // interface. Named for the property rather than using [CallerMemberName], which would
        // report this method's own name.
        public void NotifyFileLocationChanged() => this.OnPropertyChanged(nameof(this.FileLocation));

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

        // ###########################################################################################
        // Set by "Save to draft" when a NEWLY PICKED file has no folder, or one outside this board
        // and the shared folders (ContributionFileLocations.CheckNewFile): the "File location" box
        // turns red. Cleared the moment a folder is chosen. Never part of the uploaded payload.
        // ###########################################################################################
        private bool thisHasLocationError;
        [JsonIgnore]
        public bool HasLocationError
        {
            get => this.thisHasLocationError;
            set
            {
                if (this.thisHasLocationError != value)
                {
                    this.thisHasLocationError = value;
                    this.OnPropertyChanged();
                }
            }
        }
    }

    public sealed class ContributionBoardLinkRow
    {
        public string UuidV4 { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public string Name { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
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

        // ###########################################################################################
        // The board-scoped rows as this window LOADED them (or last saved them) - the baseline
        // ComponentBoardWriter merges against, so a board link or file added elsewhere (in Excel,
        // most likely) while this window was open survives a save that did not touch them. Built
        // through the same row builders as the edited rows, so the two compare like for like.
        // Moved forward after every successful save: the next save's baseline is what this one
        // wrote.
        // ###########################################################################################
        private IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow> thisBoardLocalFileRowsAtOpen = [];
        private IReadOnlyList<ComponentDraftWriter.LinkDraftRow> thisBoardLinkRowsAtOpen = [];

        private string thisHardwareName = string.Empty;
        private string thisBoardName = string.Empty;
        private string thisBoardExcelFile = string.Empty;
        private string thisBoardRevisionDate = string.Empty;
        private string thisLocalRegion = string.Empty;
        private string thisBoardLabel = string.Empty;
        private string thisComponentDisplayText = string.Empty;
        private string thisDataRoot = string.Empty;
        private string thisComponentUuidV4 = string.Empty;

        // ###########################################################################################
        // *** WHAT REDRAWS THE BOARD AFTER A SAVE, AND WHY IT IS A CALLBACK. ***
        //
        // A save writes the draft workbook and nothing else. The board the rest of the application
        // is showing came out of DataManager's cache, so until that cache is cleared and the board
        // re-read, every surface keeps showing the PRE-save data - the component popup, the
        // component list, and this window itself when it is reopened. Reported (2026-09-23): a
        // changed "Short description" saved successfully and then appeared nowhere.
        //
        // It did not exist before Phase 6a because this window used to POST to the contribution
        // server: nothing local changed, so there was nothing local to refresh. Turning it into a
        // local draft save made a refresh necessary, and that step was missed.
        //
        // A callback rather than a MainWindow reference, for the same reason TabWorkbooks activates
        // a workbook through an Action: this window is constructed directly by the headless tests,
        // which never build Main. Left null it simply does nothing.
        //
        // Main sets it to ReloadCurrentBoardFromDisk, the SAME entry point the label editor's save
        // uses - so both local-edit paths refresh through one implementation rather than two.
        //
        // Invoked at the end of SaveComponentToDraftAsync rather than in the click handler, so that
        // every route to a save carries the refresh with it.
        // ###########################################################################################
        private Action? thisRefreshBoardAfterSave;

        // ###########################################################################################
        // What the application does once "Save to draft" has landed: Main closes this window and
        // shows the Drafts tab (owner request, 2026-09-24), because that is where the next
        // step is - reviewing the draft, then submitting it. Staying on a maximized window saying
        // "Saved" left the contributor to find the Drafts tab on their own.
        //
        // A callback for the same reason as thisRefreshBoardAfterSave: the tests build this window
        // without Main. Only after a SUCCESSFUL save - a refused or failed one keeps the window,
        // and its message, in front of the contributor.
        // ###########################################################################################
        private Action? thisAfterSaved;

        // Asks whether the Drafts tab's table holds unsaved edits for a board (its ExcelDataFile) -
        // Main wires it to TabDrafts.HasUnsavedTableEditsFor. Null (as in tests building the window
        // alone) means no table to wait for.
        private Func<string, bool>? thisHasUnsavedTableEditsFor;

        // Replaces the notice's dialog in tests, which cannot wait on ShowDialog.
        internal Func<Task>? ShowSavingBlockedOverrideForTests { get; set; }

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

        public ComponentContributionWindow()
        {
            this.InitializeComponent();

            this.ComponentRowsItemsControl.ItemsSource = this.thisComponentRows;
            this.ComponentImageRowsItemsControl.ItemsSource = this.thisComponentImageRows;
            this.ComponentLocalFileRowsItemsControl.ItemsSource = this.thisComponentLocalFileRows;
            this.ComponentLinkRowsItemsControl.ItemsSource = this.thisComponentLinkRows;
            this.BoardLocalFileRowsItemsControl.ItemsSource = this.thisBoardLocalFileRows;
            this.BoardLinkRowsItemsControl.ItemsSource = this.thisBoardLinkRows;

            this.Closed += this.OnWindowClosed;

            this.UpdateSectionCounters();
        }

        // ###########################################################################################
        // Tells the window how to make the application re-read the board after a save. See
        // thisRefreshBoardAfterSave for why this is a callback and what breaks without it.
        // ###########################################################################################
        public void SetRefreshBoardAfterSave(Action refreshBoardAfterSave)
        {
            this.thisRefreshBoardAfterSave = refreshBoardAfterSave;
        }

        // Tells the window what to do once a save has landed. See thisAfterSaved.
        public void SetAfterSaved(Action afterSaved)
        {
            this.thisAfterSaved = afterSaved;
        }

        // Tells the window how to ask about the Drafts tab's table. See SubmitAsync.
        public void SetUnsavedTableEditsCheck(Func<string, bool> hasUnsavedTableEditsFor)
        {
            this.thisHasUnsavedTableEditsFor = hasUnsavedTableEditsFor;
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

            // ###########################################################################################
            // *** ALWAYS EMPTY SINCE THE UUID COLUMN WAS DROPPED (2026-09-23). ***
            //
            // BoardData no longer carries UuidV4 - the column is gone from the board workbooks
            // entirely - so there is nothing to read here. The field is kept because it was part
            // of the LEGACY contribution format, and that format was not changed: it was retired
            // with this pipeline rather than migrated (owner decision, 2026-09-23), so touching it
            // would have been work spent on something about to go.
            // ###########################################################################################
            this.thisComponentUuidV4 = string.Empty;
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
        // Discovers every folder a file row can be placed in - the data tree's AND the drafts
        // tree's, so a board that exists only as a draft offers its own folders too. See
        // ContributionFileLocations for why both, and why that needed nothing else to change.
        //
        // *** ONLY THIS BOARD'S FOLDERS AND THE SHARED ONES (owner request, 2026-09-25) - see
        // ContributionFileLocations.WritableBy. *** A row whose file already sits elsewhere keeps
        // its own folder in its list regardless (SetAvailableFileLocations adds it back).
        //
        // Rebuilt on every open, so a board created a moment ago is already in the list. Needs
        // thisBoardExcelFile, which names the board - ApplyBoardContext sets it first.
        // ###########################################################################################
        private void PopulateEndFolders(string dataRoot)
        {
            this.AvailableEndFolders.Clear();

            IEnumerable<string> writable = ContributionFileLocations.WritableBy(
                BoardDescriptorRules.BoardIdFromExcelDataFile(this.thisBoardExcelFile),
                ContributionFileLocations.FindEndFolders(dataRoot, DraftManager.DraftsRoot));

            foreach (string folder in writable)
            {
                this.AvailableEndFolders.Add(folder);
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
                    ? "Component editor - delete component"
                    : $"Component editor - delete {this.thisBoardLabel}";
            }
            else if (this.thisIsNewComponent)
            {
                this.Title = "Component editor - new component";
            }
            else
            {
                this.Title = string.IsNullOrWhiteSpace(this.thisBoardLabel)
                    ? "Component editor"
                    : $"Component editor - {this.thisBoardLabel}";
            }

            // Every way of opening the window carries the same warning - what is saved goes into
            // your own local draft, not straight into anyone else's data - so only the wording
            // changes with the mode.
            this.ContributionNoticeHeadingTextBlock.Text = this.thisIsDeleteComponent
                ? "You are deleting this component in your local draft"
                : this.thisIsNewComponent
                    ? "You are adding a component this board does not have yet"
                    : "You are modifying an existing component";

            // The delete notice replaces both of the others rather than joining them: in delete
            // mode nothing is being added or modified, so a line about the edits taking effect
            // describes a save that is not being made.
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
                    Category = row.Category,
                    Name = row.Name,
                    Url = row.Url
                });
            }

            this.thisBoardLocalFileRowsAtOpen = this.BuildLocalFileDraftRows(this.thisBoardLocalFileRows, string.Empty, isComponentScoped: false);
            this.thisBoardLinkRowsAtOpen = this.BuildLinkDraftRows(this.thisBoardLinkRows, string.Empty, isComponentScoped: false);

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
        // It is a toggle rather than a one-way door because nothing else in this window stands
        // between a mistaken click and a deletion saved to the draft; being able to back out is what
        // makes the button safe to press in order to see what it would remove. The draft is the
        // second guard: nothing leaves this computer until it is submitted from the Drafts tab.
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
            this.SubmitButton.Content = deleting ? "Save deletion to draft" : "Save to draft";

            this.PopulateHeader();

            if (deleting)
            {
                // The section holding the button is where the mode is left again, so it must not be
                // folded away behind the contributor while they are in it.
                //
                // Nothing else moves. Scrolling down would take the contributor away from the very
                // rows the notice has just told them are about to be deleted, at the moment they
                // most need to look at them - and it moves the button they may have clicked by
                // mistake off screen.
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
        // Validates the edited rows and saves them into the current board's local draft
        // (NewContributeStrategy.md Phase 2, session 2c - this window used to zip its payload and
        // POST it to the contribution server; see ComponentDraftWriter's own header for why the
        // save target changed and the UI otherwise did not).
        // ###########################################################################################
        private async void OnSubmitClick(object? sender, RoutedEventArgs e) => await this.SubmitAsync();

        // The whole of "Save to draft", awaitable - so a test can wait for a save to finish, which
        // an async void click handler cannot offer.
        private async Task SubmitAsync()
        {
            // *** NOT WHILE THE DRAFTS TAB'S TABLE HOLDS UNSAVED EDITS FOR THIS BOARD (owner
            // request, 2026-09-24). *** Both write the same draft: saving here would make the table's
            // own save refused, and its edits lost. So nothing is saved and a notice sends the
            // contributor to the Drafts tab first - this window stays open exactly as it is.
            if (this.thisHasUnsavedTableEditsFor?.Invoke(this.thisBoardExcelFile) == true)
            {
                await (this.ShowSavingBlockedOverrideForTests?.Invoke() ?? UnsavedTableEditsWindow.ShowSavingBlockedAsync(this));
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

                // After the image check, which first clears the mark of every image row it passes.
                var locationProblem = this.ValidateNewFileLocations();
                if (locationProblem != null)
                {
                    locationProblem.Value.Reveal();
                    this.ShowStatus(locationProblem.Value.Message, true);
                    return;
                }
            }

            if (string.IsNullOrWhiteSpace(this.thisBoardExcelFile))
            {
                this.ShowStatus("Could not resolve which board this draft belongs to.", true);
                return;
            }

            this.SubmitButton.IsEnabled = false;

            bool saveSucceeded = false;

            try
            {
                this.ShowStatus(string.Empty, false);

                // Under this window's "please wait" overlay (2026-09-28). Writing a draft cannot be
                // stopped halfway, so past the limit it carries on and the line below says so.
                saveSucceeded = await BusyOverlay.RunLocalAsync(
                    this,
                    CrtWaitWording.SavingDraft,
                    this.SaveComponentToDraftAsync,
                    () => this.ShowStatus(WaitWording.StillRunning("Saving to your draft"), false));

                if (saveSucceeded)
                {
                    this.ShowStatus(this.BuildSaveSuccessText(), false);
                    this.thisAfterSaved?.Invoke();
                }
                else
                {
                    this.ShowStatus("Could not save your draft - please check the logfile for details.", true);
                }
            }
            catch (Exception ex)
            {
                Logger.Warning($"Exception while saving component draft: {ex}");
                this.ShowStatus("Could not save your draft - please check the logfile for details.", true);
            }
            finally
            {
                this.ApplySubmissionOutcome(saveSucceeded);
            }
        }

        // ###########################################################################################
        // Copies any newly picked attachment files into the draft folder, applies the editing
        // session to the draft's BOARD and writes it back.
        //
        // *** IT NO LONGER LOADS THE PURE OFFICIAL BOARD (Phase 6, 2026-09-23). *** That was
        // needed to decide Added-vs-Modified per row against the published copy, a distinction
        // that only existed to describe an overlay. A draft is now the board, so the rows the
        // window hands back simply ARE this component's rows.
        //
        // The board it edits comes from DraftWorkbookStore.Edit, which reads the workbook as it
        // is ON DISK - so an edit the contributor made in Excel while this window was open is
        // not silently overwritten.
        //
        // Runs off the UI thread, mirroring TabSchematics.LabelEditor.cs's own save shape.
        //
        // *** EVERY ROW IS SNAPSHOT ON THE UI THREAD FIRST (code review, 2026-09-25). *** The
        // background half used to enumerate this window's ObservableCollections itself, while only
        // the Save button was disabled - so adding or removing a row, or a file picker completing,
        // during the save threw "Collection was modified" or wrote a torn snapshot into the draft.
        // The pool thread now sees plain immutable lists and nothing else of this window's.
        // ###########################################################################################
        private async Task<bool> SaveComponentToDraftAsync()
        {
            string draftFolder = DraftFolderLayout.GetBoardFolder(
                DraftManager.DraftsRoot,
                this.thisBoardExcelFile);

            if (string.IsNullOrWhiteSpace(draftFolder))
            {
                Logger.Warning("Component draft save failed - could not resolve the drafts folder for this board");
                return false;
            }

            // *** A BOARD WITH NO DRAFT IS SEEDED FIRST, as on every other save path. *** A draft
            // is a complete copy of the published board; starting an empty one here would produce
            // a draft that reads as "every published row deleted". Copied from the published FILE
            // as it is now, never from the board cache - see DraftSeeder.SeedFromPublishedFile.
            if (!DraftBoardSource.HasDraft(DraftManager.DraftsRoot, this.thisBoardExcelFile))
            {
                string seedDraftsRoot = DraftManager.DraftsRoot;
                string seedDataRoot = this.thisDataRoot;
                string seedExcelDataFile = this.thisBoardExcelFile;

                DraftSeedResult seeded = await Task.Run(() => DraftSeeder.SeedFromPublishedFile(seedDraftsRoot, seedDataRoot, seedExcelDataFile));

                if (!seeded.Created)
                {
                    Logger.Warning($"Component draft save failed - could not seed a draft: [{seeded.Reason}]");
                    return false;
                }
            }

            string effectiveBoardLabel = this.ResolveEffectiveBoardLabel();

            // The snapshot - see this method's header. Everything below the Task.Run reads only
            // these locals.
            string draftsRoot = DraftManager.DraftsRoot;
            string boardExcelFile = this.thisBoardExcelFile;
            string editedRegion = this.thisLocalRegion;
            bool isDelete = this.thisIsDeleteComponent;
            IReadOnlyList<(string Source, string RelativeFile)> filesToCopy = this.CollectNewlyAttachedFiles();
            IReadOnlyList<ComponentDraftWriter.ComponentDraftRow> componentRows = this.BuildComponentDraftRows(effectiveBoardLabel);
            IReadOnlyList<ComponentDraftWriter.ComponentImageDraftRow> imageRows = this.BuildComponentImageDraftRows(effectiveBoardLabel);
            IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow> componentFileRows = this.BuildLocalFileDraftRows(this.thisComponentLocalFileRows, effectiveBoardLabel, isComponentScoped: true);
            IReadOnlyList<ComponentDraftWriter.LinkDraftRow> componentLinkRows = this.BuildLinkDraftRows(this.thisComponentLinkRows, effectiveBoardLabel, isComponentScoped: true);
            IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow> boardFileRows = this.BuildLocalFileDraftRows(this.thisBoardLocalFileRows, string.Empty, isComponentScoped: false);
            IReadOnlyList<ComponentDraftWriter.LinkDraftRow> boardLinkRows = this.BuildLinkDraftRows(this.thisBoardLinkRows, string.Empty, isComponentScoped: false);
            IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow> boardFileRowsAtOpen = this.thisBoardLocalFileRowsAtOpen;
            IReadOnlyList<ComponentDraftWriter.LinkDraftRow> boardLinkRowsAtOpen = this.thisBoardLinkRowsAtOpen;

            bool writeSucceeded = await Task.Run(() =>
            {
                try
                {
                    if (isDelete)
                    {
                        return DraftWorkbookStore.Edit(
                            draftsRoot,
                            boardExcelFile,
                            board => ComponentBoardWriter.ApplyComponentDelete(board, effectiveBoardLabel));
                    }

                    // Newly picked files (OriginalFilePath set, meaning the file picker touched
                    // this row) are copied into the draft folder BEFORE the rows are built, so the
                    // stored relative path already names bytes that exist.
                    ComponentContributionWindow.CopyAttachedFilesIntoDraft(draftFolder, filesToCopy);

                    return DraftWorkbookStore.Edit(
                        draftsRoot,
                        boardExcelFile,
                        board => ComponentBoardWriter.ApplyComponentSave(
                            board,
                            effectiveBoardLabel,
                            componentRows,
                            imageRows,
                            componentFileRows,
                            componentLinkRows,
                            boardFileRows,
                            boardLinkRows,

                            // The region this window was opened on. Component images are the one
                            // section it loads filtered by region, so the writer needs to know
                            // which rows were actually on screen - see ComponentBoardWriter.
                            editedRegion,

                            // What the window loaded, so board-level rows changed elsewhere while
                            // it was open are merged rather than overwritten.
                            boardFileRowsAtOpen,
                            boardLinkRowsAtOpen));
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Component draft save failed - [{ex}]");
                    return false;
                }
            });

            // The next save's baseline is what this one wrote - otherwise it would merge against
            // the state from before this save and could re-apply this save's changes.
            if (writeSucceeded && !isDelete)
            {
                this.thisBoardLocalFileRowsAtOpen = boardFileRows;
                this.thisBoardLinkRowsAtOpen = boardLinkRows;
            }

            // ###########################################################################################
            // *** THE REFRESH LIVES HERE, NOT IN THE CLICK HANDLER, so that no route to this save can
            // skip it. *** The write above changes the draft workbook and nothing else; until the
            // board cache is cleared and the board re-read, every surface keeps showing the pre-save
            // data. See thisRefreshBoardAfterSave for the bug this fixes.
            //
            // Only on success: a failed save changed nothing, and reloading would throw away nothing
            // useful but would tell the contributor - through a board that visibly redrew - that
            // something had happened when it had not.
            // ###########################################################################################
            if (writeSucceeded)
            {
                this.thisRefreshBoardAfterSave?.Invoke();
            }

            return writeSucceeded;
        }

        // ###########################################################################################
        // Copies every file-backed row's picked file into the draft's Files/ folder, keyed by the
        // SAME FileLocation/File relative path the drafted entry will store - so nothing about the
        // stored entry needs to know it came from a draft, only DraftFileResolver's read side does.
        // A row whose file was never touched (no OriginalFilePath - it already pointed at an
        // existing official or drafted file) is skipped: its bytes already live wherever they were
        // loaded from, and re-copying them would just duplicate storage for no reason.
        //
        // Split in two: CollectNewlyAttachedFiles reads this window's rows and must run on the UI
        // thread; CopyAttachedFilesIntoDraft only copies bytes and runs on the pool thread, seeing
        // nothing of the window but the list it is handed.
        // ###########################################################################################
        private IReadOnlyList<(string Source, string RelativeFile)> CollectNewlyAttachedFiles()
        {
            var files = new List<(string Source, string RelativeFile)>();

            foreach (var row in this.thisComponentImageRows.Cast<IContributionFileRow>()
                         .Concat(this.thisComponentLocalFileRows)
                         .Concat(this.thisBoardLocalFileRows))
            {
                if (string.IsNullOrWhiteSpace(row.OriginalFilePath) || string.IsNullOrWhiteSpace(row.File))
                {
                    continue;
                }

                string relativeFile = string.IsNullOrWhiteSpace(row.FileLocation)
                    ? row.File
                    : $"{row.FileLocation.Trim().Replace('\\', '/').Trim('/')}/{row.File}";

                files.Add((row.OriginalFilePath, relativeFile));
            }

            return files;
        }

        private static void CopyAttachedFilesIntoDraft(
            string draftFolder,
            IReadOnlyList<(string Source, string RelativeFile)> files)
        {
            foreach ((string source, string relativeFile) in files)
            {
                if (!File.Exists(source))
                {
                    continue;
                }

                string destination = DraftFileResolver.BuildDraftFileDestination(draftFolder, relativeFile);

                try
                {
                    string? destinationDirectory = Path.GetDirectoryName(destination);
                    if (!string.IsNullOrWhiteSpace(destinationDirectory))
                    {
                        Directory.CreateDirectory(destinationDirectory);
                    }

                    File.Copy(source, destination, overwrite: true);
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to copy attached file [{source}] into draft - [{ex.Message}]");
                }
            }
        }

        // ###########################################################################################
        // Builds ComponentDraftWriter's row DTOs from this window's own edited rows - the mechanical
        // half of the redirect, since the writer's shapes deliberately mirror these fields one for
        // one (see ComponentDraftWriter's own header).
        // ###########################################################################################
        private IReadOnlyList<ComponentDraftWriter.ComponentDraftRow> BuildComponentDraftRows(string effectiveBoardLabel)
        {
            if (this.thisIsDeleteComponent)
            {
                return Array.Empty<ComponentDraftWriter.ComponentDraftRow>();
            }

            return this.thisComponentRows.Select(row => new ComponentDraftWriter.ComponentDraftRow
            {
                BoardLabel = effectiveBoardLabel,
                FriendlyName = row.FriendlyName?.Trim() ?? string.Empty,
                TechnicalNameOrValue = row.TechnicalNameOrValue?.Trim() ?? string.Empty,
                PartNumber = row.PartNumber?.Trim() ?? string.Empty,
                Category = row.Category?.Trim() ?? string.Empty,
                Region = row.Region?.Trim() ?? string.Empty,
                Description = row.Description?.Trim() ?? string.Empty,
            }).ToList();
        }

        private IReadOnlyList<ComponentDraftWriter.ComponentImageDraftRow> BuildComponentImageDraftRows(string effectiveBoardLabel)
        {
            if (this.thisIsDeleteComponent)
            {
                return Array.Empty<ComponentDraftWriter.ComponentImageDraftRow>();
            }

            return this.thisComponentImageRows.Select(row => new ComponentDraftWriter.ComponentImageDraftRow
            {
                BoardLabel = ResolveRowBoardLabel(row.BoardLabel, effectiveBoardLabel),
                Region = row.Region?.Trim() ?? string.Empty,
                Pin = row.Pin?.Trim() ?? string.Empty,
                Name = row.Name?.Trim() ?? string.Empty,
                ExpectedOscilloscopeReading = row.ExpectedOscilloscopeReading?.Trim() ?? string.Empty,
                VoltsDiv = row.VoltsDiv?.Trim() ?? string.Empty,
                TimeDiv = row.TimeDiv?.Trim() ?? string.Empty,
                TriggerLevelVolts = row.TriggerLevelVolts?.Trim() ?? string.Empty,
                Note = row.Note?.Trim() ?? string.Empty,
                FileLocation = row.FileLocation?.Trim() ?? string.Empty,
                File = row.File?.Trim() ?? string.Empty,
            }).ToList();
        }

        // ###########################################################################################
        // Shared by ComponentLocalFiles/BoardLocalFiles - isComponentScoped picks whether the source
        // rows carry a BoardLabel (resolved against effectiveBoardLabel the same way every other
        // component-scoped section does) or a Category (board-scoped rows have none to resolve).
        // ###########################################################################################
        private IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow> BuildLocalFileDraftRows(
            IEnumerable<ContributionComponentLocalFileRow> rows,
            string effectiveBoardLabel,
            bool isComponentScoped)
        {
            if (isComponentScoped && this.thisIsDeleteComponent)
            {
                return Array.Empty<ComponentDraftWriter.LocalFileDraftRow>();
            }

            return rows.Select(row => new ComponentDraftWriter.LocalFileDraftRow
            {
                BoardLabel = isComponentScoped ? ResolveRowBoardLabel(row.BoardLabel, effectiveBoardLabel) : string.Empty,
                Name = row.Name?.Trim() ?? string.Empty,
                FileLocation = row.FileLocation?.Trim() ?? string.Empty,
                File = row.File?.Trim() ?? string.Empty,
            }).ToList();
        }

        private IReadOnlyList<ComponentDraftWriter.LocalFileDraftRow> BuildLocalFileDraftRows(
            IEnumerable<ContributionBoardLocalFileRow> rows,
            string effectiveBoardLabel,
            bool isComponentScoped)
        {
            return rows.Select(row => new ComponentDraftWriter.LocalFileDraftRow
            {
                Category = row.Category?.Trim() ?? string.Empty,
                Name = row.Name?.Trim() ?? string.Empty,
                FileLocation = row.FileLocation?.Trim() ?? string.Empty,
                File = row.File?.Trim() ?? string.Empty,
            }).ToList();
        }

        private IReadOnlyList<ComponentDraftWriter.LinkDraftRow> BuildLinkDraftRows(
            IEnumerable<ContributionComponentLinkRow> rows,
            string effectiveBoardLabel,
            bool isComponentScoped)
        {
            if (isComponentScoped && this.thisIsDeleteComponent)
            {
                return Array.Empty<ComponentDraftWriter.LinkDraftRow>();
            }

            return rows.Select(row => new ComponentDraftWriter.LinkDraftRow
            {
                BoardLabel = isComponentScoped ? ResolveRowBoardLabel(row.BoardLabel, effectiveBoardLabel) : string.Empty,
                Name = row.Name?.Trim() ?? string.Empty,
                Url = row.Url?.Trim() ?? string.Empty,
            }).ToList();
        }

        private IReadOnlyList<ComponentDraftWriter.LinkDraftRow> BuildLinkDraftRows(
            IEnumerable<ContributionBoardLinkRow> rows,
            string effectiveBoardLabel,
            bool isComponentScoped)
        {
            return rows.Select(row => new ComponentDraftWriter.LinkDraftRow
            {
                Category = row.Category?.Trim() ?? string.Empty,
                Name = row.Name?.Trim() ?? string.Empty,
                Url = row.Url?.Trim() ?? string.Empty,
            }).ToList();
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
        // Settles what the Save button does after an attempt. Unlike the old network submission, a
        // draft save is never one-shot: saving again after a successful save is completely ordinary
        // (the window stays open so more rows can be added), so the button simply re-enables either
        // way. A failed attempt needs it back immediately so the user can fix the problem and retry.
        // ###########################################################################################
        private void ApplySubmissionOutcome(bool saveSucceeded)
        {
            this.SubmitButton.IsEnabled = true;
        }

        // ###########################################################################################
        // The line shown once the draft has been saved. A new component is worth its own wording:
        // it now exists in the local draft and can be seen on the board straight away, unlike the
        // old submit-and-wait model - saying so is what stops this reading as a save that quietly
        // did nothing, the same way the old success text had to say a submission was queued.
        // ###########################################################################################
        private string BuildSaveSuccessText()
        {
            const string SuccessText = "Saved to your local draft.";

            if (this.thisIsDeleteComponent)
            {
                string deletedLabel = this.ResolveEffectiveBoardLabel();
                string deletedText = string.IsNullOrWhiteSpace(deletedLabel)
                    ? "The component"
                    : $"The component [{deletedLabel}]";

                return $"{SuccessText} {deletedText} is removed from your local view of this board until you submit and it is reviewed.";
            }

            if (!this.thisIsNewComponent)
            {
                return $"{SuccessText} Your changes are visible on this board right away.";
            }

            string boardLabel = this.ResolveEffectiveBoardLabel();
            string componentText = string.IsNullOrWhiteSpace(boardLabel)
                ? "The new component"
                : $"The new component [{boardLabel}]";

            return $"{SuccessText} {componentText} now appears on this board.";
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
        // happens on every row, not just the first, so one pass shows the maintainer every problem;
        // rows that are fine get their mark cleared here too. A row with a note and no file is fine,
        // as in the table (see ContributionPackaging.ValidateComponentImageFile) - so it needs the
        // row's note as well as its file.
        // ###########################################################################################
        private (ContributionComponentImageRow Row, string Message)? ValidateComponentImageRows()
        {
            (ContributionComponentImageRow Row, string Message)? firstProblem = null;

            for (int index = 0; index < this.thisComponentImageRows.Count; index++)
            {
                var row = this.thisComponentImageRows[index];
                var problem = ContributionPackaging.ValidateComponentImageFile(GetStoredFilePath(row), row.Note);

                row.FileErrorText = problem switch
                {
                    ContributionPackaging.ComponentImageFileProblem.NoFileOrNote => "No file or note",
                    ContributionPackaging.ComponentImageFileProblem.NotDisplayable => "Not an image the application can display",
                    _ => string.Empty
                };

                row.HasFileError = problem != ContributionPackaging.ComponentImageFileProblem.None;

                if (row.HasFileError && firstProblem == null)
                {
                    string rowLabel = $"Component image #{index + 1}";

                    string message = problem == ContributionPackaging.ComponentImageFileProblem.NoFileOrNote
                        ? $"{rowLabel} has neither a file nor a note - add one of them"
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
            // Through the general one (code review, 2026-09-26): this used to carry its own copy of
            // the expand-post-BringIntoView dance, so a later fix to the reveal timing would have
            // had to be made twice. The index is read BEFORE posting, as RevealRow's own callers do
            // - a row no longer in the list reveals nothing.
            int index = this.thisComponentImageRows.IndexOf(row);

            if (index >= 0)
                ComponentContributionWindow.RevealRow(this.ComponentImagesExpander, this.ComponentImageRowsItemsControl, index);
        }

        // ###########################################################################################
        // *** EVERY NEWLY PICKED FILE MUST BE FILED IN THIS BOARD'S FOLDERS OR A SHARED ONE (owner
        // report, 2026-09-25). *** A file picked from outside the data folder keeps the row's
        // "File location", which starts empty - and a row saved like that stored the bare file name
        // ("HotCPU.png"), outside every board. The draft took it; the server refused the submission
        // much later ("[HotCPU.png] belongs to another board"). The rule is
        // ContributionFileLocations.CheckNewFile; this marks every row it refuses (the "File
        // location" box, and an image row's own frame and text too) and returns the first, in
        // section order, with its message.
        // ###########################################################################################
        private (Action Reveal, string Message)? ValidateNewFileLocations()
        {
            string boardId = BoardDescriptorRules.BoardIdFromExcelDataFile(this.thisBoardExcelFile);
            (Action Reveal, string Message)? firstProblem = null;

            void CheckSection(IReadOnlyList<IContributionFileRow> rows, string rowName, Expander section, ItemsControl list)
            {
                for (int index = 0; index < rows.Count; index++)
                {
                    IContributionFileRow row = rows[index];
                    ContributionFileLocations.NewFileProblem problem = this.CheckNewFileLocation(boardId, row);
                    bool noFolder = problem == ContributionFileLocations.NewFileProblem.NoFolder;

                    row.HasLocationError = problem != ContributionFileLocations.NewFileProblem.None;

                    if (!row.HasLocationError)
                    {
                        continue;
                    }

                    if (row is ContributionComponentImageRow imageRow)
                    {
                        imageRow.HasFileError = true;
                        imageRow.FileErrorText = noFolder
                            ? "Choose a file location"
                            : "Choose one of this board's folders or a shared folder";
                    }

                    if (firstProblem == null)
                    {
                        int rowIndex = index;
                        string rowLabel = $"{rowName} #{index + 1} [{row.File}]";

                        string message = noFolder
                            ? $"{rowLabel} has no file location - choose the folder it goes in: one of this board's folders, or a shared folder"
                            : $"{rowLabel} is filed in [{row.FileLocation}], which is not one of this board's folders or a shared folder - choose one of those";

                        firstProblem = (() => ComponentContributionWindow.RevealRow(section, list, rowIndex), message);
                    }
                }
            }

            CheckSection(
                this.thisComponentImageRows.Cast<IContributionFileRow>().ToList(),
                "Component image",
                this.ComponentImagesExpander,
                this.ComponentImageRowsItemsControl);

            CheckSection(
                this.thisComponentLocalFileRows.Cast<IContributionFileRow>().ToList(),
                "Component file",
                this.ComponentLocalFilesExpander,
                this.ComponentLocalFileRowsItemsControl);

            CheckSection(
                this.thisBoardLocalFileRows.Cast<IContributionFileRow>().ToList(),
                "Board file",
                this.BoardLocalFilesExpander,
                this.BoardLocalFileRowsItemsControl);

            return firstProblem;
        }

        // ###########################################################################################
        // One row's answer. Only a file picked HERE is checked - its source is a full path. A row
        // loaded from the board keeps the path it is stored under, and this window does not move a
        // stored file, so there is nothing for its location to decide.
        //
        // A file picked from the data folder and left in the folder it came from is "used in
        // place": it IS the published copy, so citing it is fine even from another board's folder.
        // Compared exactly - a different capitalisation would name a different file on the server.
        // ###########################################################################################
        private ContributionFileLocations.NewFileProblem CheckNewFileLocation(string boardId, IContributionFileRow row)
        {
            if (string.IsNullOrWhiteSpace(row.File) ||
                string.IsNullOrWhiteSpace(row.OriginalFilePath) ||
                !Path.IsPathRooted(row.OriginalFilePath))
            {
                return ContributionFileLocations.NewFileProblem.None;
            }

            string location = row.FileLocation?.Trim().Replace('\\', '/').Trim('/') ?? string.Empty;

            string? pickedFrom = ContributionPackaging.TryGetDataRootRelativeFolder(
                this.thisDataRoot,
                Path.GetDirectoryName(row.OriginalFilePath));

            bool usedInPlace = !string.IsNullOrEmpty(pickedFrom) &&
                               string.Equals(pickedFrom, location, StringComparison.Ordinal);

            return ContributionFileLocations.CheckNewFile(boardId, location, usedInPlace);
        }

        // ###########################################################################################
        // Opens a section and brings one of its rows on screen - RevealComponentImageRow's reasoning,
        // for any section: the row's container does not exist until the expander has laid it out.
        // ###########################################################################################
        private static void RevealRow(Expander section, ItemsControl list, int index)
        {
            section.IsExpanded = true;

            Dispatcher.UIThread.Post(() =>
            {
                if (list.ContainerFromIndex(index) is Control container)
                {
                    container.BringIntoView();
                }
            }, DispatcherPriority.Background);
        }

        // ###########################################################################################
        // Resolves an edited file path so it can be verified for existence and previewed.
        // Accepts both relative paths (resolved against data-root) and external absolute paths.
        // ###########################################################################################
        private string? ResolveExistingFilePath(string pathValue)
        {
            return ContributionPackaging.ResolveExistingFilePath(this.thisDataRoot, pathValue);
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

            // Containment is decided by ContributionPackaging.TryGetDataRootRelativeFolder (pure,
            // unit tested) rather than by a prefix test here - see its header for the sibling-folder
            // case a bare StartsWith let through, and why a traversing FileLocation must never reach
            // the uploaded payload.
            //
            // null means the file is OUTSIDE the data root, and FileLocation is then deliberately
            // left alone: the file keeps whatever drop-down folder the user selected.
            string? relativeFolder = ContributionPackaging.TryGetDataRootRelativeFolder(
                this.thisDataRoot,
                Path.GetDirectoryName(selectedPath));

            if (relativeFolder != null)
            {
                row.FileLocation = relativeFolder;
            }

            this.SetAvailableFileLocations(row);

            // The folder list was just rebuilt, so the bound drop-down is told to read the
            // unchanged value again - see IContributionFileRow.NotifyFileLocationChanged for why
            // this is not an assignment.
            row.NotifyFileLocationChanged();
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