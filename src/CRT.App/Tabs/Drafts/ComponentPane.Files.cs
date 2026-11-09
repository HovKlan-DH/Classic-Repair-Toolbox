using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // THE COMPONENT PANE'S OSCILLOSCOPE BASELINES, PICTURES AND FILES (CRT 4.0 step 2, slice 3b, 2026-10-26
    // - the plan: "Scope baselines: pin; a screenshot, which is required; Capture from my oscilloscope
    // appears when a scope is connected; T/DIV and V/DIV as drop-downs; trigger level, expected reading,
    // region and machine state; where it comes from", and "No 'File location' box. Every file belongs
    // to its board"). See ComponentPane.axaml.cs for the file map.
    //
    //   - A FILE ADDED GOES INTO THE DRAFT AT ONCE: copied in under a name of the board's own
    //     (ComponentPaneFileNames - never one another file has, which it would replace for every row
    //     citing it), the rows changed as one step, and the draft SAVED straight away, so its bytes are
    //     in the draft's store and listed. Left lying in the folder until the save a second later, an
    //     undo in between would leave a file no row cites - work, to the draft's retirement.
    //   - AN OSCILLOSCOPE BASELINE IS ONE ONLY WITH ITS PIN (ComponentImageKinds - else every reader takes
    //     it for a picture), and needs its signal name and an image; T/DIV, V/DIV and the trigger level may
    //     stay empty (owner request, 2026-11-02). A new one waits in its own panel until it has all three;
    //     an edit of one that would take them away is held back - the draft keeps it as it was - with the
    //     reason under it (ComponentPaneWording.BaselineProblem).
    //   - A NEW ONE'S SIGNAL NAME COMES WITH ITS PIN (owner request, 2026-11-02: "when adding a new pin and
    //     I give it a number that exists ... it should per default copy the Signal name"): from the
    //     component's pinout, else the board's KiCad data, until somebody types one (FollowNewBaselinePin).
    //   - "CAPTURE FROM MY OSCILLOSCOPE" while the Oscilloscope tab has a session (ShowScopeConnected):
    //     the scope's screen as PNG, from the host (TabOscilloscope.CaptureOscilloscopePngAsync).
    //   - THUMBNAILS are decoded small and off the UI thread, and never disposed by hand: one still drawn
    //     would take the render thread down (the Workbooks tab's lesson) - the garbage collector takes them.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // What the file picker is asked for - and so which files it offers.
        internal enum PaneFileKind
        {
            Picture,
            Screenshot,
            File,
            BoardFile
        }

        // Twice the width a thumbnail is drawn at, for a sharp one on a high-density screen.
        private const int ThumbnailWidth = 336;

        // Internal: the Workbooks tab's device photos pick the same kinds of picture.
        internal static readonly FilePickerFileType PictureFileType = new("Pictures")
        {
            Patterns = ContributionPackaging.DisplayableImageExtensions.Select(extension => "*" + extension).ToArray(),
            MimeTypes = ["image/*"],
            AppleUniformTypeIdentifiers = ["public.image"],
        };

        private readonly ObservableCollection<ComponentPaneBaselineRow> thisBaselines = new();
        private readonly ObservableCollection<ComponentPanePictureRow> thisPictures = new();
        private readonly ObservableCollection<ComponentPaneFileRow> thisFiles = new();
        private ListRowDrag<ComponentPaneFileRow>? thisFilesDrag;

        // Each file's thumbnail by its name in the board - read once while the component is open.
        // Every picture read while the pane is open - thumbnails and the Pictures part's previews alike.
        private readonly PaneImageCache thisImages = new();

        // The new baseline's screenshot until it is added: a file picked, or a capture in the temp folder.
        private string? thisNewBaselineSource;

        // Where the new baseline's screenshot waits until it is added - for tests.
        internal string? NewBaselineSourceForTests => this.thisNewBaselineSource;
        private bool thisNewBaselineIsCapture;

        private bool thisScopeConnected;

        // The scope's screen as PNG bytes, or null (the reason in the Oscilloscope tab's output) - the
        // host's, TabOscilloscope.CaptureOscilloscopePngAsync. Null: no capture is offered.
        internal Func<CancellationToken, Task<byte[]?>>? CaptureScopeImage { get; set; }

        // Stands in for the file picker in a test: the kind asked for, the path picked - null, cancelled.
        internal Func<PaneFileKind, Task<string?>>? PickFileOverrideForTests { get; set; }

        internal IReadOnlyList<ComponentPaneBaselineRow> BaselinesForTests => this.thisBaselines;

        internal IReadOnlyList<ComponentPanePictureRow> PicturesForTests => this.thisPictures;

        internal IReadOnlyList<ComponentPaneFileRow> FilesForTests => this.thisFiles;

        // Every thumbnail asked for so far, read - for a test.
        internal Task ThumbnailsLoadedForTests => this.PreviewsLoadedForTests;

        // Awaited once a picture is decoded, before it is kept - a test holds a load there.
        internal Func<string, Task>? ImageDecodedForTests { get; set; }

        private void InitializeImagesAndFiles()
        {
            this.BaselinesItemsControl.ItemsSource = this.thisBaselines;
            this.PicturesItemsControl.ItemsSource = this.thisPictures;
            this.FilesItemsControl.ItemsSource = this.thisFiles;
            this.thisFilesDrag = new ListRowDrag<ComponentPaneFileRow>(
                this, this.FilesItemsControl, this.SectionsScrollViewer, this.thisFiles, (_, _, _) => this.CommitFiles());

            this.NewBaselineTimeDivBox.ItemsSource = ComponentVocabulary.TimeDivisions;
            this.NewBaselineVoltsDivBox.ItemsSource = ComponentVocabulary.VoltDivisions;

            // What the new baseline still needs follows its boxes as they are typed in.
            this.NewBaselinePanel.AddHandler(
                TextBox.TextChangedEvent,
                (_, e) =>
                {
                    if (ReferenceEquals(e.Source, this.NewBaselinePinBox))
                    {
                        this.FollowNewBaselinePin();
                    }

                    this.ShowNewBaselineState();
                },
                handledEventsToo: true);
        }

        // ---- showing ----------------------------------------------------------------------------

        private void ShowImagesAndFiles(BoardData board)
        {
            ComponentPane.Fill(this.thisBaselines, ComponentPaneEdits.BaselinesOf(board, this.thisLabel, this.thisRegion).Select(ComponentPaneBaselineRow.From));
            ComponentPane.Fill(this.thisPictures, ComponentPaneEdits.PicturesOf(board, this.thisLabel, this.thisRegion).Select(ComponentPanePictureRow.From));
            ComponentPane.Fill(this.thisFiles, ComponentPaneEdits.AllFilesOf(board, this.thisLabel).Select(ComponentPaneFileRow.From));

            this.ShowMeasuringNotes(board);
            this.LoadThumbnails();
            this.ShowCaptureOffer();
        }

        // ###########################################################################################
        // How the board's baselines were measured - its "Scope measurements" files, each a link (owner
        // decision, 2026-10-26: the files the data already keeps, no new field) - or how to write it
        // down when the board has none.
        // ###########################################################################################
        private void ShowMeasuringNotes(BoardData board)
        {
            List<BoardLocalFileEntry> notes = ScopeMeasurementFiles.Of(board, this.thisRegion);

            this.MeasuringNotesItemsControl.ItemsSource = notes;
            this.MeasuringNotesItemsControl.IsVisible = notes.Count > 0;
            this.MeasuringNotesIntro.Text = ComponentPaneWording.MeasuringNotesLine(
                notes.Count > 0,
                ScopeMeasurementFiles.Of(board).FirstOrDefault()?.Region);
        }

        // For tests, in place of opening a file with the program the system has for it.
        internal Action<string, string?>? OpenFileOverrideForTests { get; set; }

        // A "Scope measurements" file opened.
        private void OnMeasuringNoteClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: BoardLocalFileEntry note })
            {
                this.OpenFileOfBoard(note.File, note.Name);
            }
        }

        // ###########################################################################################
        // A file the board names opened with the program the system has for it - found as every file of
        // the draft is (DraftFileResolver: a new file lying in the draft folder, the draft's store, the
        // published board's), and handed to the launcher with the root it was found under. The Files
        // part's "Open" too (2026-10-29). `name` is what the person reads, said when it cannot be.
        // ###########################################################################################
        internal void OpenFileOfBoard(string? file, string? name)
        {
            this.ShowMessage(null);

            if (this.thisSession is not { } session)
            {
                return;
            }

            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);

            if (DraftFileResolver.ResolveWithSource(DataManager.DataRoot, draftFolder, file, session.ExcelDataFile) is not { } found)
            {
                this.ShowMessage(ComponentPaneWording.FileNotFound(name));
                return;
            }

            if (this.OpenFileOverrideForTests is { } open)
            {
                open(found.FullPath, found.Root);
                return;
            }

            if (!ExternalTargetLauncher.TryOpen(found.FullPath, found.Root))
            {
                this.ShowMessage(ComponentPaneWording.FileNotOpened(name));
            }
        }

        // For tests, in place of showing a file in the file manager.
        internal Action<string, string?>? RevealFileOverrideForTests { get; set; }

        // ###########################################################################################
        // "REVEAL" (owner request, 2026-11-07: 'a button named "Reveal" and then open and select the file in File
        // Explorer'): a file the board names shown selected in the system's file manager - found as "Open file" finds it,
        // and held to the root it was found under (ExternalTargetLauncher.TryReveal).
        // ###########################################################################################
        internal void RevealFileOfBoard(string? file, string? name)
        {
            this.ShowMessage(null);

            if (this.thisSession is not { } session)
            {
                return;
            }

            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);

            if (DraftFileResolver.ResolveWithSource(DataManager.DataRoot, draftFolder, file, session.ExcelDataFile) is not { } found)
            {
                this.ShowMessage(ComponentPaneWording.FileNotFound(name));
                return;
            }

            if (this.RevealFileOverrideForTests is { } reveal)
            {
                reveal(found.FullPath, found.Root);
                return;
            }

            if (!ExternalTargetLauncher.TryReveal(found.FullPath, found.Root))
            {
                this.ShowMessage(ComponentPaneWording.FileNotRevealed(name));
            }
        }

        // "Reveal" on any row naming a file: a component's file, picture or baseline, or a board file.
        private void OnRevealFileClick(object? sender, RoutedEventArgs e)
        {
            object? tag = (sender as Control)?.Tag;

            if (ComponentPane.FileAndNameOf(tag) is (string file, string name))
            {
                this.RevealFileOfBoard(file, string.IsNullOrWhiteSpace(name) ? file : name);
            }
        }

        // The file a row names, and its name - for any row "Reveal" stands on; null for anything else.
        private static (string File, string Name)? FileAndNameOf(object? row)
        {
            if (row is ComponentPaneFileRow file)
            {
                return (file.File, file.Name);
            }

            if (row is ComponentPanePictureRow picture)
            {
                return (picture.File, picture.Name);
            }

            if (row is ComponentPaneBaselineRow baseline)
            {
                return (baseline.File, baseline.Name);
            }

            return row is ComponentPaneBoardFileRow boardFile ? (boardFile.File, boardFile.Name) : null;
        }

        // "Open file" on a component's file (owner request, 2026-11-07: "it misses the same "Open file" button, which is
        // available in the "Board: Files"") - the board file's own way (OpenFileOfBoard).
        private void OnOpenComponentFileClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneFileRow row })
            {
                this.OpenFileOfBoard(row.File, row.Name.Trim().Length > 0 ? row.Name : row.File);
            }
        }

        // How many each holds is its tab's to say (ShowPageTabs).
        private void ShowImageAndFileHeadings()
        {
            this.NoBaselinesText.IsVisible = this.thisBaselines.Count == 0;
            this.NoPicturesText.IsVisible = this.thisPictures.Count == 0;
            this.NoFilesText.IsVisible = this.thisFiles.Count == 0;
            this.FilesHeading.IsVisible = this.thisFiles.Count > 0;

            // A picture moves one place at a time, as a schematic image does (owner request, 2026-11-02).
            for (int index = 0; index < this.thisPictures.Count; index++)
            {
                // An order is each component's own - none moves while several are shown (RowsMove).
                this.thisPictures[index].CanMoveUp = this.RowsMove && index > 0;
                this.thisPictures[index].CanMoveDown = this.RowsMove && index < this.thisPictures.Count - 1;
            }
        }

        private void ForgetImagesAndFiles()
        {
            this.thisBaselines.Clear();
            this.thisPictures.Clear();
            this.thisFiles.Clear();
            this.thisImages.Clear();
            this.LetGoOfWhatIsOfTheComponentShown();
        }

        // ###########################################################################################
        // What belongs to the component shown and to no other, gone as the pane moves on (code review,
        // 2026-11-02). The page stays as another component is picked, on purpose - and a half-filled new
        // baseline on it stayed too, whose Add then filed U8's pin, signal and picture under U9. A capture
        // waiting in that panel is deleted with it (ForgetNewBaselineSource). And what the pane said of
        // it: "Moved to NTSC" stood under the title of the next component.
        // ###########################################################################################
        private void LetGoOfWhatIsOfTheComponentShown()
        {
            this.CloseNewBaseline();
            this.ShowMessage(null);
        }

        // What an action could not do, under the title - null clears it.
        private void ShowMessage(string? message)
        {
            this.PaneMessageText.Text = message ?? string.Empty;
            this.PaneMessageText.IsVisible = !string.IsNullOrEmpty(message);
        }

        // ###########################################################################################
        // "Capture from my oscilloscope" while the Oscilloscope tab has a session - from the host, as
        // the tab's state moves (Main.UpdateComponentInfoWindowsOscilloscopeSessionState).
        // ###########################################################################################
        internal void ShowScopeConnected(bool connected)
        {
            this.thisScopeConnected = connected;
            this.ShowCaptureOffer();
        }

        private void ShowCaptureOffer()
        {
            bool offered = this.thisScopeConnected && this.CaptureScopeImage is not null;

            this.CaptureBaselineButton.IsVisible = offered;
            this.NewBaselineCaptureButton.IsVisible = offered;
        }

        // ---- into the draft ---------------------------------------------------------------------

        // ###########################################################################################
        // The baselines as the pane holds them, into the draft as one step - each one that is still a
        // baseline as it reads now, and the one that would not be (its pin, its signal name or its image
        // cleared) as the draft has it, the reason said under it.
        // ###########################################################################################
        private void CommitBaselines()
        {
            // A baseline is measured on one component: none is shown, nor written, while several are (ComponentPane.Several.cs).
            if (this.IsSeveral)
            {
                return;
            }

            var rows = new List<ComponentImageEntry>(this.thisBaselines.Count);
            var taken = new List<(ComponentPaneBaselineRow Row, ComponentImageEntry Entry)>();

            foreach (ComponentPaneBaselineRow row in this.thisBaselines)
            {
                ComponentImageEntry wanted = row.ToEntry();
                string? problem = ComponentPaneWording.BaselineProblem(wanted.Pin, wanted.Name, wanted.File);

                row.Problem = problem is null ? string.Empty : ComponentPaneWording.BaselineHeldBack(problem);
                rows.Add(problem is null ? wanted : row.Committed);

                if (problem is null)
                {
                    taken.Add((row, wanted));
                }
            }

            this.Apply(
                DraftStepWording.ComponentBaselines(this.thisLabel),
                board => ComponentPaneEdits.WithBaselines(board, this.thisLabel, this.thisRegion, rows));

            foreach ((ComponentPaneBaselineRow row, ComponentImageEntry entry) in taken)
            {
                row.Committed = entry;
            }
        }

        private void CommitPictures()
        {
            if (this.IsSeveral)
            {
                this.CommitPicturesOfSeveral();
                return;
            }

            List<ComponentImageEntry> pictures = this.thisPictures.Select(row => row.ToEntry()).ToList();

            this.Apply(
                DraftStepWording.ComponentPictures(this.thisLabel),
                board => ComponentPaneEdits.WithPictures(board, this.thisLabel, this.thisRegion, pictures));
        }

        private void CommitFiles()
        {
            if (this.IsSeveral)
            {
                this.CommitFilesOfSeveral();
                return;
            }

            List<ComponentLocalFileEntry> files = this.thisFiles.Select(row => row.ToEntry()).ToList();

            this.Apply(
                DraftStepWording.ComponentFiles(this.thisLabel),
                board => ComponentPaneEdits.WithAllFiles(board, this.thisLabel, files));
        }

        // ###########################################################################################
        // A picked file into the draft under `name`, then `change` (the rows citing it, as one step),
        // then the draft saved - all under CRT's "please wait". False, with the reason under the title,
        // when the file could not be copied in; a save that fails says so in the draft's header, as
        // every save of the pane does.
        // ###########################################################################################
        private Task<bool> AddToDraftAsync(string source, string name, Action change) =>
            this.AddToDraftAsync([(source, name)], _ => change());

        // ###########################################################################################
        // Several picked files at once - the board's pictures - each under its name, one step for all.
        // `change` is handed the names actually COPIED: a file that could not be (locked by another
        // program, gone since it was picked) gets no row (code review, 2026-10-28).
        // ###########################################################################################
        private async Task<bool> AddToDraftAsync(IReadOnlyList<(string Source, string Name)> files, Action<IReadOnlyList<string>> change)
        {
            if (this.thisSession is not { } session)
            {
                return false;
            }

            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);

            return await BusyOverlay.RunLocalAsync(this, CrtWaitWording.AddingFile, async () =>
            {
                DraftPickedFiles copies = await Task.Run(() => DraftPickedFiles.CopyInto(draftFolder, files));

                if (!copies.AnyLying)
                {
                    // Said for what it is: a file of no bytes is not added at all (2026-11-02).
                    this.ShowMessage(copies.EmptyFiles.Count > 0 ? ComponentPaneWording.FileEmpty : ComponentPaneWording.FileNotCopied);
                    return false;
                }

                List<string> copied = copies.Copied.ToList();
                copies.Keep();

                this.thisFilesArriving = true;

                try
                {
                    change(copied);
                    this.ShowEmptyLines();
                    await this.SaveNowAsync();
                }
                finally
                {
                    this.thisFilesArriving = false;
                }

                this.LoadThumbnails();
                this.LoadPreviews();
                return true;
            });
        }

        // ###########################################################################################
        // *** NO PICTURE IS READ WHILE THE FILES JUST COPIED IN ARE ON THEIR WAY INTO THE DRAFT (owner
        // report, 2026-11-01: a new board's pictures showed as grey boxes until the part was opened
        // again). *** A copy lies in the draft folder until the save lists it and moves it into the
        // draft's store (DraftFiles); a preview read meanwhile found it there, and the save took it away
        // before it was decoded - read as a picture that does not decode, never asked for again. So the
        // thumbnails and previews are read once that save is done (AddToDraftAsync).
        // ###########################################################################################
        private bool thisFilesArriving;

        // The names the board's files have, or an undo can bring back - none of them is given to a file
        // added here (DraftSession.FileNamesInUse).
        private List<string> TakenNames() =>
            this.thisSession is not { } session
                ? []
                : session.FileNamesInUse.ToList();

        private async Task<string?> PickFileAsync(PaneFileKind kind)
        {
            if (this.PickFileOverrideForTests is { } pick)
            {
                return await pick(kind);
            }

            if (TopLevel.GetTopLevel(this)?.StorageProvider is not { } storage)
            {
                return null;
            }

            IReadOnlyList<IStorageFile> files = await storage.OpenFilePickerAsync(new FilePickerOpenOptions
            {
                Title = kind switch
                {
                    PaneFileKind.Picture => "Pick a picture of the component",
                    PaneFileKind.Screenshot => "Pick the screenshot from the oscilloscope",
                    PaneFileKind.BoardFile => "Pick a file about the board",
                    _ => "Pick a file about the component",
                },
                AllowMultiple = false,
                FileTypeFilter = kind is PaneFileKind.File or PaneFileKind.BoardFile ? null : [ComponentPane.PictureFileType],
            });

            return files.Count == 0 ? null : files[0].TryGetLocalPath();
        }

        // A picked picture or screenshot CRT can show - the picker's filter is only a suggestion.
        private bool IsPicture(string path)
        {
            if (ContributionPackaging.IsDisplayableImageFile(path))
            {
                return true;
            }

            this.ShowMessage(ComponentPaneWording.NotAnImage);
            return false;
        }

        // ---- pictures ---------------------------------------------------------------------------

        private async void OnAddPictureClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AddPictureAsync();
            }
            catch (Exception ex)
            {
                // Nothing may leave an async void handler: App lets such an exception end CRT.
                Logger.Critical($"Adding a picture in the component pane failed - [{ex}]");
            }
        }

        internal async Task AddPictureAsync()
        {
            this.CommitHeld();
            this.ShowMessage(null);

            if (await this.PickFileAsync(PaneFileKind.Picture) is not { } path || !this.IsPicture(path))
            {
                return;
            }

            string name = ComponentPaneFileNames.ForPicture(path, this.TakenNames());

            await this.AddToDraftAsync(path, name, () =>
            {
                // For every region, as nearly all of them are - its "Applies to" narrows it.
                this.thisPictures.Add(new ComponentPanePictureRow
                {
                    Name = Path.GetFileNameWithoutExtension(path),
                    File = name,
                });

                this.CommitPictures();
            });
        }

        private async void OnChangePictureClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Control { Tag: ComponentPanePictureRow row })
                {
                    await this.ChangePictureAsync(row);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Changing a picture in the component pane failed - [{ex}]");
            }
        }

        internal async Task ChangePictureAsync(ComponentPanePictureRow row)
        {
            this.CommitHeld();
            this.ShowMessage(null);

            if (await this.PickFileAsync(PaneFileKind.Picture) is not { } path || !this.IsPicture(path))
            {
                return;
            }

            string name = ComponentPaneFileNames.ForPicture(path, this.TakenNames());

            await this.AddToDraftAsync(path, name, () =>
            {
                row.File = name;
                row.Thumbnail = null;
                this.CommitPictures();
            });
        }

        private void OnRemovePictureClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPanePictureRow row } && this.thisPictures.Remove(row))
            {
                this.CommitPictures();
                this.ShowEmptyLines();
            }
        }

        private void OnMoveComponentPictureUpClick(object? sender, RoutedEventArgs e) => this.MoveComponentPicture(sender, -1);

        private void OnMoveComponentPictureDownClick(object? sender, RoutedEventArgs e) => this.MoveComponentPicture(sender, +1);

        private void MoveComponentPicture(object? sender, int by)
        {
            if (sender is Control { Tag: ComponentPanePictureRow row })
            {
                this.MoveComponentPicture(row, by);
            }
        }

        // One place up (-1) or down (+1) among the pictures shown, ONE step - what the pane holds handed
        // over first. The row stays the same object: its thumbnail stays with it.
        internal bool MoveComponentPicture(ComponentPanePictureRow row, int by)
        {
            if (!this.MoveRow(this.thisPictures, row, by))
            {
                return false;
            }

            this.CommitPictures();
            this.ShowEmptyLines();
            return true;
        }

        // ---- files ------------------------------------------------------------------------------

        private async void OnAddFileClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AddFileAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Adding a file in the component pane failed - [{ex}]");
            }
        }

        internal async Task AddFileAsync()
        {
            this.CommitHeld();
            this.ShowMessage(null);

            if (await this.PickFileAsync(PaneFileKind.File) is not { } path)
            {
                return;
            }

            string name = ComponentPaneFileNames.ForFile(path, this.TakenNames());

            await this.AddToDraftAsync(path, name, () =>
            {
                this.thisFiles.Add(new ComponentPaneFileRow { Name = Path.GetFileNameWithoutExtension(path), File = name });
                this.CommitFiles();
            });
        }

        private async void OnChangeFileClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Control { Tag: ComponentPaneFileRow row })
                {
                    await this.ChangeFileAsync(row);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Changing a file in the component pane failed - [{ex}]");
            }
        }

        internal async Task ChangeFileAsync(ComponentPaneFileRow row)
        {
            this.CommitHeld();
            this.ShowMessage(null);

            if (await this.PickFileAsync(PaneFileKind.File) is not { } path)
            {
                return;
            }

            string name = ComponentPaneFileNames.ForFile(path, this.TakenNames());

            await this.AddToDraftAsync(path, name, () =>
            {
                row.File = name;
                this.CommitFiles();
            });
        }

        private void OnRemoveFileClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneFileRow row } && this.thisFiles.Remove(row))
            {
                this.CommitFiles();
                this.ShowEmptyLines();
            }
        }

        // A file dragged by its handle, or moved by Alt+Up and Alt+Down, as the board's are (owner request,
        // 2026-11-02: "the reordering is missing in the component part").
        private void OnFileHandlePressed(object? sender, Avalonia.Input.PointerPressedEventArgs e)
        {
            // An order is each component's own - none moves while several are shown (RowsMove).
            if (this.RowsMove)
            {
                this.thisFilesDrag?.Press(sender, e);
            }
        }

        internal bool MoveFile(ComponentPaneFileRow row, int by)
        {
            if (!this.MoveRow(this.thisFiles, row, by))
            {
                return false;
            }

            this.CommitFiles();
            return true;
        }

        // ---- oscilloscope baselines -------------------------------------------------------------

        private async void OnChangeBaselineScreenshotClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                if (sender is Control { Tag: ComponentPaneBaselineRow row })
                {
                    await this.ChangeBaselineScreenshotAsync(row);
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Changing an oscilloscope baseline's image in the component pane failed - [{ex}]");
            }
        }

        internal async Task ChangeBaselineScreenshotAsync(ComponentPaneBaselineRow row)
        {
            this.CommitHeld();
            this.ShowMessage(null);

            if (await this.PickFileAsync(PaneFileKind.Screenshot) is not { } path || !this.IsPicture(path))
            {
                return;
            }

            string name = ComponentPaneFileNames.ForBaseline(this.thisLabel, row.Committed.Pin, row.Committed.Name, path, this.TakenNames());

            await this.AddToDraftAsync(path, name, () =>
            {
                row.File = name;
                row.Thumbnail = null;
                this.CommitBaselines();
            });
        }

        private void OnRemoveBaselineClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneBaselineRow row } && this.thisBaselines.Remove(row))
            {
                this.CommitBaselines();
                this.ShowEmptyLines();
            }
        }

        // ---- a new oscilloscope baseline --------------------------------------------------------

        private void OnAddBaselineClick(object? sender, RoutedEventArgs e) => this.OpenNewBaseline();

        // ###########################################################################################
        // The panel a new baseline waits in. Its scope settings start as the component's last
        // baseline's - most of a chip's pins are measured alike.
        // ###########################################################################################
        internal void OpenNewBaseline()
        {
            this.CommitHeld();
            this.ShowMessage(null);
            this.ForgetNewBaselineSource();

            ComponentPaneBaselineRow? last = this.thisBaselines.LastOrDefault();

            this.NewBaselinePinBox.Text = string.Empty;
            this.NewBaselineNameBox.Text = string.Empty;
            this.thisNewBaselineNameOffered = string.Empty;
            this.NewBaselineTimeDivBox.Text = last?.TimeDiv ?? string.Empty;
            this.NewBaselineVoltsDivBox.Text = last?.VoltsDiv ?? string.Empty;
            this.NewBaselineTriggerBox.Text = last?.TriggerLevel ?? string.Empty;
            this.NewBaselineThumbnail.Source = null;

            // Its region: the one shown, as nearly all baselines are of one region (owner decision,
            // 2026-11-01) - on a board that has regions.
            string region = this.RegionForNewBaseline();
            this.NewBaselineRegionPanel.IsVisible = this.thisBoardHasRegions;
            this.NewBaselineRegionBox.ItemsSource = ComponentPaneWording.AppliesToChoices(this.thisBoardRegions, region);
            this.NewBaselineRegionBox.SelectedItem = ComponentPaneWording.AppliesToOf(this.thisBoardRegions, region);

            this.NewBaselinePanel.IsVisible = true;
            this.AddBaselineButtons.IsVisible = false;
            this.ShowNewBaselineState();
            this.NewBaselinePinBox.Focus();
        }

        private void OnCancelNewBaselineClick(object? sender, RoutedEventArgs e) => this.CloseNewBaseline();

        private void CloseNewBaseline()
        {
            this.ForgetNewBaselineSource();
            this.NewBaselineThumbnail.Source = null;
            this.NewBaselinePanel.IsVisible = false;
            this.AddBaselineButtons.IsVisible = true;
        }

        // A capture written to the temp folder is gone once the baseline is added or let go of.
        private void ForgetNewBaselineSource()
        {
            if (this.thisNewBaselineIsCapture && this.thisNewBaselineSource is { } capture)
            {
                try
                {
                    File.Delete(capture);
                }
                catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
                {
                    Logger.Warning($"Could not remove the scope capture [{capture}] - [{ex.Message}]");
                }
            }

            this.thisNewBaselineSource = null;
            this.thisNewBaselineIsCapture = false;
        }

        // ###########################################################################################
        // The signal name that goes with the pin typed - the pinout's, else the board's KiCad data's
        // (KiCadComponentPins.NameOfPin) - while the box is empty or still holds what was offered for the
        // pin before: a name somebody typed is never written over. A pin with none takes back what was
        // offered, so a slip of the keys leaves no name behind.
        // ###########################################################################################
        private void FollowNewBaselinePin()
        {
            string shown = this.NewBaselineNameBox.Text ?? string.Empty;

            if (shown.Length > 0 && !string.Equals(shown, this.thisNewBaselineNameOffered, StringComparison.Ordinal))
            {
                return;
            }

            string offered = KiCadComponentPins.NameOfPin(
                this.thisPins.Select(row => row.ToEntry()).ToList(),
                this.thisKiCadPins,
                this.NewBaselinePinBox.Text);

            this.thisNewBaselineNameOffered = offered;
            this.NewBaselineNameBox.Text = offered;
        }

        // What the new baseline's name box was last given (FollowNewBaselinePin).
        private string thisNewBaselineNameOffered = string.Empty;

        // What the new baseline still needs, and Add only once it has nothing more to need.
        private void ShowNewBaselineState()
        {
            string? problem = ComponentPaneWording.BaselineProblem(
                this.NewBaselinePinBox.Text,
                this.NewBaselineNameBox.Text,
                this.thisNewBaselineSource) ?? this.NewBaselinePinProblem();

            this.NewBaselineProblemText.Text = problem ?? string.Empty;
            this.NewBaselineProblemText.IsVisible = problem is not null;
            this.AddNewBaselineButton.IsEnabled = problem is null;
        }

        // The region the new baseline's drop-down list says - blank for every region.
        private string NewBaselineRegion() =>
            this.NewBaselineRegionPanel.IsVisible
                ? ComponentPaneWording.RegionOfAppliesTo(this.NewBaselineRegionBox.SelectedItem as string)
                : string.Empty;

        // The new baseline's pin, when the component's case has none such (owner request, 2026-11-01).
        private string? NewBaselinePinProblem() =>
            this.thisSession is { } session &&
            BoardDataChecks.PackageRefusing(session.Content.Board.Components, this.thisLabel, this.NewBaselineRegion(), this.NewBaselinePinBox.Text) is { } package
                ? BoardDataChecks.PinNotOnPackage(this.NewBaselinePinBox.Text, package)
                : null;

        private async void OnChooseNewBaselineScreenshotClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.ChooseNewBaselineScreenshotAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Choosing a new oscilloscope baseline's image failed - [{ex}]");
            }
        }

        internal async Task ChooseNewBaselineScreenshotAsync()
        {
            this.ShowMessage(null);

            if (await this.PickFileAsync(PaneFileKind.Screenshot) is not { } path || !this.IsPicture(path))
            {
                return;
            }

            this.UseNewBaselineSource(path, isCapture: false);
        }

        private async void OnCaptureBaselineClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                this.OpenNewBaseline();
                await this.CaptureNewBaselineAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Capturing a new oscilloscope baseline failed - [{ex}]");
            }
        }

        private async void OnCaptureNewBaselineClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.CaptureNewBaselineAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Capturing a new oscilloscope baseline failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The oscilloscope's screen as the new baseline's screenshot - through the host, under CRT's
        // "please wait", written to the user's own folder until the baseline is added
        // (AppConfig.ScopeCaptureWaitFolder - never the shared temporary one).
        // ###########################################################################################
        internal async Task CaptureNewBaselineAsync()
        {
            if (this.CaptureScopeImage is not { } capture)
            {
                return;
            }

            this.ShowMessage(null);

            byte[]? png = await BusyOverlay.RunLocalAsync(this, CrtWaitWording.CapturingOscilloscope, () => capture(CancellationToken.None));

            if (png is not { Length: > 0 })
            {
                this.ShowMessage(ComponentPaneWording.CaptureFailed);
                return;
            }

            string path = Path.Combine(AppConfig.ScopeCaptureWaitFolder, $"{Guid.NewGuid():N}.png");

            try
            {
                Directory.CreateDirectory(Path.GetDirectoryName(path)!);
                await File.WriteAllBytesAsync(path, png);
            }
            catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
            {
                Logger.Warning($"Could not keep the scope capture at [{path}] - [{ex.Message}]");
                this.ShowMessage(ComponentPaneWording.CaptureFailed);
                return;
            }

            this.UseNewBaselineSource(path, isCapture: true);
        }

        private void UseNewBaselineSource(string path, bool isCapture)
        {
            this.ForgetNewBaselineSource();
            this.thisNewBaselineSource = path;
            this.thisNewBaselineIsCapture = isCapture;
            this.ShowNewBaselineState();

            this.TrackPreviewLoad(this.LoadNewBaselineThumbnailAsync(path));
        }

        private async Task LoadNewBaselineThumbnailAsync(string path)
        {
            Bitmap? bitmap = await Task.Run(() => ComponentPane.DecodeThumbnail(path));

            if (string.Equals(this.thisNewBaselineSource, path, StringComparison.Ordinal))
            {
                this.NewBaselineThumbnail.Source = bitmap;
            }
        }

        private async void OnAddNewBaselineClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                await this.AddNewBaselineAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Adding a new oscilloscope baseline failed - [{ex}]");
            }
        }

        // ###########################################################################################
        // The new baseline into the draft - its screenshot under the name the published ones have
        // ("Scope baseline/<label>_<pin>_<signal>"), in the region of the component's row, the rows as
        // one step, the draft saved. Nothing while it still needs something.
        // ###########################################################################################
        internal async Task AddNewBaselineAsync()
        {
            if (this.thisNewBaselineSource is not { } source ||
                ComponentPaneWording.BaselineProblem(
                    this.NewBaselinePinBox.Text,
                    this.NewBaselineNameBox.Text,
                    source) is not null ||
                this.NewBaselinePinProblem() is not null)
            {
                return;
            }

            this.CommitHeld();

            string pin = this.NewBaselinePinBox.Text?.Trim() ?? string.Empty;
            string signal = this.NewBaselineNameBox.Text?.Trim() ?? string.Empty;
            string name = ComponentPaneFileNames.ForBaseline(this.thisLabel, pin, signal, source, this.TakenNames());

            var baseline = new ComponentImageEntry
            {
                Region = this.NewBaselineRegion(),
                Pin = pin,
                Name = signal,
                File = name,
                TimeDiv = this.NewBaselineTimeDivBox.Text?.Trim() ?? string.Empty,
                VoltsDiv = this.NewBaselineVoltsDivBox.Text?.Trim() ?? string.Empty,
                TriggerLevelVolts = this.NewBaselineTriggerBox.Text?.Trim() ?? string.Empty,
            };

            bool added = await this.AddToDraftAsync(source, name, () =>
            {
                this.thisBaselines.Add(ComponentPaneBaselineRow.From(baseline));
                this.CommitBaselines();
            });

            if (added)
            {
                this.CloseNewBaseline();
                this.ShowEmptyLines();
            }
        }

        // ---- thumbnails -------------------------------------------------------------------------

        // Every row without its thumbnail gets it - read once per file's bytes while the pane is open
        // (PaneImageCache).
        private void LoadThumbnails()
        {
            if (this.thisSession is not { } session || this.thisFilesArriving)
            {
                return;
            }

            string dataRoot = DataManager.DataRoot;
            string draftFolder = DraftFolderLayout.GetBoardFolder(session.DraftsRoot, session.ExcelDataFile);
            string identity = session.ExcelDataFile;

            IEnumerable<IComponentPaneThumbnailRow> rows = this.thisBaselines.Cast<IComponentPaneThumbnailRow>().Concat(this.thisPictures);

            foreach (IComponentPaneThumbnailRow row in rows.Where(row => row.Thumbnail is null && row.File.Length > 0))
            {
                this.TrackPreviewLoad(this.LoadThumbnailAsync(row, dataRoot, draftFolder, identity));
            }
        }

        private async Task LoadThumbnailAsync(IComponentPaneThumbnailRow row, string dataRoot, string draftFolder, string identity)
        {
            string file = row.File;
            int generation = this.thisImages.Generation;
            Func<string, Task>? decoded = this.ImageDecodedForTests;

            (string? path, string? key) = await Task.Run(() =>
            {
                string? found = DraftFileResolver.Resolve(dataRoot, draftFolder, file, identity);
                return (found, PaneImageCache.KeyOf(found, ComponentPane.ThumbnailWidth));
            });

            if (key is null || !this.thisImages.TryGet(key, out PaneImage image))
            {
                Bitmap? bitmap = await Task.Run(() => ComponentPane.DecodeThumbnail(path));

                if (decoded is not null)
                {
                    await decoded(file);
                }

                image = new PaneImage(bitmap, default);

                if (key is not null)
                {
                    this.thisImages.Keep(key, image, generation);
                }
            }

            // Not for a row of a component or board shown before (the cache let go since).
            if (generation == this.thisImages.Generation && string.Equals(row.File, file, StringComparison.Ordinal))
            {
                row.Thumbnail = image.Bitmap;
            }
        }

        // A picture decoded small - null for one that is not there or cannot be read.
        private static Bitmap? DecodeThumbnail(string? path)
        {
            if (string.IsNullOrWhiteSpace(path) || !File.Exists(path) || !ContributionPackaging.IsDisplayableImageFile(path))
            {
                return null;
            }

            try
            {
                using FileStream stream = File.OpenRead(path);
                return Bitmap.DecodeToWidth(stream, ComponentPane.ThumbnailWidth, BitmapInterpolationMode.MediumQuality);
            }
            catch (Exception ex)
            {
                // A file that does not decode is a row without a thumbnail, never a pane that fails.
                Logger.Warning($"Could not read the picture [{path}] for a thumbnail - [{ex.Message}]");
                return null;
            }
        }
    }

    // A row of the pane with a picture: its file in the board, and the thumbnail drawn of it.
    internal interface IComponentPaneThumbnailRow
    {
        string File { get; }

        Bitmap? Thumbnail { get; set; }
    }

    // ###########################################################################################
    // One oscilloscope baseline as the pane holds it. `Committed` is the baseline as the draft has it - what
    // goes back in while an edit would make it no baseline (ComponentPane.CommitBaselines).
    // ###########################################################################################
    public sealed class ComponentPaneBaselineRow : ComponentPaneRegionalRow, IComponentPaneThumbnailRow
    {
        private string thisPin = string.Empty;
        private string thisName = string.Empty;
        private string thisExpectedReading = string.Empty;
        private string thisTimeDiv = string.Empty;
        private string thisVoltsDiv = string.Empty;
        private string thisTriggerLevel = string.Empty;
        private string thisMachineState = string.Empty;
        private string thisNote = string.Empty;
        private string thisFile = string.Empty;
        private string thisProblem = string.Empty;
        private string thisPinProblem = string.Empty;
        private Bitmap? thisThumbnail;

        public ComponentImageEntry Committed { get; set; } = new();

        // Held back, the draft still has it where it was - whatever its "Applies to" says by now.
        public override string RegionInDraft => this.Committed.Region ?? string.Empty;

        public string Pin { get => this.thisPin; set => this.Set(ref this.thisPin, value ?? string.Empty); }

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value ?? string.Empty); }

        public string ExpectedReading { get => this.thisExpectedReading; set => this.Set(ref this.thisExpectedReading, value ?? string.Empty); }

        public string TimeDiv { get => this.thisTimeDiv; set => this.Set(ref this.thisTimeDiv, value ?? string.Empty); }

        public string VoltsDiv { get => this.thisVoltsDiv; set => this.Set(ref this.thisVoltsDiv, value ?? string.Empty); }

        public string TriggerLevel { get => this.thisTriggerLevel; set => this.Set(ref this.thisTriggerLevel, value ?? string.Empty); }

        public string MachineState { get => this.thisMachineState; set => this.Set(ref this.thisMachineState, value ?? string.Empty); }

        public string Note { get => this.thisNote; set => this.Set(ref this.thisNote, value ?? string.Empty); }

        public string File { get => this.thisFile; set => this.Set(ref this.thisFile, value ?? string.Empty); }

        // Why the latest edit is held back - empty when it is not.
        public string Problem
        {
            get => this.thisProblem;
            set
            {
                this.Set(ref this.thisProblem, value ?? string.Empty);
                this.Raise(nameof(this.HasProblem));
            }
        }

        public bool HasProblem => this.thisProblem.Length > 0;

        // The draft's baseline lacks its pin, its signal name or its image - an error Submit stops on,
        // typed in the table or held back here (baseline.required_empty).
        public bool LacksWhatItNeeds => ComponentPaneWording.BaselineProblem(this.Committed.Pin, this.Committed.Name, this.Committed.File) is not null;

        // A pin the component's case does not have - an error, marked on its box (ShowPinProblems).
        public string PinProblem => this.thisPinProblem;

        public bool PinOutside => this.thisPinProblem.Length > 0;

        // The case that has no such pin - null when it can be on it.
        public void UseCaseThatRefuses(string? package)
        {
            if (this.Set(ref this.thisPinProblem, package is null ? string.Empty : BoardDataChecks.PinNotOnPackage(this.thisPin, package), nameof(this.PinProblem)))
            {
                this.Raise(nameof(this.PinOutside));
            }
        }

        public Bitmap? Thumbnail
        {
            get => this.thisThumbnail;
            set
            {
                if (!ReferenceEquals(this.thisThumbnail, value))
                {
                    this.thisThumbnail = value;
                    this.Raise(nameof(this.Thumbnail));
                }
            }
        }

        public static ComponentPaneBaselineRow From(ComponentImageEntry row) => new()
        {
            Region = row.Region ?? string.Empty,
            Committed = row,
            Pin = row.Pin ?? string.Empty,
            Name = row.Name ?? string.Empty,
            ExpectedReading = row.ExpectedOscilloscopeReading ?? string.Empty,
            TimeDiv = row.TimeDiv ?? string.Empty,
            VoltsDiv = row.VoltsDiv ?? string.Empty,
            TriggerLevel = row.TriggerLevelVolts ?? string.Empty,
            MachineState = row.MachineState ?? string.Empty,
            Note = row.Note ?? string.Empty,
            File = row.File ?? string.Empty,

            // Said as it is shown, not only once a commit holds an edit back: its tab is marked already.
            Problem = ComponentPaneWording.BaselineProblem(row.Pin, row.Name, row.File) is { } problem
                ? ComponentPaneWording.BaselineHeldBack(problem)
                : string.Empty,
        };

        public ComponentImageEntry ToEntry() => new()
        {
            Region = this.Region,
            Pin = this.thisPin,
            Name = this.thisName,
            ExpectedOscilloscopeReading = this.thisExpectedReading,
            File = this.thisFile,
            Note = this.thisNote,
            TimeDiv = this.thisTimeDiv,
            VoltsDiv = this.thisVoltsDiv,
            TriggerLevelVolts = this.thisTriggerLevel,
            MachineState = this.thisMachineState,
        };
    }

    // ###########################################################################################
    // One picture as the pane holds it - what the pane does not show of the row (a pin, a reading a
    // picture may carry) kept from the row it was read from.
    // ###########################################################################################
    public sealed class ComponentPanePictureRow : ComponentPaneRegionalRow, IComponentPaneThumbnailRow
    {
        private ComponentImageEntry thisRead = new();
        private string thisName = string.Empty;
        private string thisNote = string.Empty;
        private string thisFile = string.Empty;
        private Bitmap? thisThumbnail;
        private bool thisCanMoveUp;
        private bool thisCanMoveDown;

        public bool CanMoveUp
        {
            get => this.thisCanMoveUp;
            set
            {
                if (this.thisCanMoveUp != value)
                {
                    this.thisCanMoveUp = value;
                    this.Raise(nameof(this.CanMoveUp));
                }
            }
        }

        public bool CanMoveDown
        {
            get => this.thisCanMoveDown;
            set
            {
                if (this.thisCanMoveDown != value)
                {
                    this.thisCanMoveDown = value;
                    this.Raise(nameof(this.CanMoveDown));
                }
            }
        }

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value ?? string.Empty); }

        public string Note { get => this.thisNote; set => this.Set(ref this.thisNote, value ?? string.Empty); }

        public string File { get => this.thisFile; set => this.Set(ref this.thisFile, value ?? string.Empty); }

        public Bitmap? Thumbnail
        {
            get => this.thisThumbnail;
            set
            {
                if (!ReferenceEquals(this.thisThumbnail, value))
                {
                    this.thisThumbnail = value;
                    this.Raise(nameof(this.Thumbnail));
                }
            }
        }

        public static ComponentPanePictureRow From(ComponentImageEntry row) => new()
        {
            thisRead = row,
            Region = row.Region ?? string.Empty,
            Name = row.Name ?? string.Empty,
            Note = row.Note ?? string.Empty,
            File = row.File ?? string.Empty,
        };

        public ComponentImageEntry ToEntry() => new()
        {
            Region = this.Region,
            Pin = this.thisRead.Pin,
            Name = this.thisName,
            ExpectedOscilloscopeReading = this.thisRead.ExpectedOscilloscopeReading,
            File = this.thisFile,
            Note = this.thisNote,
            TimeDiv = this.thisRead.TimeDiv,
            VoltsDiv = this.thisRead.VoltsDiv,
            TriggerLevelVolts = this.thisRead.TriggerLevelVolts,
            MachineState = this.thisRead.MachineState,
        };
    }

    // One file of the component as the pane holds it.
    public sealed class ComponentPaneFileRow : ComponentPaneRegionalRow
    {
        private string thisName = string.Empty;
        private string thisFile = string.Empty;

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value ?? string.Empty); }

        public string File
        {
            get => this.thisFile;
            set
            {
                this.Set(ref this.thisFile, value ?? string.Empty);
                this.Raise(nameof(this.HasFile));
            }
        }

        // "Open file" and "Reveal" are on while the row names a file - a board file's rule.
        public bool HasFile => this.thisFile.Trim().Length > 0;

        public static ComponentPaneFileRow From(ComponentLocalFileEntry row) => new()
        {
            Region = row.Region ?? string.Empty,
            Name = row.Name ?? string.Empty,
            File = row.File ?? string.Empty,
        };

        public ComponentLocalFileEntry ToEntry() => new() { Region = this.Region, Name = this.thisName, File = this.thisFile };
    }
}
