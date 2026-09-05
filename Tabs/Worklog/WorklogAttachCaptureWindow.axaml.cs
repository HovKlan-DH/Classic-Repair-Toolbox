using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Documents;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // Files a just-captured image into one worklog entry: pick the entry, add a comment, attach.
    //
    // Opened from the component popup's captured-image banner after an oscilloscope capture, where
    // the WORKBOOK is already settled (ResolveActiveWorkbook decides it app-wide) but the ENTRY is
    // not. Choosing the entry is the whole reason this dialog exists.
    //
    // Deliberately NOT a third mode on WorklogAddPhotoWindow, even though the two look alike. That
    // window's first half is the file-picker path - Browse, drag-and-drop, ValidateSourceFile,
    // AppleUniformTypeIdentifiers - and none of it applies when the source is a PNG the app just
    // wrote itself. What the two genuinely share is the WRITE, and that is now one implementation
    // in WorklogAttachmentWriter, which both call.
    //
    // The dialog only RESOLVES a choice; it does not write. It returns the chosen entry id and the
    // comment, and the caller performs the attach - the same division WorklogAddPhotoWindow keeps,
    // and it is what lets the caller report a failure in its own UI rather than this modal owning
    // an error state it cannot recover from.
    //
    // Keyboard contract matches the other worklog modals: Escape cancels, Ctrl+Enter attaches.
    // ###########################################################################################
    public partial class WorklogAttachCaptureWindow : Window
    {
        // ###########################################################################################
        // What the dialog returns: which entry to file into, and the comment (which may be empty -
        // an oscilloscope capture of a named pin largely speaks for itself).
        //
        // EntryId is null for the "Create new worklog" choice, which the caller answers by opening
        // the full editor on a new entry rather than attaching to an existing one.
        // ###########################################################################################
        public sealed record AttachResult(int? EntryId, string Comment);

        // One ComboBox row. Holding the target rather than a formatted string keeps the id
        // available on selection, and ToString is what the ComboBox renders.
        private sealed record EntryChoice(WorklogAttachTargets.AttachTarget? Target, string Label)
        {
            public override string ToString() => this.Label;
        }

        // The "Create new worklog" row carries no target; it is the one choice that is not an
        // existing entry, so it is recognised by a null Target rather than by matching its label.
        private static EntryChoice NewEntryChoice => new(null, "Create new worklog");

        private static EntryChoice BuildChoice(WorklogAttachTargets.AttachTarget target) =>
            new(target, WorklogAttachTargets.FormatLabel(target.Entry));

        // ###########################################################################################
        // The matched band's heading, with the component itself picked out: "Worklogs with [U6] in
        // scope", the brackets and the bold naming the thing the grouping is actually keyed on.
        //
        // Built from Runs rather than as one string because a TextBlock cannot mix weights within a
        // single Text - the identical reason TabWorkbooks.Summary.cs walks WorkbookSummary's Stat
        // parts instead of its finished strings. A block carrying Inlines has Text == null, so a
        // test reading only Text sees it as blank; VisibleTextOf below is what reads these back.
        // ###########################################################################################
        private static ComboBoxItem BuildComponentGroupHeader(string component)
        {
            var block = new TextBlock { TextTrimming = TextTrimming.CharacterEllipsis };

            block.Inlines!.Add(new Run("Worklogs with ["));
            block.Inlines.Add(new Run(component) { FontWeight = FontWeight.Bold });
            block.Inlines.Add(new Run("] in scope"));

            return BuildGroupHeader(block);
        }

        // ###########################################################################################
        // A non-selectable heading inside the dropdown, naming the band of worklogs under it.
        //
        // The list is ordered by two rules - component matches first, then everything else, each by
        // ascending id - and the SECOND of those is invisible on screen: a reader sees "#2, #1, #3"
        // and concludes the list is simply unordered, which was reported twice. A caption under the
        // box explaining it did not help, because the explanation is not where the disorder is. So
        // the grouping is stated in the list itself, where it is being applied.
        //
        // Built as a ComboBoxItem with IsEnabled false rather than as another EntryChoice: a
        // disabled ComboBoxItem cannot be selected with the mouse OR with the keyboard, so a header
        // can never end up as SelectedItem. Doing it with a data row plus a SelectionChanged guard
        // means the selection lands on the header first and is then bounced, which flickers and
        // fights arrow-key navigation.
        //
        // Takes the content as an object so it can be a plain string OR the TextBlock above, which
        // carries the bolded component. Avalonia accepts a mix of raw ComboBoxItems and data objects
        // in one ItemsSource; the data rows are wrapped in generated containers as usual, and these
        // are used as-is.
        // ###########################################################################################
        private static ComboBoxItem BuildGroupHeader(object content) => new()
        {
            Content = content,
            IsEnabled = false,

            // Faint enough that it cannot be read as a selectable row. At 0.75 it still looked like
            // an ordinary option - reported - because a ComboBox's own rows carry no styling of
            // their own to contrast against. Opacity rather than a colour: this dialog needs one
            // muted foreground and the theme defines no key for it, so dimming the inherited Fg
            // stays correct in BOTH themes where a hardcoded grey would fail one of them.
            Opacity = 0.45,
            FontSize = 11,
            FontWeight = FontWeight.SemiBold,

            // Left-aligned with the entries BELOW it rather than with the ComboBox's own text
            // padding, so the header sits at the outer edge and its members are indented under it.
            // The two paddings together are what draw the hierarchy - see EntryIndentPadding.
            Padding = new Thickness(6, 4, 6, 2)
        };

        // Indents the worklog rows under their heading, so the grouping reads as a hierarchy rather
        // than as headings interleaved with unrelated rows at the same level. Applied ONLY when the
        // list is actually grouped: with no headers there is nothing to be indented under, and a
        // uniformly shifted list would just look misaligned against the closed box's own text.
        private static readonly Thickness EntryIndentPadding = new(20, 4, 6, 4);

        private string thisSourcePath = string.Empty;

        // Whether the dropdown currently carries group headers, and therefore whether the worklog
        // rows under them should be indented. Set by Initialize, read by the container hook below.
        private bool thisIsGroupedList;

        public WorklogAttachCaptureWindow()
        {
            this.InitializeComponent();

            this.Opened += (_, _) =>
                Dispatcher.UIThread.Post(() => this.CommentTextBox.Focus(), DispatcherPriority.Background);

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);

            // The indent is applied to the generated CONTAINER, which is the only place it can be:
            // the data rows are plain records, so they carry no Padding of their own, and a style
            // targeting ComboBoxItem would hit the headers too - they are ComboBoxItems already and
            // set their own. ContainerPrepared fires for every row as it is realised, including
            // after the list is virtualised and a container is recycled for a different row.
            this.EntryComboBox.ContainerPrepared += this.OnEntryContainerPrepared;

            // The preview bitmap holds an unmanaged surface and would otherwise outlive the dialog.
            this.Closed += (_, _) => (this.ImagePreview.Source as IDisposable)?.Dispose();
        }

        // ###########################################################################################
        // Indents one realised worklog row under its heading.
        //
        // Skips a container that IS one of the headers: those are raw ComboBoxItems placed straight
        // into the list, so Avalonia uses them as their own container and overwriting their padding
        // would undo the outer alignment BuildGroupHeader sets.
        // ###########################################################################################
        private void OnEntryContainerPrepared(object? sender, ContainerPreparedEventArgs e)
        {
            // A header is a raw ComboBoxItem placed straight into the list, so Avalonia uses it as
            // its own container and its Content is the heading (a string, or a TextBlock for the
            // one carrying a bolded component). A worklog row is a generated container whose Content
            // is the EntryChoice - which is what tells the two apart, and what keeps this from
            // overwriting the outer alignment BuildGroupHeader sets.
            if (e.Container is not ComboBoxItem container || container.Content is not EntryChoice)
            {
                return;
            }

            container.Padding = this.thisIsGroupedList ? EntryIndentPadding : new Thickness(6, 4);
        }

        // ###########################################################################################
        // Fills the dialog in. Call before showing.
        //
        // sourcePath is the captured PNG, workbook the one it will be filed into, entries that
        // workbook's entries as read from disk, and componentLabel the board label of the component
        // the capture was taken on ("U8") - which is what floats the matching entries to the top of
        // the list. See WorklogAttachTargets for the ranking.
        // ###########################################################################################
        public void Initialize(
            string sourcePath,
            WorkbookRecord? workbook,
            IReadOnlyList<WorklogEntryRecord>? entries,
            string? componentLabel)
        {
            this.thisSourcePath = sourcePath ?? string.Empty;

            this.SelectedFileText.Text = this.thisSourcePath.Length > 0
                ? Path.GetFileName(this.thisSourcePath)
                : string.Empty;

            this.LoadPreview();

            this.WorkbookText.Text = WorklogAttachTargets.FormatWorkbookLabel(workbook);

            var targets = WorklogAttachTargets.Rank(entries, componentLabel);

            var matched = targets.Where(target => target.IsComponentMatch).ToList();
            var others = targets.Where(target => !target.IsComponentMatch).ToList();

            // Headers are added only when there IS a component match to separate out. With none,
            // every worklog is in one band and a single "All other worklogs" heading over the whole
            // list would name a distinction that is not being drawn - noise, and slightly
            // misleading, since there is no other group for them to be "other" than.
            bool isGrouped = matched.Count > 0;

            var items = new List<object>();

            if (isGrouped)
            {
                string component = componentLabel?.Trim() ?? string.Empty;

                // The component is picked out in bold inside brackets when it is known; with no
                // component to name there is nothing to emphasise, so that heading stays a plain
                // string rather than a TextBlock built to hold no bold run.
                items.Add(component.Length > 0
                    ? BuildComponentGroupHeader(component)
                    : BuildGroupHeader("Worklogs with this component in scope"));
            }

            items.AddRange(matched.Select(BuildChoice));

            if (isGrouped && others.Count > 0)
            {
                items.Add(BuildGroupHeader("All other worklogs"));
            }

            items.AddRange(others.Select(BuildChoice));

            items.Add(NewEntryChoice);

            // Set BEFORE ItemsSource, since assigning it realises containers immediately and the
            // hook reads this flag to decide whether to indent them.
            this.thisIsGroupedList = isGrouped;

            this.EntryComboBox.ItemsSource = items;

            // The best guess - the first component match when there was one, else the first worklog
            // by id - so the common case is a single Attach click. Selected by VALUE rather than by
            // index, because a leading group header occupies index 0 whenever the list is grouped
            // and cannot be selected at all. With no entries the only row is "Create new worklog",
            // which is then the right selection too.
            this.EntryComboBox.SelectedItem = items.FirstOrDefault(item => item is EntryChoice);

            this.UpdateAttachButtonText();
        }

        // ###########################################################################################
        // Shows the captured image. A failure here is not fatal to the dialog - the file name is
        // still shown and the attach still works - because the bytes have already been written to
        // the oscilloscope image folder either way, and refusing to file a capture the app cannot
        // preview would lose the more valuable half of the operation.
        // ###########################################################################################
        private void LoadPreview()
        {
            if (this.thisSourcePath.Length == 0 || !File.Exists(this.thisSourcePath))
            {
                this.ImagePreview.IsVisible = false;
                return;
            }

            try
            {
                this.ImagePreview.Source = new Bitmap(this.thisSourcePath);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not preview captured image [{this.thisSourcePath}]: {ex.Message}");
                this.ImagePreview.IsVisible = false;
            }
        }

        // ###########################################################################################
        // The button says what the click will actually do. Picking "Create new worklog" opens the
        // full editor rather than attaching then and there, and a button still reading "Attach"
        // would misdescribe that - the user would expect the dialog to be the end of it.
        // ###########################################################################################
        private void UpdateAttachButtonText()
        {
            bool isNewEntry = this.EntryComboBox.SelectedItem is EntryChoice choice && choice.Target == null;
            this.AttachButton.Content = isNewEntry ? "Create worklog" : "Attach to existing worklog";
        }

        private void OnEntrySelectionChanged(object? sender, SelectionChangedEventArgs e) =>
            this.UpdateAttachButtonText();

        private void OnAttachClick(object? sender, RoutedEventArgs e)
        {
            if (this.EntryComboBox.SelectedItem is not EntryChoice choice)
            {
                return;
            }

            this.Close(new AttachResult(choice.Target?.Entry.Id, this.CommentTextBox.Text ?? string.Empty));
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter && e.KeyModifiers.HasFlag(KeyModifiers.Control))
            {
                this.OnAttachClick(sender, e);
                e.Handled = true;
            }
        }
    }
}
