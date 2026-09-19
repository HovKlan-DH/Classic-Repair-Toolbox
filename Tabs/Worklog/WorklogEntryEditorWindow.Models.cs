using System;

namespace CRT
{
    // ###########################################################################################
    // Row types for the editor's ItemsControls. Public and top-level so the compiled DataTemplates
    // in WorklogEntryEditorWindow.axaml can bind to them - same reasoning as WorklogEntryComponentRow
    // in TabSchematics.Worklog.cs.
    // ###########################################################################################
    public sealed class WorklogLinkRow
    {
        public int Id { get; set; }
        public string Headline { get; set; } = string.Empty;
        public string Url { get; set; } = string.Empty;
    }

    public sealed class WorklogCommentRow
    {
        public int Id { get; set; }
        public string Text { get; set; } = string.Empty;
        public string DateText { get; set; } = string.Empty;
    }

    public sealed class WorklogWorkDoneRow
    {
        public int Id { get; set; }
        public string Text { get; set; } = string.Empty;
        public string DateText { get; set; } = string.Empty;
        public string SummaryText { get; set; } = string.Empty;
    }

    public sealed class WorklogAttachmentRow : System.ComponentModel.INotifyPropertyChanged
    {
        public int Id { get; set; }
        public string FileName { get; set; } = string.Empty;
        public string Comment { get; set; } = string.Empty;

        // ###########################################################################################
        // The file name without the "{id}_" storage prefix - what the Files list shows as its link
        // text. The prefix keeps names unique on disk and means nothing to the user.
        // ###########################################################################################
        public string DisplayFileName { get; set; } = string.Empty;

        // ###########################################################################################
        // Thumbnail for a photo row, decoded once when the row is built rather than by a binding
        // converter, so a file that has gone missing or will not decode simply leaves this null and
        // the row still lists its name and comment. Always null for file rows, which show no image.
        // ###########################################################################################
        public Avalonia.Media.Imaging.Bitmap? Thumbnail { get; set; }

        public bool HasThumbnail => this.Thumbnail != null;

        // ###########################################################################################
        // Shown in place of the thumbnail when the image is unavailable, so a broken photo row reads
        // as broken instead of as a blank square.
        // ###########################################################################################
        public bool HasNoThumbnail => this.Thumbnail == null;

        // ###########################################################################################
        // Hides the comment line entirely when there is none, keeping rows compact - a photo is
        // allowed to carry no comment.
        // ###########################################################################################
        public bool HasComment => !string.IsNullOrWhiteSpace(this.Comment);

        // ###########################################################################################
        // True while this row is the one being dragged, which draws it as an empty outlined slot
        // showing where a drop would land. Following SchematicThumbnail's IsDropPlaceholder: the
        // template swaps between the placeholder box and the real content on this flag.
        //
        // Unlike the thumbnail list, no separate placeholder object is inserted - the dragged row
        // moves within the collection and renders as the placeholder itself, so the gap is exactly
        // the height of the row being moved and the list shows the order it will end up in.
        // ###########################################################################################
        public bool IsDropPlaceholder
        {
            get => this.thisIsDropPlaceholder;
            set
            {
                if (this.thisIsDropPlaceholder == value)
                {
                    return;
                }

                this.thisIsDropPlaceholder = value;
                this.PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(this.IsDropPlaceholder)));
                this.PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(this.IsNotDropPlaceholder)));
            }
        }

        public bool IsNotDropPlaceholder => !this.thisIsDropPlaceholder;

        private bool thisIsDropPlaceholder;

        // ###########################################################################################
        // The row's own height while it is the placeholder, so the gap matches the row being dragged
        // rather than collapsing to the empty box's natural size.
        //
        // MUST raise PropertyChanged: BeginPhotoDragPlaceholder measures the row and assigns this
        // immediately before setting IsDropPlaceholder, and without notification the Height binding
        // kept whatever value it first read. The placeholder then drew at a fixed size regardless of
        // the row - unnoticeable for photo rows, which happen to be about that tall, and obvious in
        // the Files list, where a ~50px row left a gap three times its height.
        //
        // The starting value is only used if a drag somehow begins before the row has been measured.
        // It is deliberately small: too short is a brief visual glitch, too tall is the bug above.
        // ###########################################################################################
        public double PlaceholderHeight
        {
            get => this.thisPlaceholderHeight;
            set
            {
                // The value is ALWAYS stored; only the notification is gated. Returning early
                // without assigning kept the previous row's height in the field, so the property
                // and the row it describes disagreed - and a row that genuinely measured within
                // half a pixel of the seed value was indistinguishable from one never measured at
                // all. Storing first keeps the state honest; the threshold still suppresses the
                // sub-pixel churn that made the list flicker mid-drag.
                bool isMeaningfulChange = Math.Abs(this.thisPlaceholderHeight - value) >= 0.5;

                this.thisPlaceholderHeight = value;

                if (isMeaningfulChange)
                {
                    this.PropertyChanged?.Invoke(this, new System.ComponentModel.PropertyChangedEventArgs(nameof(this.PlaceholderHeight)));
                }
            }
        }

        private double thisPlaceholderHeight = 48.0;

        public event System.ComponentModel.PropertyChangedEventHandler? PropertyChanged;
    }
}
