using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Data;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A TEXT BOX THAT FORMATS ITS NOTE AS IT IS TYPED (owner request, 2026-11-07: "a very simple WYSIWYG ... Bold,
    // Italic, Underscore, Strikethrough, Bulletlist (nested) ... I do not want a full fledged editor ... nor should it feel
    // 'heavy' and slow in usage") - for the notes of a device, a workbook and a worklog, and of a component's pictures and
    // oscilloscope baselines.
    //
    // *** A TextBox, so typing, the caret, selection, the clipboard and a keyboard's composition are Avalonia's own. *** It
    // holds the note as its plain text (Text - what is shown and edited) and each character's style; what is stored is
    // MARKUP (RichNote - **bold**, *italic*, __underlined__, ~~struck~~, "- " bullets), two-way bindable as any value. The
    // styles are drawn by its RichNoteTextPresenter, in the template RichNoteTheme.axaml gives it - Fluent's TextBox's own
    // with a row of five small buttons over the text.
    //
    // KEYS: Ctrl+B bold, Ctrl+I italic, Ctrl+U underline, Ctrl+Shift+X struck through, Ctrl+Shift+8 bullets on or off -
    // on the selection, or, with nothing selected, for what is typed next; Tab and Shift+Tab nest a bullet deeper or less
    // (elsewhere Tab is Tab as ever); Enter on a bullet starts the next, on an empty one ends the list; Backspace at a
    // bullet's words takes it away. Every rule is RichNoteEditing's.
    //
    // UNDO IS ITS OWN (Ctrl+Z, Ctrl+Y or Ctrl+Shift+Z): the TextBox's undoes the text alone, and would leave the styles of
    // what it took back behind. Typed characters in a row are one step, as the TextBox's were.
    //
    // Code reads and writes Markup - never Text, which is the note without its formatting.
    // ###########################################################################################
    public class RichNoteTextBox : TextBox
    {
        public static readonly StyledProperty<string> MarkupProperty =
            AvaloniaProperty.Register<RichNoteTextBox, string>(nameof(Markup), string.Empty, defaultBindingMode: BindingMode.TwoWay);

        private List<RichNoteStyle> thisStyles = new();

        private string thisShownText = string.Empty;

        // The markup last shown or written - one handed in again changes nothing.
        private string thisPublishedMarkup = string.Empty;

        // Set while the box writes its own Text or Markup, so its handlers do not take that for a change from outside.
        private bool thisWriting;

        // A style asked for with nothing selected (Ctrl+B and typing): what the next character typed AT `thisTypingAt` takes.
        private RichNoteStyle? thisTypingStyle;

        private int thisTypingAt = -1;

        private RichNoteTextPresenter? thisPresenter;

        private readonly List<Snapshot> thisUndo = new();

        private readonly List<Snapshot> thisRedo = new();

        // When the last step was taken by typing one character - a run of them is one step.
        private DateTime thisLastTypedUtc = DateTime.MinValue;

        private const int MostSteps = 200;

        private static readonly TimeSpan TypingRun = TimeSpan.FromSeconds(1.5);

        private sealed record Snapshot(string Text, RichNoteStyle[] Styles, int Caret);

        public RichNoteTextBox()
        {
            this.AcceptsReturn = true;

            // ###########################################################################################
            // *** ENTER INSERTS "\n", ON WINDOWS TOO (code review, 2026-11-07). *** A TextBox inserts its NewLine, which is
            // Environment.NewLine unless set - "\r\n" on Windows - and a note holding "\r" never read back as written
            // (RichNote reads line ends as "\n"), so every note of more than one line was stored with a "\" before each
            // literal mark character. A note pasted with "\r\n" is written with "\n" by RichNote itself (Tidy).
            // ###########################################################################################
            this.NewLine = "\n";
            this.TextWrapping = TextWrapping.Wrap;
            this.IsUndoEnabled = false;
            this.AddHandler(KeyDownEvent, this.OnKeyDownFirst, RoutingStrategies.Tunnel);
        }

        // Its own look - Fluent's TextBox with the format row (RichNoteTheme.axaml).
        protected override Type StyleKeyOverride => typeof(RichNoteTextBox);

        // The note as stored - markup. Set it to show a note; it follows every edit.
        public string Markup
        {
            get => this.GetValue(MarkupProperty);
            set => this.SetValue(MarkupProperty, value);
        }

        // Each character's style - for the tests.
        internal IReadOnlyList<RichNoteStyle> CharacterStylesForTests => this.thisStyles;

        protected override void OnApplyTemplate(TemplateAppliedEventArgs e)
        {
            base.OnApplyTemplate(e);

            this.thisPresenter = e.NameScope.Find<RichNoteTextPresenter>("PART_TextPresenter");

            this.thisStyleButtons.Clear();
            this.thisStyleButtons.Add((RichNoteStyle.Bold, Wire(e, "PART_BoldButton", () => this.ToggleStyle(RichNoteStyle.Bold))));
            this.thisStyleButtons.Add((RichNoteStyle.Italic, Wire(e, "PART_ItalicButton", () => this.ToggleStyle(RichNoteStyle.Italic))));
            this.thisStyleButtons.Add((RichNoteStyle.Underline, Wire(e, "PART_UnderlineButton", () => this.ToggleStyle(RichNoteStyle.Underline))));
            this.thisStyleButtons.Add((RichNoteStyle.Strike, Wire(e, "PART_StrikeButton", () => this.ToggleStyle(RichNoteStyle.Strike))));
            this.thisBulletButton = Wire(e, "PART_BulletButton", this.ToggleBullets);

            this.ShowStyles();
        }

        private static Button? Wire(TemplateAppliedEventArgs e, string name, Action action)
        {
            if (e.NameScope.Find<Button>(name) is not Button button)
            {
                return null;
            }

            button.Click += (_, _) => action();
            return button;
        }

        // ###########################################################################################
        // *** A FORMAT BUTTON IS SHOWN PRESSED WHILE A PRESS WOULD TURN ITS STYLE OFF (owner decision, 2026-11-08: "I think
        // that would be nice so do it"; RichNoteEditing.Pressed - the very rule a press goes by): the selection all bold,
        // or, with nothing selected, bold being what is typed next. Its class "On" (RichNoteTheme.axaml). Shown after every
        // change of the text, its styles, the caret, the selection and the style asked for with nothing selected.
        // ###########################################################################################
        private readonly List<(RichNoteStyle Flag, Button? Button)> thisStyleButtons = new();

        private Button? thisBulletButton;

        private void ShowPressedButtons()
        {
            RichNoteStyle? asked = this.thisTypingStyle is { } style && this.thisTypingAt == this.CaretIndex ? style : null;

            (RichNoteStyle pressed, bool bullets) = RichNoteEditing.Pressed(
                this.thisShownText, this.thisStyles, this.SelectionStart, this.SelectionEnd, asked);

            foreach ((RichNoteStyle flag, Button? button) in this.thisStyleButtons)
            {
                button?.Classes.Set("On", (pressed & flag) != 0);
            }

            this.thisBulletButton?.Classes.Set("On", bullets);
        }

        // ###########################################################################################
        // *** WHERE A CHANGE OF THE TEXT BEGAN (code review, 2026-11-08). *** Compared alone, the text before and after a
        // change cannot tell where a character went beside one like it ("b" typed after a bold "a" in "ab" read as typed
        // after its plain "b", and came out plain). The TextBox sets its new text BEFORE it moves the caret or the
        // selection (Avalonia 12.1.4, decompiled: HandleTextInput, DeleteSelection, Backspace, Delete) and coerces them into
        // the new text only in its own handling of the change - so as this is entered, they still say where it began: the
        // selection's start, or the caret. Backspace with nothing selected takes out what stands BEFORE the caret, so it is
        // noted as its key goes through the box (OnKeyDown).
        // ###########################################################################################
        private bool thisDeletingBackward;

        protected override void OnPropertyChanged(AvaloniaPropertyChangedEventArgs change)
        {
            int? changeAt = !this.thisWriting && change.Property == TextProperty
                ? Math.Min(this.SelectionStart, this.SelectionEnd)
                : null;
            bool nothingSelected = this.SelectionStart == this.SelectionEnd;

            base.OnPropertyChanged(change);

            if (this.thisWriting)
            {
                return;
            }

            if (change.Property == MarkupProperty)
            {
                this.ShowMarkup(change.GetNewValue<string>());
            }
            else if (change.Property == DataContextProperty)
            {
                // Made again for another row (the pane's lists build only the rows in view): none of the last row's steps -
                // with its note the same as the one written last, ShowMarkup shows nothing new and kept them (code review,
                // 2026-11-13: Ctrl+Z brought another card's words into this one's note).
                this.thisUndo.Clear();
                this.thisRedo.Clear();
                this.thisTypingStyle = null;
            }
            else if (change.Property == SelectionStartProperty || change.Property == SelectionEndProperty || change.Property == CaretIndexProperty)
            {
                this.ShowPressedButtons();
            }
            else if (change.Property == TextProperty)
            {
                string text = change.GetNewValue<string?>() ?? string.Empty;

                if (changeAt is { } at && nothingSelected && this.thisDeletingBackward && text.Length < this.thisShownText.Length)
                {
                    changeAt = Math.Max(0, at - (this.thisShownText.Length - text.Length));
                }

                this.FollowTyping(text, changeAt);
            }
        }

        // The TextBox's own handling of a key, Backspace noted for the change it makes (thisDeletingBackward).
        protected override void OnKeyDown(KeyEventArgs e)
        {
            this.thisDeletingBackward = e.Key == Key.Back;

            try
            {
                base.OnKeyDown(e);
            }
            finally
            {
                this.thisDeletingBackward = false;
            }
        }

        // ###########################################################################################
        // A note handed in: shown, its styles drawn - and no step to undo, as it is where editing starts.
        // ###########################################################################################
        private void ShowMarkup(string? markup)
        {
            // The markup the box wrote itself, handed back (a binding's source echoing it a moment later): nothing new to
            // show, and the steps to undo stay.
            if (string.Equals(markup ?? string.Empty, this.thisPublishedMarkup, StringComparison.Ordinal))
            {
                return;
            }

            RichNoteText note = RichNote.Parse(markup);
            this.thisPublishedMarkup = markup ?? string.Empty;

            this.thisWriting = true;

            try
            {
                this.SetCurrentValue(TextProperty, note.Text);
                this.thisShownText = note.Text;
                this.thisStyles = note.Styles.ToList();
            }
            finally
            {
                this.thisWriting = false;
            }

            this.thisUndo.Clear();
            this.thisRedo.Clear();
            this.thisTypingStyle = null;
            this.ShowStyles();
        }

        // ###########################################################################################
        // The text changed by typing, a paste, a cut or a delete - beginning at `changeAt` - the styles follow
        // (RichNoteEditing.FollowTextChange), one character typed after another is one step to undo, and the markup
        // follows.
        // ###########################################################################################
        private void FollowTyping(string text, int? changeAt)
        {
            string before = this.thisShownText;
            bool oneCharacterTyped = text.Length == before.Length + 1;
            DateTime now = DateTime.UtcNow;

            if (!oneCharacterTyped || now - this.thisLastTypedUtc > TypingRun || this.thisUndo.Count == 0)
            {
                this.PushUndo(new Snapshot(before, [.. this.thisStyles], this.CaretIndex));
            }

            this.thisLastTypedUtc = oneCharacterTyped ? now : DateTime.MinValue;

            (int at, _, int inserted) = RichNoteEditing.ChangeOf(before, text, changeAt);
            RichNoteStyle? typing = this.thisTypingStyle is { } asked && inserted > 0 && at == this.thisTypingAt ? asked : null;

            this.thisStyles = RichNoteEditing.FollowTextChange(before, this.thisStyles, text, typing, changeAt);
            this.thisShownText = text;

            if (typing is null)
            {
                this.thisTypingStyle = null;
            }
            else
            {
                this.thisTypingAt = at + inserted;
            }

            this.ShowStyles();
            this.PublishMarkup();
        }

        // ###########################################################################################
        // A style on the selection - or, with nothing selected, for what is typed next at the caret.
        // ###########################################################################################
        public void ToggleStyle(RichNoteStyle flag)
        {
            if (this.IsReadOnly)
            {
                return;
            }

            int start = Math.Min(this.SelectionStart, this.SelectionEnd);
            int end = Math.Max(this.SelectionStart, this.SelectionEnd);

            if (start == end)
            {
                RichNoteStyle now = this.thisTypingStyle is { } asked && this.thisTypingAt == this.CaretIndex
                    ? asked
                    : this.StyleBefore(this.CaretIndex);

                this.thisTypingStyle = now ^ flag;
                this.thisTypingAt = this.CaretIndex;
                this.ShowPressedButtons();
                return;
            }

            this.Apply(RichNoteEditing.Toggle(this.thisShownText, this.thisStyles, start, end, flag));
        }

        // Bullets on or off for the lines the selection touches.
        public void ToggleBullets() =>
            this.Apply(RichNoteEditing.ToggleBullets(this.thisShownText, this.thisStyles, this.SelectionStart, this.SelectionEnd));

        // The style a character typed at `caret` takes without being asked - the one before it on its line.
        private RichNoteStyle StyleBefore(int caret) => RichNoteEditing.StyleTypedAt(this.thisShownText, this.thisStyles, caret);

        // ###########################################################################################
        // The box's own keys - styles, bullets, nesting, its undo - ahead of the TextBox's. None in a box only to be read
        // (code review, 2026-11-07: the component popup's picture note, IsReadOnly, took Enter, Backspace and Tab on its
        // bullets, the style keys and its own Undo, all of which change the note shown); the TextBox's own keys, copying
        // among them, are left to it.
        // ###########################################################################################
        private void OnKeyDownFirst(object? sender, KeyEventArgs e)
        {
            if (this.IsReadOnly)
            {
                return;
            }

            KeyModifiers command = Application.Current?.PlatformSettings?.HotkeyConfiguration.CommandModifiers ?? KeyModifiers.Control;
            bool isCommand = e.KeyModifiers == command;
            bool isCommandShift = e.KeyModifiers == (command | KeyModifiers.Shift);

            RichNoteEdit? edit = null;
            bool handled = true;

            switch (e.Key)
            {
                case Key.B when isCommand:
                    this.ToggleStyle(RichNoteStyle.Bold);
                    break;

                case Key.I when isCommand:
                    this.ToggleStyle(RichNoteStyle.Italic);
                    break;

                case Key.U when isCommand:
                    this.ToggleStyle(RichNoteStyle.Underline);
                    break;

                case Key.X when isCommandShift:
                    this.ToggleStyle(RichNoteStyle.Strike);
                    break;

                case Key.D8 when isCommandShift:
                    this.ToggleBullets();
                    break;

                case Key.Z when isCommand:
                    this.Undo();
                    break;

                case Key.Y when isCommand:
                case Key.Z when isCommandShift:
                    this.Redo();
                    break;

                case Key.Tab when e.KeyModifiers is KeyModifiers.None or KeyModifiers.Shift:
                    edit = RichNoteEditing.Indent(this.thisShownText, this.thisStyles, this.SelectionStart, this.SelectionEnd, e.KeyModifiers == KeyModifiers.Shift ? -1 : 1);
                    handled = edit is not null;
                    break;

                case Key.Enter when e.KeyModifiers == KeyModifiers.None && this.SelectionStart == this.SelectionEnd:
                    edit = RichNoteEditing.EnterOnBullet(this.thisShownText, this.thisStyles, this.CaretIndex);
                    handled = edit is not null;
                    break;

                case Key.Back when e.KeyModifiers == KeyModifiers.None && this.SelectionStart == this.SelectionEnd:
                    edit = RichNoteEditing.BackspaceAtBullet(this.thisShownText, this.thisStyles, this.CaretIndex);
                    handled = edit is not null;
                    break;

                default:
                    handled = false;
                    break;
            }

            if (edit is not null)
            {
                this.Apply(edit);
            }

            if (handled)
            {
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // An edit of the box's own - a style, a bullet, a nesting: one step to undo, the text and the styles shown, the
        // selection where the edit puts it, and the markup following.
        // ###########################################################################################
        private void Apply(RichNoteEdit edit)
        {
            // A box only to be read changes nothing of its own (OnKeyDownFirst); nor does an edit that changes nothing -
            // a style toggled on a bullet's start alone - leave a step to undo (code review, 2026-11-08).
            if (this.IsReadOnly ||
                (string.Equals(edit.Text, this.thisShownText, StringComparison.Ordinal) && edit.Styles.SequenceEqual(this.thisStyles)))
            {
                return;
            }

            this.PushUndo(new Snapshot(this.thisShownText, [.. this.thisStyles], this.CaretIndex));
            this.thisLastTypedUtc = DateTime.MinValue;
            this.Show(edit.Text, edit.Styles, edit.SelectionStart, edit.SelectionEnd);
        }

        private void Show(string text, IReadOnlyList<RichNoteStyle> styles, int selectionStart, int selectionEnd)
        {
            this.thisWriting = true;

            try
            {
                this.SetCurrentValue(TextProperty, text);
                this.thisShownText = text;
                this.thisStyles = styles.ToList();

                // The caret first: a TextBox lets go of its selection as its caret is set (Avalonia 12.1.4's
                // OnCaretIndexChanged), and a selection set after it keeps the caret at its end.
                this.CaretIndex = Math.Clamp(selectionEnd, 0, text.Length);
                this.SelectionStart = Math.Clamp(selectionStart, 0, text.Length);
                this.SelectionEnd = Math.Clamp(selectionEnd, 0, text.Length);
            }
            finally
            {
                this.thisWriting = false;
            }

            this.thisTypingStyle = null;
            this.ShowStyles();
            this.PublishMarkup();
        }

        private void PushUndo(Snapshot snapshot)
        {
            this.thisUndo.Add(snapshot);
            this.thisRedo.Clear();

            if (this.thisUndo.Count > MostSteps)
            {
                this.thisUndo.RemoveAt(0);
            }
        }

        // The step before the last one, back - its text, styles and caret.
        public new void Undo()
        {
            if (this.IsReadOnly || this.thisUndo.Count == 0)
            {
                return;
            }

            Snapshot back = this.thisUndo[^1];
            this.thisUndo.RemoveAt(this.thisUndo.Count - 1);
            this.thisRedo.Add(new Snapshot(this.thisShownText, [.. this.thisStyles], this.CaretIndex));
            this.thisLastTypedUtc = DateTime.MinValue;
            this.Show(back.Text, back.Styles, back.Caret, back.Caret);
        }

        public new void Redo()
        {
            if (this.IsReadOnly || this.thisRedo.Count == 0)
            {
                return;
            }

            Snapshot again = this.thisRedo[^1];
            this.thisRedo.RemoveAt(this.thisRedo.Count - 1);
            this.thisUndo.Add(new Snapshot(this.thisShownText, [.. this.thisStyles], this.CaretIndex));
            this.thisLastTypedUtc = DateTime.MinValue;
            this.Show(again.Text, again.Styles, again.Caret, again.Caret);
        }

        private void ShowStyles()
        {
            if (this.thisPresenter is not null)
            {
                this.thisPresenter.CharacterStyles = this.thisStyles.ToArray();
            }

            this.ShowPressedButtons();
        }

        private void PublishMarkup()
        {
            this.thisWriting = true;

            try
            {
                this.thisPublishedMarkup = RichNote.Serialize(new RichNoteText(this.thisShownText, this.thisStyles));
                this.SetCurrentValue(MarkupProperty, this.thisPublishedMarkup);
            }
            finally
            {
                this.thisWriting = false;
            }
        }
    }
}
