using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.ComponentModel;

namespace CRT
{
    // ###########################################################################################
    // THE LIVE PART OF A DRAFT'S ROW - its save state and its Undo and Redo (CRT 4.0 step 1's last
    // slice, 2026-10-25 - the plan's header: "Draft: 3 changes - Saved   [Undo] [Redo]").
    //
    // Read off the draft's session (DraftManager.Sessions) as it changes, and told to the row through
    // property changes - NOT by rebuilding the row: the row is kept while it draws the same
    // (TabDrafts.KeepUnchangedRows), because a rebuilt row threw the pressed button away ("I need to
    // click it twice"), and a save state changes every second while someone types. The tab keeps one
    // of these per draft for the whole run, so every rebuilt row holds the same one.
    //
    // A draft no editor has touched this run has no session: it is Saved, with nothing to undo.
    // ###########################################################################################
    public sealed class DraftHeaderState : INotifyPropertyChanged
    {
        private DraftSession? thisSession;

        public event PropertyChangedEventHandler? PropertyChanged;

        // "Saved", "Saving..." or "Could not save - <reason>" (DraftSaveWording.StateLine).
        public string SaveStateText { get; private set; } = DraftSaveWording.Saved;

        // A save that could not be made - the text in the failure colour, and Try again.
        public bool IsFailed { get; private set; }

        public bool CanUndo { get; private set; }

        public bool CanRedo { get; private set; }

        public string UndoTip { get; private set; } = DraftSaveWording.UndoTip(null);

        public string RedoTip { get; private set; } = DraftSaveWording.RedoTip(null);

        // ###########################################################################################
        // The session this row follows - none, the one an editor opened, or (after the draft was
        // discarded and made again) another. Following a new one lets go of the old one. Read again
        // either way: a save on the pool tells the row by a post, and the list refreshed after it
        // (a step from the row, TabDrafts.AfterDraftChanged) must not show "Saving..." until then.
        // ###########################################################################################
        public void Follow(DraftSession? session)
        {
            if (ReferenceEquals(session, this.thisSession))
            {
                this.Read();
                return;
            }

            if (this.thisSession is { } previous)
            {
                previous.Changed -= this.OnChanged;
                previous.SaveStateChanged -= this.OnSaveStateChanged;
            }

            this.thisSession = session;

            if (session is not null)
            {
                session.Changed += this.OnChanged;
                session.SaveStateChanged += this.OnSaveStateChanged;
            }

            this.Read();
        }

        // Lets go of the session - the draft is no longer listed.
        public void Forget() => this.Follow(null);

        internal DraftSession? SessionForTests => this.thisSession;

        private void OnChanged(object? sender, DraftSessionChange change) => this.ReadOnUiThread();

        private void OnSaveStateChanged(object? sender, EventArgs e) => this.ReadOnUiThread();

        // A save runs on the pool and says so there; the row is drawn on the UI thread.
        private void ReadOnUiThread()
        {
            if (Dispatcher.UIThread.CheckAccess())
            {
                this.Read();
            }
            else
            {
                Dispatcher.UIThread.Post(this.Read);
            }
        }

        // ###########################################################################################
        // Everything the row shows, read again - and only what changed is told, so a save that
        // changed nothing on screen redraws nothing.
        // ###########################################################################################
        internal void Read()
        {
            DraftSession? session = this.thisSession;

            DraftSaveState state = session?.SaveState ?? DraftSaveState.Saved;

            this.Set(nameof(this.SaveStateText), DraftSaveWording.StateLine(state, session?.SaveFailure), this.SaveStateText, value => this.SaveStateText = value);
            this.Set(nameof(this.IsFailed), state == DraftSaveState.Failed, this.IsFailed, value => this.IsFailed = value);
            this.Set(nameof(this.CanUndo), session?.CanUndo == true, this.CanUndo, value => this.CanUndo = value);
            this.Set(nameof(this.CanRedo), session?.CanRedo == true, this.CanRedo, value => this.CanRedo = value);
            this.Set(nameof(this.UndoTip), DraftSaveWording.UndoTip(session?.UndoDescription), this.UndoTip, value => this.UndoTip = value);
            this.Set(nameof(this.RedoTip), DraftSaveWording.RedoTip(session?.RedoDescription), this.RedoTip, value => this.RedoTip = value);
        }

        private void Set<T>(string name, T value, T current, Action<T> assign)
        {
            if (Equals(value, current))
            {
                return;
            }

            assign(value);
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }
}
