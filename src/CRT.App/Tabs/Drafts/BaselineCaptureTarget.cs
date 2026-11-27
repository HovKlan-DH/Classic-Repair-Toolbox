using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Oscilloscope;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // A line of the window's list: a step of the run - its pin, its checkbox, its baseline (null for a new one not accepted yet), a
    // capture of it not yet accepted, the baseline the run started from, and what it shows of them, said to the window as its
    // checkbox changes.
    // ###########################################################################################
    public sealed class BaselineCaptureTarget : INotifyPropertyChanged
    {
        private readonly Action thisPickedChanged;
        private ComponentPaneBaselineRow? thisRow;
        private bool thisIsNext;

        internal BaselineCaptureTarget(BaselineRunStep step, Action pickedChanged)
        {
            this.Step = step;
            this.thisPickedChanged = pickedChanged;
        }

        public event PropertyChangedEventHandler? PropertyChanged;

        public BaselineRunStep Step { get; }

        public string Label => ComponentPaneWording.RunTargetLabel(this.Step);

        public string StateText => ComponentPaneWording.RunTargetState(this.Step);

        public bool IsAccepted => this.Step.IsAccepted;

        // Captured, and not accepted yet - "Not yet accepted".
        public bool IsWaiting => this.Step.IsWaiting;

        public bool IsNext => this.thisIsNext;

        public bool IsNotPicked => !this.Step.IsPicked;

        // Whether there is a picture to see large - a capture waiting, or the baseline's.
        public bool HasPicture => this.Waiting is not null || this.thisRow?.Committed.File is { Length: > 0 };

        // Whether a capture not yet accepted shows in the line, in the baseline's place.
        public bool HasWaitingPicture => this.WaitingPicture is not null;

        // The checkbox: picked to capture.
        public bool IsPicked
        {
            get => this.Step.IsPicked;
            set
            {
                if (this.SetPicked(value))
                {
                    this.thisPickedChanged();
                }
            }
        }

        public ComponentPaneBaselineRow? Row
        {
            get => this.thisRow;
            set
            {
                if (!ReferenceEquals(this.thisRow, value))
                {
                    this.thisRow = value;
                    this.Raise(nameof(this.Row));
                }

                this.Raise(nameof(this.Settings));
                this.Raise(nameof(this.HasSettings));
                this.Raise(nameof(this.HasPicture));
            }
        }

        // The baseline as the run started - null for a pin that had none.
        internal ComponentImageEntry? BeforeEntry { get; init; }

        // The old baseline's picture, read before the first capture - null until then, or when it could not be read.
        internal byte[]? BeforeBytes { get; set; }

        // Those bytes as a picture, read once, off the UI thread.
        internal Task<Bitmap?>? OldPictureRead { get; set; }

        internal Bitmap? OldPicture => this.OldPictureRead is { IsCompletedSuccessfully: true } read ? read.Result : null;

        // A capture not yet accepted - waiting in the user's own folder - and its picture, read off the UI thread.
        internal ComponentPane.ScopeCapture? Waiting { get; private set; }

        internal Task<Bitmap?>? WaitingPictureRead { get; private set; }

        public Bitmap? WaitingPicture => this.WaitingPictureRead is { IsCompletedSuccessfully: true } read ? read.Result : null;

        // Captured again in this run, from a baseline it had: the old one is shown beside it.
        internal bool IsComparable => this.Step.IsCaptured && this.Step.IsUpdate;

        // ###########################################################################################
        // `capture` held, not yet accepted - the one before it, not accepted either, let go of - its picture read from where it waits.
        // ###########################################################################################
        internal void Hold(ComponentPane.ScopeCapture capture)
        {
            ArgumentNullException.ThrowIfNull(capture);

            if (this.Waiting is { } before)
            {
                ComponentPane.DeleteCapture(before.Path);
            }

            this.Waiting = capture;
            this.WaitingPictureRead = Task.Run(() => BaselineCaptureTarget.ReadPicture(capture.Path), CancellationToken.None);
            this.Step.Captured();
            this.ShowPictures();
        }

        // Accepted into the draft as `file` - the waiting screen gone with it (ComponentPane.PutCaptureAsync).
        internal void Accepted(string file)
        {
            this.Step.Accepted(file);
            this.Waiting = null;
            this.WaitingPictureRead = null;
            this.ShowPictures();
        }

        // Let go of, not accepted - its screen gone from the user's folder.
        internal void LetGoOfCapture()
        {
            if (this.Waiting is { } capture)
            {
                ComponentPane.DeleteCapture(capture.Path);
            }

            this.Waiting = null;
            this.WaitingPictureRead = null;
        }

        // A capture's picture, as it waits - none for one that cannot be read.
        private static Bitmap? ReadPicture(string path)
        {
            try
            {
                using FileStream stream = BaselineCaptureTarget.OpenCapture(path);
                return new Bitmap(stream);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Could not read the capture [{path}] - [{ex.Message}]");
                return null;
            }
        }

        // ###########################################################################################
        // *** A CAPTURE IS OPENED SO THAT IT CAN BE DELETED WHILE IT IS READ (FileShare.Delete). *** Its picture is read off the UI
        // thread, and the capture let go of meanwhile - "Capture again" (Hold), another component shown, the window closed - deletes
        // it then: opened with File.OpenRead, Windows refused that delete, which was only logged, and the screen stayed in the user's
        // folder for good (2026-11-26: ComponentPaneBaselineRunTests failed on it at random, under load).
        // ###########################################################################################
        internal static FileStream OpenCapture(string path) =>
            new(path, FileMode.Open, FileAccess.Read, FileShare.Read | FileShare.Delete);

        // ###########################################################################################
        // The baseline as the capture waiting would make it - `row` as the pane holds it, or a new one of the pin in `machineState`, in
        // `newRegion` (the pane's RegionForCapturedBaseline) - by the very rule the pane puts it in by once accepted (BaselineSettingWords
        // .Captured; code review, 2026-11-21: shown as "All regions", put in the region shown). Null with no capture waiting.
        // ###########################################################################################
        internal ComponentImageEntry? WaitingEntry(ComponentPaneBaselineRow? row, string machineState, string newRegion)
        {
            if (this.Waiting is not { } capture)
            {
                return null;
            }

            ComponentImageEntry? held = row?.ToEntry();
            return BaselineSettingWords.Captured(held, capture.Settings, held?.File ?? string.Empty, this.Step.PinText, this.Step.Name, machineState, newRegion);
        }

        // How the oscilloscope was set for what the line shows - a capture waiting, else its baseline: "100us, 2V, 1.25V".
        public string Settings =>
            this.WaitingEntry(this.thisRow, string.Empty, string.Empty) is { } waiting ? ComponentPaneWording.RunSettings(waiting)
            : this.thisRow is null ? string.Empty
            : ComponentPaneWording.RunSettings(this.thisRow.Committed);

        public bool HasSettings => this.Settings.Length > 0;

        // Picked or not, without telling the window - a quick pick tells it once, after them all.
        internal bool SetPicked(bool picked)
        {
            if (this.Step.IsPicked == picked)
            {
                return false;
            }

            this.Step.IsPicked = picked;
            this.Raise(nameof(this.IsPicked));
            this.Raise(nameof(this.IsNotPicked));
            return true;
        }

        internal void ShowState(bool isNext)
        {
            this.thisIsNext = isNext;
            this.Raise(nameof(this.IsNext));
            this.Raise(nameof(this.IsAccepted));
            this.Raise(nameof(this.IsWaiting));
            this.Raise(nameof(this.StateText));
        }

        // What the line shows of a capture changed - held, read, accepted.
        internal void ShowPictures()
        {
            this.Raise(nameof(this.WaitingPicture));
            this.Raise(nameof(this.HasWaitingPicture));
            this.Raise(nameof(this.HasPicture));
            this.Raise(nameof(this.Settings));
            this.Raise(nameof(this.HasSettings));
            this.Raise(nameof(this.IsWaiting));
            this.Raise(nameof(this.IsAccepted));
            this.Raise(nameof(this.StateText));
        }

        internal void ShowLabel() => this.Raise(nameof(this.Label));

        private void Raise(string name) => this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
    }

    // A key of the window's and what it does, along its foot.
    public sealed record BaselineCaptureKey(string Keys, string What);
}
