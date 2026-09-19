using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // The background worker that syncs a component image's oscilloscope baseline settings
    // (trigger level, time/div, volts/div) to the connected scope whenever the selected component
    // image changes. See TabOscilloscope.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class TabOscilloscope
    {
        private string thisLastOscilloscopeImageSyncSignature = string.Empty;
        private readonly object thisPendingImageSyncLock = new();
        private PendingComponentImageSyncRequest? thisPendingComponentImageSyncRequest;
        private readonly SemaphoreSlim thisImageSyncSignal = new(0);
        private CancellationTokenSource? thisImageSyncWorkerCts;
        private Task? thisImageSyncWorkerTask;

        // ###########################################################################################
        // Applies oscilloscope settings from a board Excel component image entry using the standard
        // SetTimeDiv, SetVoltsDiv and SetTriggerLevel palettes so all SCPI traffic is logged.
        // This overload captures a UI snapshot first, then delegates to the fully decoupled worker
        // implementation that no longer reads UI controls directly.
        // ###########################################################################################
        public async Task ApplyComponentImageOscilloscopeSettingsAsync(
            ComponentImageEntry componentImageEntry,
            CancellationToken cancellationToken)
        {
            OscilloscopeSelectionSnapshot selectionSnapshot;

            if (Dispatcher.UIThread.CheckAccess())
            {
                selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();
            }
            else
            {
                selectionSnapshot = await Dispatcher.UIThread.InvokeAsync(
                    this.CreateOscilloscopeSelectionSnapshot,
                    DispatcherPriority.Background);
            }

            await this.ApplyComponentImageOscilloscopeSettingsAsync(
                componentImageEntry,
                selectionSnapshot,
                cancellationToken).ConfigureAwait(false);
        }

        // ###########################################################################################
        // Applies oscilloscope settings using a previously captured oscilloscope snapshot so the
        // entire auto-sync path can run off the UI thread while still reusing the live SCPI session.
        // ###########################################################################################
        private async Task ApplyComponentImageOscilloscopeSettingsAsync(
            ComponentImageEntry componentImageEntry,
            OscilloscopeSelectionSnapshot selectionSnapshot,
            CancellationToken cancellationToken)
        {
            bool enteredSemaphore = false;

            try
            {
                await this.thisOscilloscopeImageSyncSemaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
                enteredSemaphore = true;

                if (componentImageEntry == null ||
                    string.IsNullOrWhiteSpace(componentImageEntry.Pin) ||
                    (string.IsNullOrWhiteSpace(componentImageEntry.TimeDiv) &&
                     string.IsNullOrWhiteSpace(componentImageEntry.VoltsDiv) &&
                     string.IsNullOrWhiteSpace(componentImageEntry.TriggerLevelVolts)))
                {
                    this.thisLastOscilloscopeImageSyncSignature = string.Empty;
                    return;
                }

                if (!selectionSnapshot.HasActiveEstablishedSession)
                {
                    return;
                }

                var selectedOscilloscope = selectionSnapshot.SelectedOscilloscope;
                if (selectedOscilloscope == null)
                {
                    return;
                }

                ScopeMappedValue? mappedTimeDiv = null;
                if (!string.IsNullOrWhiteSpace(componentImageEntry.TimeDiv))
                {
                    if (ScopeValueMapper.TryMapTimeDiv(componentImageEntry, selectedOscilloscope, out ScopeMappedValue timeDiv))
                    {
                        mappedTimeDiv = timeDiv;
                    }
                    else
                    {
                        this.AppendOutputLine("Warning", $"Could not map image T/DIV value [{componentImageEntry.TimeDiv}] to a supported oscilloscope value");
                    }
                }

                ScopeMappedValue? mappedVoltsDiv = null;
                if (!string.IsNullOrWhiteSpace(componentImageEntry.VoltsDiv))
                {
                    if (ScopeValueMapper.TryMapVoltsDiv(componentImageEntry, selectedOscilloscope, out ScopeMappedValue voltsDiv))
                    {
                        mappedVoltsDiv = voltsDiv;
                    }
                    else
                    {
                        this.AppendOutputLine("Warning", $"Could not map image VOLTS/DIV value [{componentImageEntry.VoltsDiv}] to a supported oscilloscope value");
                    }
                }

                ScopeMappedValue? mappedTriggerLevel = null;
                if (!string.IsNullOrWhiteSpace(componentImageEntry.TriggerLevelVolts))
                {
                    if (ScopeValueMapper.TryMapTriggerLevel(componentImageEntry, out ScopeMappedValue triggerLevel))
                    {
                        mappedTriggerLevel = triggerLevel;
                    }
                    else
                    {
                        this.AppendOutputLine("Warning", $"Could not parse image trigger level value [{componentImageEntry.TriggerLevelVolts}]");
                    }
                }

                if (mappedTimeDiv == null &&
                    mappedVoltsDiv == null &&
                    mappedTriggerLevel == null)
                {
                    return;
                }

                string signature = BuildComponentImageSyncSignature(
                    componentImageEntry,
                    selectionSnapshot,
                    mappedTimeDiv,
                    mappedVoltsDiv,
                    mappedTriggerLevel);

                if (string.Equals(this.thisLastOscilloscopeImageSyncSignature, signature, StringComparison.Ordinal))
                {
                    return;
                }

                this.AppendOutputLine("Debug", "---");
                this.AppendOutputLine("Info", $"Auto-syncing oscilloscope settings from image pin [{componentImageEntry.Pin.Trim()}]");

                if (mappedTimeDiv != null)
                {
                    this.AppendOutputLine("Info", $"Image TIME/DIV [{mappedTimeDiv.RawValue}] mapped to scope value [{mappedTimeDiv.MatchedDisplayValue}]");
                }

                if (mappedVoltsDiv != null)
                {
                    this.AppendOutputLine("Info", $"Image VOLTS/DIV [{mappedVoltsDiv.RawValue}] mapped to scope value [{mappedVoltsDiv.MatchedDisplayValue}]");
                }

                if (mappedTriggerLevel != null)
                {
                    this.AppendOutputLine("Info", $"Image trigger level [{mappedTriggerLevel.RawValue}] mapped to SCPI value [{mappedTriggerLevel.ScpiValue}]");
                }

                await this.RunWithEstablishedOscilloscopeSessionAsync(
                    selectionSnapshot,
                    async (scopeClient, oscilloscopeEntry, token) =>
                    {
                        if (mappedTimeDiv != null)
                        {
                            this.thisLastTimeDivSeconds = mappedTimeDiv.NumericValue;
                            await this.ExecutePaletteAsync(
                                scopeClient,
                                oscilloscopeEntry,
                                ScopeCommandPalette.SetTimeDiv,
                                token).ConfigureAwait(false);
                        }

                        if (mappedVoltsDiv != null)
                        {
                            this.thisLastVoltsDivVolts = mappedVoltsDiv.NumericValue;
                            await this.ExecutePaletteAsync(
                                scopeClient,
                                oscilloscopeEntry,
                                ScopeCommandPalette.SetVoltsDiv,
                                token).ConfigureAwait(false);
                        }

                        if (mappedTriggerLevel != null)
                        {
                            this.thisLastTriggerLevelVolts = mappedTriggerLevel.NumericValue;
                            await this.ExecutePaletteAsync(
                                scopeClient,
                                oscilloscopeEntry,
                                ScopeCommandPalette.SetTriggerLevel,
                                token).ConfigureAwait(false);
                        }
                    },
                    cancellationToken,
                    writeWarnings: false).ConfigureAwait(false);

                this.thisLastOscilloscopeImageSyncSignature = signature;
            }
            catch (OperationCanceledException)
            {
            }
            finally
            {
                if (enteredSemaphore)
                {
                    this.thisOscilloscopeImageSyncSemaphore.Release();
                }
            }
        }

        // ###########################################################################################
        // Builds a stable signature for the currently selected component image so repeated callbacks
        // for the same image are skipped, while actual image changes still resend SCPI commands.
        // ###########################################################################################
        private static string BuildComponentImageSyncSignature(
            ComponentImageEntry componentImageEntry,
            OscilloscopeSelectionSnapshot selectionSnapshot,
            ScopeMappedValue? mappedTimeDiv,
            ScopeMappedValue? mappedVoltsDiv,
            ScopeMappedValue? mappedTriggerLevel)
        {
            return string.Join(
                "|",
                selectionSnapshot.Host,
                selectionSnapshot.Port,
                selectionSnapshot.SelectedOscilloscope?.Brand ?? string.Empty,
                selectionSnapshot.SelectedOscilloscope?.SeriesOrModel ?? string.Empty,
                componentImageEntry.BoardLabel?.Trim() ?? string.Empty,
                componentImageEntry.Region?.Trim() ?? string.Empty,
                componentImageEntry.Pin?.Trim() ?? string.Empty,
                componentImageEntry.Name?.Trim() ?? string.Empty,
                componentImageEntry.File?.Trim() ?? string.Empty,
                mappedTimeDiv?.ScpiValue ?? string.Empty,
                mappedVoltsDiv?.ScpiValue ?? string.Empty,
                mappedTriggerLevel?.ScpiValue ?? string.Empty);
        }

        // ###########################################################################################
        // Returns the configured per-scope image auto-sync debounce delay in milliseconds.
        // Falls back to 250ms when the current oscilloscope has no valid Debounce-Time value.
        // ###########################################################################################
        public int GetComponentImageSyncDebounceDelayMilliseconds()
        {
            return this.CreateOscilloscopeSelectionSnapshot().DebounceDelayMilliseconds;
        }

        // ###########################################################################################
        // Queues the latest component image for oscilloscope sync using latest-wins semantics.
        // Intermediate rapid image changes are collapsed so the UI remains responsive.
        // ###########################################################################################
        public void QueueComponentImageOscilloscopeSync(ComponentImageEntry? componentImageEntry)
        {
            lock (this.thisPendingImageSyncLock)
            {
                this.thisPendingComponentImageSyncRequest =
                    componentImageEntry == null
                        ? null
                        : new PendingComponentImageSyncRequest
                        {
                            ComponentImageEntry = componentImageEntry,
                            SelectionSnapshot = this.CreateOscilloscopeSelectionSnapshot()
                        };
            }

            if (componentImageEntry == null)
            {
                this.thisLastOscilloscopeImageSyncSignature = string.Empty;
            }

            this.EnsureImageSyncWorkerStarted();

            if (this.thisImageSyncSignal.CurrentCount == 0)
            {
                this.thisImageSyncSignal.Release();
            }
        }

        // ###########################################################################################
        // Starts the background image-sync worker once and lets it process only the latest request.
        // ###########################################################################################
        private void EnsureImageSyncWorkerStarted()
        {
            if (this.thisImageSyncWorkerTask != null && !this.thisImageSyncWorkerTask.IsCompleted)
            {
                return;
            }

            this.thisImageSyncWorkerCts?.Cancel();
            this.thisImageSyncWorkerCts?.Dispose();
            this.thisImageSyncWorkerCts = new CancellationTokenSource();
            this.thisImageSyncWorkerTask = this.RunImageSyncWorkerAsync(this.thisImageSyncWorkerCts.Token);
        }

        // ###########################################################################################
        // Processes queued image-sync requests on a fully decoupled background worker and applies
        // only the latest request after the per-scope debounce interval has elapsed.
        // ###########################################################################################
        private async Task RunImageSyncWorkerAsync(CancellationToken cancellationToken)
        {
            string lastProcessedRequestSignature = string.Empty;

            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    await this.thisImageSyncSignal.WaitAsync(cancellationToken).ConfigureAwait(false);

                    PendingComponentImageSyncRequest? pendingRequest;
                    lock (this.thisPendingImageSyncLock)
                    {
                        pendingRequest = this.thisPendingComponentImageSyncRequest;
                    }

                    if (pendingRequest == null)
                    {
                        lastProcessedRequestSignature = string.Empty;
                        continue;
                    }

                    await Task.Delay(
                        pendingRequest.SelectionSnapshot.DebounceDelayMilliseconds,
                        cancellationToken).ConfigureAwait(false);

                    lock (this.thisPendingImageSyncLock)
                    {
                        pendingRequest = this.thisPendingComponentImageSyncRequest;
                        this.thisPendingComponentImageSyncRequest = null;
                    }

                    if (pendingRequest == null)
                    {
                        lastProcessedRequestSignature = string.Empty;
                        continue;
                    }

                    string requestSignature = string.Join(
                        "|",
                        pendingRequest.SelectionSnapshot.Host,
                        pendingRequest.SelectionSnapshot.Port,
                        pendingRequest.SelectionSnapshot.SelectedOscilloscope?.Brand ?? string.Empty,
                        pendingRequest.SelectionSnapshot.SelectedOscilloscope?.SeriesOrModel ?? string.Empty,
                        pendingRequest.ComponentImageEntry.BoardLabel?.Trim() ?? string.Empty,
                        pendingRequest.ComponentImageEntry.Region?.Trim() ?? string.Empty,
                        pendingRequest.ComponentImageEntry.Pin?.Trim() ?? string.Empty,
                        pendingRequest.ComponentImageEntry.Name?.Trim() ?? string.Empty,
                        pendingRequest.ComponentImageEntry.File?.Trim() ?? string.Empty);

                    if (string.Equals(lastProcessedRequestSignature, requestSignature, StringComparison.Ordinal))
                    {
                        continue;
                    }

                    await this.ApplyComponentImageOscilloscopeSettingsAsync(
                        pendingRequest.ComponentImageEntry,
                        pendingRequest.SelectionSnapshot,
                        cancellationToken).ConfigureAwait(false);

                    lastProcessedRequestSignature = requestSignature;
                }
            }
            catch (OperationCanceledException)
            {
            }
        }

        // ###########################################################################################
        // Container for one queued component-image sync request, pairing the image entry with the
        // oscilloscope UI snapshot that was current when the image was selected.
        // ###########################################################################################
        private sealed class PendingComponentImageSyncRequest
        {
            public ComponentImageEntry ComponentImageEntry { get; init; } = new();
            public OscilloscopeSelectionSnapshot SelectionSnapshot { get; init; } = new();
        }
    }
}
