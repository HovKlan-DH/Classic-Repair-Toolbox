using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // The trigger-level, time/div and volts/div keyboard-step workers: three near-identical
    // clusters that queue a keyboard-driven adjustment, coalesce rapid keystrokes and send the
    // resulting SCPI command without waiting for a response on every press. See
    // TabOscilloscope.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class TabOscilloscope
    {
        private readonly SemaphoreSlim thisTriggerLevelKeyboardSignal = new(0);
        private CancellationTokenSource? thisTriggerLevelKeyboardWorkerCts;
        private Task? thisTriggerLevelKeyboardWorkerTask;
        private int thisPendingTriggerLevelKeyboardSteps;

        private readonly SemaphoreSlim thisTimeDivKeyboardSignal = new(0);
        private CancellationTokenSource? thisTimeDivKeyboardWorkerCts;
        private Task? thisTimeDivKeyboardWorkerTask;
        private int thisPendingTimeDivKeyboardSteps;

        private readonly object thisPendingVoltsDivKeyboardLock = new();
        private double? thisPendingVoltsDivKeyboardTargetVolts;
        private readonly SemaphoreSlim thisVoltsDivKeyboardSignal = new(0);
        private CancellationTokenSource? thisVoltsDivKeyboardWorkerCts;
        private Task? thisVoltsDivKeyboardWorkerTask;

        // ###########################################################################################
        // Clears the cached trigger level so the next keyboard-driven trigger step will refresh it
        // from the oscilloscope before sending a new SetTriggerLevel command.
        // ###########################################################################################
        private void InvalidateCachedTriggerLevelVolts()
        {
            this.thisLastTriggerLevelVolts = null;
        }

        // ###########################################################################################
        // Ensures a cached trigger level is available for keyboard stepping. When the cache is empty,
        // it performs one QueryTriggerLevel palette execution and stores the returned value.
        // ###########################################################################################
        private async Task<bool> EnsureCachedTriggerLevelVoltsAsync(
            IScopeClient scopeClient,
            OscilloscopeEntry oscilloscopeEntry,
            CancellationToken cancellationToken)
        {
            if (this.thisLastTriggerLevelVolts.HasValue)
            {
                return true;
            }

            await this.ExecutePaletteAsync(
                scopeClient,
                oscilloscopeEntry,
                ScopeCommandPalette.QueryTriggerLevel,
                cancellationToken).ConfigureAwait(false);

            return this.thisLastTriggerLevelVolts.HasValue;
        }

        // ###########################################################################################
        // Queues one keyboard-driven trigger-level step and starts the dedicated background worker
        // that drains pending Up/Down requests without dropping held-key repeats.
        // ###########################################################################################
        public void QueueTriggerLevelKeyboardStep(int direction)
        {
            if (direction == 0)
            {
                return;
            }

            Interlocked.Add(
                ref this.thisPendingTriggerLevelKeyboardSteps,
                direction > 0 ? 1 : -1);

            this.EnsureTriggerLevelKeyboardWorkerStarted();

            if (this.thisTriggerLevelKeyboardSignal.CurrentCount == 0)
            {
                this.thisTriggerLevelKeyboardSignal.Release();
            }
        }

        // ###########################################################################################
        // Starts the dedicated trigger-level keyboard worker once so held Up/Down keys can be
        // processed as a latest-pending batch instead of being dropped by the popup window.
        // ###########################################################################################
        private void EnsureTriggerLevelKeyboardWorkerStarted()
        {
            if (this.thisTriggerLevelKeyboardWorkerTask != null &&
                !this.thisTriggerLevelKeyboardWorkerTask.IsCompleted)
            {
                return;
            }

            this.thisTriggerLevelKeyboardWorkerCts?.Cancel();
            this.thisTriggerLevelKeyboardWorkerCts?.Dispose();
            this.thisTriggerLevelKeyboardWorkerCts = new CancellationTokenSource();
            this.thisTriggerLevelKeyboardWorkerTask = this.RunTriggerLevelKeyboardWorkerAsync(
                this.thisTriggerLevelKeyboardWorkerCts.Token);
        }

        // ###########################################################################################
        // Stops any queued keyboard trigger stepping and shuts down the dedicated background worker.
        // ###########################################################################################
        private void StopTriggerLevelKeyboardWorker()
        {
            Interlocked.Exchange(ref this.thisPendingTriggerLevelKeyboardSteps, 0);

            if (this.thisTriggerLevelKeyboardWorkerCts == null)
            {
                return;
            }

            try
            {
                this.thisTriggerLevelKeyboardWorkerCts.Cancel();
            }
            catch
            {
                // ObjectDisposedException when the trigger-level keyboard worker was already
                // torn down - Cancel() on a disposed CancellationTokenSource throws, and every
                // caller here is a "stop it if it is running" path that has nothing to do
                // differently when it was not. The Dispose() below still runs.
            }

            this.thisTriggerLevelKeyboardWorkerCts.Dispose();
            this.thisTriggerLevelKeyboardWorkerCts = null;
            this.thisTriggerLevelKeyboardWorkerTask = null;
        }

        // ###########################################################################################
        // Drains queued keyboard trigger steps in batches so repeated Up/Down keypresses remain
        // responsive even while one SCPI write is still in flight.
        // ###########################################################################################
        private async Task RunTriggerLevelKeyboardWorkerAsync(CancellationToken cancellationToken)
        {
            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    await this.thisTriggerLevelKeyboardSignal.WaitAsync(cancellationToken).ConfigureAwait(false);

                    while (!cancellationToken.IsCancellationRequested)
                    {
                        int pendingSteps = Interlocked.Exchange(ref this.thisPendingTriggerLevelKeyboardSteps, 0);
                        if (pendingSteps == 0)
                        {
                            break;
                        }

                        int direction = Math.Sign(pendingSteps);
                        int stepCount = Math.Abs(pendingSteps);
                        OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();

                        await this.RunWithEstablishedOscilloscopeSessionAsync(
                            selectionSnapshot,
                            async (scopeClient, oscilloscopeEntry, token) =>
                            {
                                if (!await this.EnsureCachedTriggerLevelVoltsAsync(
                                    scopeClient,
                                    oscilloscopeEntry,
                                    token).ConfigureAwait(false))
                                {
                                    this.AppendOutputLine("Warning", "Could not read current trigger level from oscilloscope");
                                    return;
                                }

                                double startingTriggerLevelVolts = this.thisLastTriggerLevelVolts!.Value;
                                double currentTriggerLevelVolts = startingTriggerLevelVolts;

                                for (int i = 0; i < stepCount; i++)
                                {
                                    double targetTriggerLevelVolts = ScopeFormatting.GetNextSnappedTriggerLevelVolts(
                                        currentTriggerLevelVolts,
                                        direction);

                                    await this.SendTriggerLevelFastAsync(
                                        scopeClient,
                                        oscilloscopeEntry,
                                        targetTriggerLevelVolts,
                                        token).ConfigureAwait(false);

                                    currentTriggerLevelVolts = targetTriggerLevelVolts;
                                }

                                this.thisLastOscilloscopeImageSyncSignature = string.Empty;

                                this.AppendOutputLine(
                                    "Info",
                                    stepCount == 1
                                        ? $"Keyboard trigger level step: {ScopeFormatting.FormatVoltage(startingTriggerLevelVolts)} -> {ScopeFormatting.FormatVoltage(currentTriggerLevelVolts)}"
                                        : $"Keyboard trigger level step: {ScopeFormatting.FormatVoltage(startingTriggerLevelVolts)} -> {ScopeFormatting.FormatVoltage(currentTriggerLevelVolts)} ({stepCount} steps)");
                            },
                            cancellationToken,
                            writeWarnings: true).ConfigureAwait(false);
                    }
                }
            }
            catch (OperationCanceledException)
            {
            }
        }

        // ###########################################################################################
        // Sends one trigger-level Set command through a lightweight SCPI fast path so repeated
        // keyboard stepping avoids the heavier full palette logging pipeline.
        // ###########################################################################################
        private async Task SendTriggerLevelFastAsync(
            IScopeClient scopeClient,
            OscilloscopeEntry oscilloscopeEntry,
            double targetTriggerLevelVolts,
            CancellationToken cancellationToken)
        {
            this.thisLastTriggerLevelVolts = targetTriggerLevelVolts;

            string commandText = ScopeCommandResolver.GetCommandText(
                oscilloscopeEntry,
                ScopeCommand.SetTriggerLevel);

            if (string.IsNullOrWhiteSpace(commandText))
            {
                throw new InvalidOperationException("No SCPI command text is defined for SetTriggerLevel");
            }

            string effectiveCommandText = this.BuildEffectiveCommandText(
                ScopeCommand.SetTriggerLevel,
                commandText);

            if (string.IsNullOrWhiteSpace(effectiveCommandText))
            {
                throw new InvalidOperationException("No trigger level value is available for SetTriggerLevel");
            }

            await scopeClient.SendAsync(effectiveCommandText, cancellationToken).ConfigureAwait(false);
        }

        // ###########################################################################################
        // Clears the cached TIME/DIV value so the next keyboard-driven TIME/DIV step will refresh it
        // from the oscilloscope before sending a new SetTimeDiv command.
        // ###########################################################################################
        private void InvalidateCachedTimeDivSeconds()
        {
            this.thisLastTimeDivSeconds = null;
        }

        // ###########################################################################################
        // Ensures a cached TIME/DIV value is available for keyboard stepping. When the cache is empty,
        // it performs one QueryTimeDiv palette execution and stores the returned value.
        // ###########################################################################################
        private async Task<bool> EnsureCachedTimeDivSecondsAsync(
            IScopeClient scopeClient,
            OscilloscopeEntry oscilloscopeEntry,
            CancellationToken cancellationToken)
        {
            if (this.thisLastTimeDivSeconds.HasValue)
            {
                return true;
            }

            await this.ExecutePaletteAsync(
                scopeClient,
                oscilloscopeEntry,
                ScopeCommandPalette.QueryTimeDiv,
                cancellationToken).ConfigureAwait(false);

            return this.thisLastTimeDivSeconds.HasValue;
        }

        // ###########################################################################################
        // Queues one keyboard-driven TIME/DIV step and starts the dedicated background worker that
        // drains pending Add/Subtract requests without dropping held-key repeats.
        // ###########################################################################################
        public void QueueTimeDivKeyboardStep(int offset)
        {
            if (offset == 0)
            {
                return;
            }

            Interlocked.Add(
                ref this.thisPendingTimeDivKeyboardSteps,
                offset > 0 ? 1 : -1);

            this.EnsureTimeDivKeyboardWorkerStarted();

            if (this.thisTimeDivKeyboardSignal.CurrentCount == 0)
            {
                this.thisTimeDivKeyboardSignal.Release();
            }
        }

        // ###########################################################################################
        // Starts the dedicated TIME/DIV keyboard worker once so held Add/Subtract keys can be
        // processed as a pending batch instead of being dropped by the popup window.
        // ###########################################################################################
        private void EnsureTimeDivKeyboardWorkerStarted()
        {
            if (this.thisTimeDivKeyboardWorkerTask != null &&
                !this.thisTimeDivKeyboardWorkerTask.IsCompleted)
            {
                return;
            }

            this.thisTimeDivKeyboardWorkerCts?.Cancel();
            this.thisTimeDivKeyboardWorkerCts?.Dispose();
            this.thisTimeDivKeyboardWorkerCts = new CancellationTokenSource();
            this.thisTimeDivKeyboardWorkerTask = this.RunTimeDivKeyboardWorkerAsync(
                this.thisTimeDivKeyboardWorkerCts.Token);
        }

        // ###########################################################################################
        // Stops any queued keyboard TIME/DIV stepping and shuts down the dedicated background worker.
        // ###########################################################################################
        private void StopTimeDivKeyboardWorker()
        {
            Interlocked.Exchange(ref this.thisPendingTimeDivKeyboardSteps, 0);

            if (this.thisTimeDivKeyboardWorkerCts == null)
            {
                return;
            }

            try
            {
                this.thisTimeDivKeyboardWorkerCts.Cancel();
            }
            catch
            {
                // ObjectDisposedException when the time/div keyboard worker was already
                // torn down - Cancel() on a disposed CancellationTokenSource throws, and every
                // caller here is a "stop it if it is running" path that has nothing to do
                // differently when it was not. The Dispose() below still runs.
            }

            this.thisTimeDivKeyboardWorkerCts.Dispose();
            this.thisTimeDivKeyboardWorkerCts = null;
            this.thisTimeDivKeyboardWorkerTask = null;
        }

        // ###########################################################################################
        // Drains queued keyboard TIME/DIV steps in batches so repeated Add/Subtract keypresses remain
        // responsive even while one SCPI write is still in flight.
        // ###########################################################################################
        private async Task RunTimeDivKeyboardWorkerAsync(CancellationToken cancellationToken)
        {
            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    await this.thisTimeDivKeyboardSignal.WaitAsync(cancellationToken).ConfigureAwait(false);

                    while (!cancellationToken.IsCancellationRequested)
                    {
                        int pendingSteps = Interlocked.Exchange(ref this.thisPendingTimeDivKeyboardSteps, 0);
                        if (pendingSteps == 0)
                        {
                            break;
                        }

                        int direction = Math.Sign(pendingSteps);
                        int requestedStepCount = Math.Abs(pendingSteps);
                        OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();

                        await this.RunWithEstablishedOscilloscopeSessionAsync(
                            selectionSnapshot,
                            async (scopeClient, oscilloscopeEntry, token) =>
                            {
                                if (!await this.EnsureCachedTimeDivSecondsAsync(
                                    scopeClient,
                                    oscilloscopeEntry,
                                    token).ConfigureAwait(false))
                                {
                                    this.AppendOutputLine("Warning", "Could not read current TIME/DIV from oscilloscope");
                                    return;
                                }

                                double startingTimeDivSeconds = this.thisLastTimeDivSeconds!.Value;
                                double currentTimeDivSeconds = startingTimeDivSeconds;
                                int appliedStepCount = 0;

                                for (int i = 0; i < requestedStepCount; i++)
                                {
                                    if (!ScopeValueMapper.TryGetAdjacentTimeDivValue(
                                        oscilloscopeEntry,
                                        currentTimeDivSeconds,
                                        direction,
                                        out ScopeMappedValue mappedTimeDiv))
                                    {
                                        if (appliedStepCount == 0)
                                        {
                                            string directionLabel = direction < 0 ? "previous" : "next";
                                            this.AppendOutputLine(
                                                "Warning",
                                                $"Could not resolve the {directionLabel} TIME/DIV value from the oscilloscope definition list");
                                        }

                                        break;
                                    }

                                    await this.SendTimeDivFastAsync(
                                        scopeClient,
                                        oscilloscopeEntry,
                                        mappedTimeDiv.NumericValue,
                                        token).ConfigureAwait(false);

                                    currentTimeDivSeconds = mappedTimeDiv.NumericValue;
                                    appliedStepCount++;
                                }

                                if (appliedStepCount <= 0)
                                {
                                    return;
                                }

                                this.thisLastOscilloscopeImageSyncSignature = string.Empty;

                                this.AppendOutputLine(
                                    "Info",
                                    appliedStepCount == 1
                                        ? $"Keyboard TIME/DIV step: {ScopeFormatting.FormatTime(startingTimeDivSeconds)} -> {ScopeFormatting.FormatTime(currentTimeDivSeconds)}"
                                        : $"Keyboard TIME/DIV step: {ScopeFormatting.FormatTime(startingTimeDivSeconds)} -> {ScopeFormatting.FormatTime(currentTimeDivSeconds)} ({appliedStepCount} steps)");
                            },
                            cancellationToken,
                            writeWarnings: true).ConfigureAwait(false);
                    }
                }
            }
            catch (OperationCanceledException)
            {
            }
        }

        // ###########################################################################################
        // Sends one TIME/DIV Set command through a lightweight SCPI fast path so repeated keyboard
        // stepping avoids the heavier full palette logging pipeline.
        // ###########################################################################################
        private async Task SendTimeDivFastAsync(
            IScopeClient scopeClient,
            OscilloscopeEntry oscilloscopeEntry,
            double targetTimeDivSeconds,
            CancellationToken cancellationToken)
        {
            this.thisLastTimeDivSeconds = targetTimeDivSeconds;

            string commandText = ScopeCommandResolver.GetCommandText(
                oscilloscopeEntry,
                ScopeCommand.SetTimeDiv);

            if (string.IsNullOrWhiteSpace(commandText))
            {
                throw new InvalidOperationException("No SCPI command text is defined for SetTimeDiv");
            }

            string effectiveCommandText = this.BuildEffectiveCommandText(
                ScopeCommand.SetTimeDiv,
                commandText);

            if (string.IsNullOrWhiteSpace(effectiveCommandText))
            {
                throw new InvalidOperationException("No TIME/DIV value is available for SetTimeDiv");
            }

            await scopeClient.SendAsync(effectiveCommandText, cancellationToken).ConfigureAwait(false);
        }

        // ###########################################################################################
        // Queues one keyboard-driven VOLTS/DIV target and starts the dedicated background worker that
        // applies the latest requested value without dropping rapid keypresses.
        // ###########################################################################################
        public void QueueVoltsDivKeyboardSet(double targetVoltsDivVolts)
        {
            if (targetVoltsDivVolts <= 0)
            {
                return;
            }

            lock (this.thisPendingVoltsDivKeyboardLock)
            {
                this.thisPendingVoltsDivKeyboardTargetVolts = targetVoltsDivVolts;
            }

            this.EnsureVoltsDivKeyboardWorkerStarted();

            if (this.thisVoltsDivKeyboardSignal.CurrentCount == 0)
            {
                this.thisVoltsDivKeyboardSignal.Release();
            }
        }

        // ###########################################################################################
        // Starts the dedicated VOLTS/DIV keyboard worker once so rapid fixed-value requests can be
        // handled with latest-wins behavior.
        // ###########################################################################################
        private void EnsureVoltsDivKeyboardWorkerStarted()
        {
            if (this.thisVoltsDivKeyboardWorkerTask != null &&
                !this.thisVoltsDivKeyboardWorkerTask.IsCompleted)
            {
                return;
            }

            this.thisVoltsDivKeyboardWorkerCts?.Cancel();
            this.thisVoltsDivKeyboardWorkerCts?.Dispose();
            this.thisVoltsDivKeyboardWorkerCts = new CancellationTokenSource();
            this.thisVoltsDivKeyboardWorkerTask = this.RunVoltsDivKeyboardWorkerAsync(
                this.thisVoltsDivKeyboardWorkerCts.Token);
        }

        // ###########################################################################################
        // Stops any queued keyboard VOLTS/DIV requests and shuts down the dedicated background worker.
        // ###########################################################################################
        private void StopVoltsDivKeyboardWorker()
        {
            lock (this.thisPendingVoltsDivKeyboardLock)
            {
                this.thisPendingVoltsDivKeyboardTargetVolts = null;
            }

            if (this.thisVoltsDivKeyboardWorkerCts == null)
            {
                return;
            }

            try
            {
                this.thisVoltsDivKeyboardWorkerCts.Cancel();
            }
            catch
            {
                // ObjectDisposedException when the volts/div keyboard worker was already
                // torn down - Cancel() on a disposed CancellationTokenSource throws, and every
                // caller here is a "stop it if it is running" path that has nothing to do
                // differently when it was not. The Dispose() below still runs.
            }

            this.thisVoltsDivKeyboardWorkerCts.Dispose();
            this.thisVoltsDivKeyboardWorkerCts = null;
            this.thisVoltsDivKeyboardWorkerTask = null;
        }

        // ###########################################################################################
        // Drains queued keyboard VOLTS/DIV requests using latest-wins semantics so rapid fixed-value
        // keypresses remain responsive without replaying stale intermediate selections.
        // ###########################################################################################
        private async Task RunVoltsDivKeyboardWorkerAsync(CancellationToken cancellationToken)
        {
            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    await this.thisVoltsDivKeyboardSignal.WaitAsync(cancellationToken).ConfigureAwait(false);

                    double? pendingTargetVoltsDivVolts;
                    lock (this.thisPendingVoltsDivKeyboardLock)
                    {
                        pendingTargetVoltsDivVolts = this.thisPendingVoltsDivKeyboardTargetVolts;
                        this.thisPendingVoltsDivKeyboardTargetVolts = null;
                    }

                    if (!pendingTargetVoltsDivVolts.HasValue)
                    {
                        continue;
                    }

                    OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();

                    await this.RunWithEstablishedOscilloscopeSessionAsync(
                        selectionSnapshot,
                        async (scopeClient, oscilloscopeEntry, token) =>
                        {
                            if (!ScopeValueMapper.TryGetSupportedVoltsDivValue(
                                oscilloscopeEntry,
                                pendingTargetVoltsDivVolts.Value,
                                out ScopeMappedValue mappedVoltsDiv))
                            {
                                this.AppendOutputLine(
                                    "Warning",
                                    $"Could not resolve keyboard VOLTS/DIV value [{ScopeFormatting.FormatVoltage(pendingTargetVoltsDivVolts.Value)}] from the oscilloscope definition list");
                                return;
                            }

                            await this.SendVoltsDivFastAsync(
                                scopeClient,
                                oscilloscopeEntry,
                                mappedVoltsDiv.NumericValue,
                                token).ConfigureAwait(false);

                            this.thisLastOscilloscopeImageSyncSignature = string.Empty;

                            this.AppendOutputLine(
                                "Info",
                                $"Keyboard VOLTS/DIV set: {mappedVoltsDiv.MatchedDisplayValue} per division");
                        },
                        cancellationToken,
                        writeWarnings: true).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException)
            {
            }
        }

        // ###########################################################################################
        // Sends one VOLTS/DIV Set command through a lightweight SCPI fast path so repeated keyboard
        // selections avoid the heavier full palette logging pipeline.
        // ###########################################################################################
        private async Task SendVoltsDivFastAsync(
            IScopeClient scopeClient,
            OscilloscopeEntry oscilloscopeEntry,
            double targetVoltsDivVolts,
            CancellationToken cancellationToken)
        {
            this.thisLastVoltsDivVolts = targetVoltsDivVolts;

            string commandText = ScopeCommandResolver.GetCommandText(
                oscilloscopeEntry,
                ScopeCommand.SetVoltsDiv);

            if (string.IsNullOrWhiteSpace(commandText))
            {
                throw new InvalidOperationException("No SCPI command text is defined for SetVoltsDiv");
            }

            string effectiveCommandText = this.BuildEffectiveCommandText(
                ScopeCommand.SetVoltsDiv,
                commandText);

            if (string.IsNullOrWhiteSpace(effectiveCommandText))
            {
                throw new InvalidOperationException("No VOLTS/DIV value is available for SetVoltsDiv");
            }

            await scopeClient.SendAsync(effectiveCommandText, cancellationToken).ConfigureAwait(false);
        }
    }
}
