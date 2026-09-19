using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System;
using System.Globalization;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // The connect/disconnect flow: building and validating a connection snapshot, opening and
    // tearing down the SCPI session, background connectivity monitoring (ping), auto-reconnect,
    // and the main-window title/session-state updates that follow a connection change. See
    // TabOscilloscope.axaml.cs for the file map of the whole partial class.
    // ###########################################################################################
    public partial class TabOscilloscope
    {
        private CancellationTokenSource? thisOscilloscopeMonitorCancellationTokenSource;
        private string thisMainWindowTitleBase = string.Empty;
        private bool? thisLastOscilloscopeConnectionState;
        private bool thisHasEstablishedOscilloscopeSession;
        private readonly SemaphoreSlim thisOscilloscopeImageSyncSemaphore = new(1, 1);
        private readonly SemaphoreSlim thisOscilloscopeSessionSemaphore = new(1, 1);
        private ScopeScpiClient? thisConnectedScopeClient;
        private bool thisShouldAutoReconnectEstablishedOscilloscopeSession;
        private int thisAutoReconnectAttemptInProgress;
        private DateTime thisLastAutoReconnectAttemptUtc = DateTime.MinValue;
        private bool thisHasSeenEstablishedOscilloscopeSession;
        private CancellationTokenSource? thisOscilloscopeAutoConnectCancellationTokenSource;
        private Task? thisOscilloscopeAutoConnectTask;
        private bool thisHasLoggedAutomaticConnectPendingMessage;

        private Main? thisMainWindow;

        // ###########################################################################################
        // Captures the current oscilloscope UI state on the UI thread so background workers can use
        // the values without reading UI controls directly.
        // ###########################################################################################
        private OscilloscopeSelectionSnapshot CreateOscilloscopeSelectionSnapshot()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                return Dispatcher.UIThread.InvokeAsync(
                    this.CreateOscilloscopeSelectionSnapshot,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
            }

            const int defaultDelayMilliseconds = 250;

            var selectedOscilloscope = this.GetSelectedOscilloscope();
            string host = this.HostTextBox.Text?.Trim() ?? string.Empty;
            int port = int.TryParse(
                this.PortTextBox.Text?.Trim(),
                NumberStyles.None,
                CultureInfo.InvariantCulture,
                out int parsedPort)
                ? parsedPort
                : 0;

            int debounceDelayMilliseconds = defaultDelayMilliseconds;
            if (selectedOscilloscope != null &&
                !string.IsNullOrWhiteSpace(selectedOscilloscope.DebounceTime) &&
                ScopeValueMapper.TryParseTimeValue(selectedOscilloscope.DebounceTime, out double seconds))
            {
                double milliseconds = seconds * 1000.0;
                if (!double.IsNaN(milliseconds) && !double.IsInfinity(milliseconds))
                {
                    debounceDelayMilliseconds = Math.Clamp(
                        (int)Math.Round(milliseconds, MidpointRounding.AwayFromZero),
                        0,
                        60000);
                }
            }

            return new OscilloscopeSelectionSnapshot
            {
                SelectedOscilloscope = selectedOscilloscope,
                Host = host,
                Port = port,
                DebounceDelayMilliseconds = debounceDelayMilliseconds,
                HasActiveEstablishedSession =
                    this.thisHasEstablishedOscilloscopeSession &&
                    this.thisLastOscilloscopeConnectionState == true &&
                    this.thisConnectedScopeClient != null
            };
        }

        // ###########################################################################################
        // Validates the supplied oscilloscope snapshot and returns false when it does not describe a
        // usable oscilloscope target and endpoint.
        // ###########################################################################################
        private bool TryValidateOscilloscopeSelectionSnapshot(
            OscilloscopeSelectionSnapshot selectionSnapshot,
            bool writeWarnings)
        {
            if (selectionSnapshot.SelectedOscilloscope == null)
            {
                if (writeWarnings)
                {
                    this.AppendOutputLine("Warning", "Select a vendor and series or model first");
                }

                return false;
            }

            if (string.IsNullOrWhiteSpace(selectionSnapshot.Host))
            {
                if (writeWarnings)
                {
                    this.AppendOutputLine("Warning", "Enter an IP address or FQDN first");
                }

                return false;
            }

            if (selectionSnapshot.Port < 1 || selectionSnapshot.Port > 65535)
            {
                if (writeWarnings)
                {
                    this.AppendOutputLine("Warning", "TCP port must be within 1-65535");
                }

                return false;
            }

            return true;
        }

        // ###########################################################################################
        // Creates and stores one persistent SCPI client session after an explicit user connect.
        // The connection stays alive and is reused by later popup image auto-sync operations.
        // Automatic reconnect attempts can reuse the same path with different log text.
        // ###########################################################################################
        private async Task<bool> ConnectSelectedOscilloscopeAsync(
            CancellationToken externalCancellationToken,
            bool isAutomaticReconnect = false)
        {
            OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();
            if (!this.TryValidateOscilloscopeSelectionSnapshot(selectionSnapshot, writeWarnings: !isAutomaticReconnect))
            {
                return false;
            }

            bool enteredSemaphore = false;
            ScopeScpiClient? newScopeClient = null;
            OscilloscopeEntry selectedOscilloscope = selectionSnapshot.SelectedOscilloscope!;
            bool isReestablishingSession = this.thisHasSeenEstablishedOscilloscopeSession;

            this.SetOscilloscopeButtonsEnabled(false);

            if (!isAutomaticReconnect)
            {
                this.AppendOutputLine(
                    "Info",
                    $"Connecting to {selectedOscilloscope.Brand} {selectedOscilloscope.SeriesOrModel} at {selectionSnapshot.Host}:{selectionSnapshot.Port}");
            }
            else if (!this.thisHasLoggedAutomaticConnectPendingMessage)
            {
                this.AppendOutputLine(
                    "Info",
                    $"Automatically connecting to {selectedOscilloscope.Brand} {selectedOscilloscope.SeriesOrModel} at {selectionSnapshot.Host}:{selectionSnapshot.Port} ...");
                this.thisHasLoggedAutomaticConnectPendingMessage = true;
            }

            try
            {
                await this.thisOscilloscopeSessionSemaphore.WaitAsync(externalCancellationToken);
                enteredSemaphore = true;

                using var timeoutCts = new CancellationTokenSource(
                    GetOscilloscopeConnectionAttemptTimeout(isAutomaticReconnect));
                using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
                    timeoutCts.Token,
                    externalCancellationToken);

                await this.DisposeConnectedScopeClientCoreAsync();

                newScopeClient = new ScopeScpiClient();
                await newScopeClient.ConnectAsync(
                    selectionSnapshot.Host,
                    selectionSnapshot.Port,
                    linkedCts.Token).ConfigureAwait(false);

                this.AppendOutputLine("Info", "Network session established");

                await this.ExecutePaletteAsync(
                    newScopeClient,
                    selectedOscilloscope,
                    ScopeCommandPalette.Identify,
                    linkedCts.Token).ConfigureAwait(false);

                this.thisConnectedScopeClient = newScopeClient;
                newScopeClient = null;

                this.thisHasEstablishedOscilloscopeSession = true;
                this.thisHasSeenEstablishedOscilloscopeSession = true;
                this.thisShouldAutoReconnectEstablishedOscilloscopeSession = UserSettings.OscilloscopeAutoConnect;
                this.thisLastOscilloscopeConnectionState = true;
                this.thisLastOscilloscopeImageSyncSignature = string.Empty;
                this.thisHasLoggedAutomaticConnectPendingMessage = false;
                this.InvalidateCachedTriggerLevelVolts();
                this.InvalidateCachedTimeDivSeconds();
                this.thisLastVoltsDivVolts = null;

                this.AppendOutputLine(
                    "Info",
                    isAutomaticReconnect
                        ? (isReestablishingSession
                            ? "Oscilloscope session re-established automatically"
                            : "Oscilloscope session established automatically")
                        : "Oscilloscope session established");

                await Dispatcher.UIThread.InvokeAsync(() =>
                {
                    this.StartOscilloscopeConnectionMonitoring();
                    this.UpdateMainWindowOscilloscopeSessionState();
                },
                DispatcherPriority.Background);

                return true;
            }
            catch (OperationCanceledException) when (externalCancellationToken.IsCancellationRequested)
            {
                return false;
            }
            catch (OperationCanceledException)
            {
                if (!isAutomaticReconnect)
                {
                    this.AppendOutputLine("Warning", "Connection or SCPI communication timed out");
                }
            }
            catch (Exception ex)
            {
                if (!isAutomaticReconnect)
                {
                    this.AppendOutputLine("Critical", $"Oscilloscope communication failed: {ex.Message}");
                }
            }
            finally
            {
                if (newScopeClient != null)
                {
                    await newScopeClient.DisposeAsync();
                }

                if (enteredSemaphore)
                {
                    this.thisOscilloscopeSessionSemaphore.Release();
                }

                await Dispatcher.UIThread.InvokeAsync(
                    () => this.SetOscilloscopeButtonsEnabled(true),
                    DispatcherPriority.Background);
            }

            return false;
        }

        // ###########################################################################################
        // Returns the timeout to use for one oscilloscope connection attempt.
        // Automatic background retries use a shorter timeout than manual connects.
        // ###########################################################################################
        private static TimeSpan GetOscilloscopeConnectionAttemptTimeout(bool isAutomaticReconnect)
        {
            return isAutomaticReconnect
                ? TimeSpan.FromSeconds(5)
                : TimeSpan.FromSeconds(30);
        }

        // ###########################################################################################
        // Returns the delay between automatic oscilloscope reconnect attempts.
        // ###########################################################################################
        private static TimeSpan GetOscilloscopeAutoConnectRetryDelay()
        {
            return TimeSpan.FromSeconds(1);
        }

        // ###########################################################################################
        // Runs work against the already established persistent SCPI client without reconnecting.
        // The supplied snapshot is captured on the UI thread so this method can safely run from the
        // background image-sync worker.
        // ###########################################################################################
        private async Task RunWithEstablishedOscilloscopeSessionAsync(
            OscilloscopeSelectionSnapshot selectionSnapshot,
            Func<IScopeClient, OscilloscopeEntry, CancellationToken, Task> runAsync,
            CancellationToken externalCancellationToken,
            bool writeWarnings)
        {
            if (!this.TryValidateOscilloscopeSelectionSnapshot(selectionSnapshot, writeWarnings))
            {
                return;
            }

            bool enteredSemaphore = false;

            try
            {
                await this.thisOscilloscopeSessionSemaphore.WaitAsync(externalCancellationToken).ConfigureAwait(false);
                enteredSemaphore = true;

                if (!selectionSnapshot.HasActiveEstablishedSession || this.thisConnectedScopeClient == null)
                {
                    if (writeWarnings)
                    {
                        this.AppendOutputLine("Warning", "Connect to oscilloscope first");
                    }

                    return;
                }

                await Dispatcher.UIThread.InvokeAsync(
                    () => this.SetOscilloscopeButtonsEnabled(false),
                    DispatcherPriority.Background);

                using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
                    timeoutCts.Token,
                    externalCancellationToken);

                await runAsync(
                    this.thisConnectedScopeClient,
                    selectionSnapshot.SelectedOscilloscope!,
                    linkedCts.Token).ConfigureAwait(false);
            }
            catch (OperationCanceledException) when (externalCancellationToken.IsCancellationRequested)
            {
            }
            catch (OperationCanceledException)
            {
                await this.HandleOscilloscopeSessionFailureCoreAsync("Connection or SCPI communication timed out");
            }
            catch (Exception ex)
            {
                await this.HandleOscilloscopeSessionFailureCoreAsync(ex.Message);
            }
            finally
            {
                if (enteredSemaphore)
                {
                    this.thisOscilloscopeSessionSemaphore.Release();
                }

                await Dispatcher.UIThread.InvokeAsync(
                    () => this.SetOscilloscopeButtonsEnabled(true),
                    DispatcherPriority.Background);
            }
        }

        // ###########################################################################################
        // Runs work against the already established persistent SCPI client without reconnecting.
        // Used by direct UI actions that are initiated from the oscilloscope tab itself.
        // ###########################################################################################
        private async Task RunWithEstablishedOscilloscopeSessionAsync(
            Func<IScopeClient, OscilloscopeEntry, CancellationToken, Task> runAsync,
            CancellationToken externalCancellationToken)
        {
            OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();

            await this.RunWithEstablishedOscilloscopeSessionAsync(
                selectionSnapshot,
                runAsync,
                externalCancellationToken,
                writeWarnings: true);
        }

        // ###########################################################################################
        // Connects to the selected oscilloscope and runs the Identify command palette over SCPI.
        // Starts background connectivity monitoring only after a successful connection.
        // ###########################################################################################
        private async void OnConnectToOscilloscopeClick(object? sender, RoutedEventArgs e)
        {
            await this.ConnectSelectedOscilloscopeAsync(
                CancellationToken.None,
                isAutomaticReconnect: false);
        }

        // ###########################################################################################
        // Starts or restarts background ICMP monitoring for the currently configured oscilloscope host.
        // ###########################################################################################
        private void StartOscilloscopeConnectionMonitoring()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    this.StartOscilloscopeConnectionMonitoring,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            string host = this.HostTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(host))
            {
                return;
            }

            this.StopOscilloscopeConnectionMonitoring();

            Window? window = this.thisMainWindow ?? TopLevel.GetTopLevel(this) as Window;
            if (window != null)
            {
                this.thisMainWindowTitleBase = ScopeFormatting.GetMainWindowTitleBase(window.Title ?? string.Empty);
            }

            this.thisLastOscilloscopeConnectionState = null;
            this.thisOscilloscopeMonitorCancellationTokenSource = new CancellationTokenSource();

            _ = this.RunOscilloscopeConnectionMonitoringAsync(
                host,
                this.thisOscilloscopeMonitorCancellationTokenSource.Token);
        }

        // ###########################################################################################
        // Stops any existing background oscilloscope connectivity monitor.
        // ###########################################################################################
        private void StopOscilloscopeConnectionMonitoring()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    this.StopOscilloscopeConnectionMonitoring,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            if (this.thisOscilloscopeMonitorCancellationTokenSource != null)
            {
                try
                {
                    this.thisOscilloscopeMonitorCancellationTokenSource.Cancel();
                }
                catch
                {
                    // ObjectDisposedException when the connectivity monitor was already
                    // torn down - Cancel() on a disposed CancellationTokenSource throws, and every
                    // caller here is a "stop it if it is running" path that has nothing to do
                    // differently when it was not. The Dispose() below still runs, via the
                    // statement after this block.
                }

                this.thisOscilloscopeMonitorCancellationTokenSource.Dispose();
                this.thisOscilloscopeMonitorCancellationTokenSource = null;
            }

            this.thisLastOscilloscopeConnectionState = null;
            this.thisLastOscilloscopeImageSyncSignature = string.Empty;
            this.UpdateMainWindowOscilloscopeSessionState();
        }

        // ###########################################################################################
        // Repeatedly pings the oscilloscope host, updates the main window title on state changes,
        // and triggers automatic reconnect attempts when the host becomes reachable again.
        // ###########################################################################################
        private async Task RunOscilloscopeConnectionMonitoringAsync(string host, CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                bool isConnected = await this.PingOscilloscopeAsync(host);

                await Dispatcher.UIThread.InvokeAsync(() =>
                {
                    this.UpdateMainWindowOscilloscopeConnectionState(isConnected);

                    if (isConnected)
                    {
                        this.QueueAutomaticOscilloscopeReconnectIfNeeded();
                    }
                },
                DispatcherPriority.Background);

                try
                {
                    await Task.Delay(500, cancellationToken);
                }
                catch (OperationCanceledException)
                {
                    break;
                }
            }
        }

        // ###########################################################################################
        // Performs a simple ICMP ping against the oscilloscope host.
        // ###########################################################################################
        private async Task<bool> PingOscilloscopeAsync(string host)
        {
            try
            {
                using var ping = new Ping();
                var reply = await ping.SendPingAsync(host, 300);
                return reply.Status == IPStatus.Success;
            }
            catch
            {
                return false;
            }
        }

        // ###########################################################################################
        // Updates ping reachability state for the oscilloscope.
        // Ping state now drives logging and reconnect behavior, while the main window title reflects
        // the actual SCPI session state through UpdateMainWindowOscilloscopeSessionState().
        // ###########################################################################################
        private void UpdateMainWindowOscilloscopeConnectionState(bool isConnected)
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    () => this.UpdateMainWindowOscilloscopeConnectionState(isConnected),
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            bool? previousConnectionState = this.thisLastOscilloscopeConnectionState;

            if (previousConnectionState == isConnected)
            {
                return;
            }

            if (previousConnectionState == true && !isConnected)
            {
                this.AppendOutputLine("Warning", "Oscilloscope disconnected");
            }
            else if (previousConnectionState == false && isConnected)
            {
                this.AppendOutputLine("Info", "Oscilloscope reachable again");
            }

            this.thisLastOscilloscopeConnectionState = isConnected;

            if (!isConnected)
            {
                if (this.thisHasEstablishedOscilloscopeSession || this.thisConnectedScopeClient != null)
                {
                    this.thisHasEstablishedOscilloscopeSession = false;
                    this.thisLastOscilloscopeImageSyncSignature = string.Empty;
                    this.thisHasLoggedAutomaticConnectPendingMessage = false;
                    this.InvalidateCachedTriggerLevelVolts();
                    this.StopTriggerLevelKeyboardWorker();
                    _ = this.DisposeConnectedScopeClientAsync();
                }
            }

            this.UpdateMainWindowOscilloscopeSessionState();
        }

        // ###########################################################################################
        // Invalidates the explicit user-established oscilloscope session when the selected device or
        // endpoint changes, so popup image changes cannot auto-connect implicitly.
        // ###########################################################################################
        private async void InvalidateEstablishedOscilloscopeSession()
        {
            this.thisHasEstablishedOscilloscopeSession = false;
            this.thisHasSeenEstablishedOscilloscopeSession = false;
            this.thisShouldAutoReconnectEstablishedOscilloscopeSession = false;
            this.thisLastOscilloscopeImageSyncSignature = string.Empty;
            this.thisHasLoggedAutomaticConnectPendingMessage = false;
            this.InvalidateCachedTriggerLevelVolts();
            this.InvalidateCachedTimeDivSeconds();
            this.thisLastVoltsDivVolts = null;
            this.StopTriggerLevelKeyboardWorker();
            this.StopTimeDivKeyboardWorker();
            this.StopVoltsDivKeyboardWorker();
            Interlocked.Exchange(ref this.thisAutoReconnectAttemptInProgress, 0);
            this.StopOscilloscopeConnectionMonitoring();
            await this.DisposeConnectedScopeClientAsync();
        }

        // ###########################################################################################
        // Disposes the currently stored persistent SCPI client while serializing against other scope
        // operations so the connection cannot be torn down mid-command.
        // ###########################################################################################
        private async Task DisposeConnectedScopeClientAsync()
        {
            bool enteredSemaphore = false;

            try
            {
                await this.thisOscilloscopeSessionSemaphore.WaitAsync();
                enteredSemaphore = true;
                await this.DisposeConnectedScopeClientCoreAsync();
            }
            catch (Exception ex)
            {
                // Tearing a connection down is best-effort: the socket is being abandoned either
                // way, and throwing here would take out whatever is disconnecting (a tab switch, a
                // host change, application shutdown) over a scope that is already gone.
                //
                // LOGGED rather than swallowed silently, because this try covers the semaphore WAIT
                // as well as the dispose. A failure to acquire the session semaphore is not a
                // teardown race - it means scope commands are no longer being serialized, which
                // presents to the user only as the oscilloscope going unresponsive, with nothing
                // anywhere to explain it. The two cases are indistinguishable without this line.
                Logger.Warning($"Failed to dispose the oscilloscope connection cleanly: [{ex.Message}]");
            }
            finally
            {
                if (enteredSemaphore)
                {
                    this.thisOscilloscopeSessionSemaphore.Release();
                }
            }
        }

        // ###########################################################################################
        // Disposes the stored persistent SCPI client. Caller must already hold the session semaphore.
        // ###########################################################################################
        private async Task DisposeConnectedScopeClientCoreAsync()
        {
            if (this.thisConnectedScopeClient == null)
            {
                return;
            }

            try
            {
                await this.thisConnectedScopeClient.DisposeAsync();
            }
            catch
            {
                // Deliberately silent, and the one bare catch here that should stay that way: the
                // caller (DisposeConnectedScopeClientAsync) already logs, and a scope that dropped
                // the link mid-session throws here on EVERY disconnect - logging it again would put
                // a warning in the log for the ordinary case of unplugging the instrument.
            }
            finally
            {
                this.thisConnectedScopeClient = null;
            }
        }

        // ###########################################################################################
        // Tears down the persistent session after a communication failure and reports the failure in
        // the output pane. Connectivity monitoring remains active so automatic reconnect can occur
        // once the oscilloscope becomes reachable again.
        // ###########################################################################################
        private async Task HandleOscilloscopeSessionFailureCoreAsync(string message)
        {
            this.AppendOutputLine("Critical", $"Oscilloscope communication failed: {message}");
            this.thisHasEstablishedOscilloscopeSession = false;
            this.thisLastOscilloscopeImageSyncSignature = string.Empty;
            this.thisHasLoggedAutomaticConnectPendingMessage = false;
            this.InvalidateCachedTriggerLevelVolts();
            this.InvalidateCachedTimeDivSeconds();
            this.thisLastVoltsDivVolts = null;
            this.StopTriggerLevelKeyboardWorker();
            this.StopTimeDivKeyboardWorker();
            this.StopVoltsDivKeyboardWorker();
            await this.DisposeConnectedScopeClientCoreAsync();
            this.UpdateMainWindowOscilloscopeSessionState();
        }

        // ###########################################################################################
        // Immutable snapshot of the current oscilloscope UI state, captured on the UI thread and
        // later consumed by the background image-sync worker without touching UI controls.
        // ###########################################################################################
        private sealed class OscilloscopeSelectionSnapshot
        {
            public OscilloscopeEntry? SelectedOscilloscope { get; init; }
            public string Host { get; init; } = string.Empty;
            public int Port { get; init; }
            public int DebounceDelayMilliseconds { get; init; }
            public bool HasActiveEstablishedSession { get; init; }
        }

        // ###########################################################################################
        // Queues an automatic reconnect attempt when the oscilloscope has previously been connected,
        // is reachable again over ping, and no active SCPI session currently exists.
        // ###########################################################################################
        private void QueueAutomaticOscilloscopeReconnectIfNeeded()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    this.QueueAutomaticOscilloscopeReconnectIfNeeded,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            if (UserSettings.OscilloscopeAutoConnect)
            {
                return;
            }

            if (!UserSettings.EnableNetworkConnectedOscilloscopeTab)
            {
                return;
            }

            if (!this.thisShouldAutoReconnectEstablishedOscilloscopeSession)
            {
                return;
            }

            if (this.thisConnectedScopeClient != null || this.thisHasEstablishedOscilloscopeSession)
            {
                return;
            }

            if (DateTime.UtcNow - this.thisLastAutoReconnectAttemptUtc < TimeSpan.FromSeconds(3))
            {
                return;
            }

            if (Interlocked.CompareExchange(ref this.thisAutoReconnectAttemptInProgress, 1, 0) != 0)
            {
                return;
            }

            this.thisLastAutoReconnectAttemptUtc = DateTime.UtcNow;
            this.AppendOutputLine("Info", "Oscilloscope reachable - attempting automatic reconnect");

            _ = this.RunAutomaticOscilloscopeReconnectAsync();
        }

        // ###########################################################################################
        // Runs one automatic reconnect attempt and releases the reconnect gate afterward so future
        // retries can occur if the oscilloscope is still not ready for SCPI connections.
        // ###########################################################################################
        private async Task RunAutomaticOscilloscopeReconnectAsync()
        {
            try
            {
                await this.ConnectSelectedOscilloscopeAsync(
                    CancellationToken.None,
                    isAutomaticReconnect: true);
            }
            finally
            {
                Interlocked.Exchange(ref this.thisAutoReconnectAttemptInProgress, 0);
            }
        }

        // ###########################################################################################
        // Updates the main window title and button state based on the actual SCPI session state.
        // The oscilloscope suffix is only shown after a session has existed, or while auto-connect
        // is enabled and startup/background detection is expected to be active.
        // ###########################################################################################
        private void UpdateMainWindowOscilloscopeSessionState()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    this.UpdateMainWindowOscilloscopeSessionState,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            bool hasEstablishedSession =
                this.thisHasEstablishedOscilloscopeSession &&
                this.thisConnectedScopeClient != null;

            // Auto-connect alone is enough to report a state, so that a user who expects the scope to
            // come up on its own can see it is still pending.
            bool shouldReportSessionState =
                this.thisHasSeenEstablishedOscilloscopeSession ||
                UserSettings.OscilloscopeAutoConnect;

            this.RunFullTestSuiteButton.IsEnabled = hasEstablishedSession;

            Window? window = this.thisMainWindow ?? TopLevel.GetTopLevel(this) as Window;
            if (window != null)
            {
                if (string.IsNullOrWhiteSpace(this.thisMainWindowTitleBase))
                {
                    this.thisMainWindowTitleBase = ScopeFormatting.GetMainWindowTitleBase(window.Title ?? string.Empty);
                }

                string newTitle = ScopeFormatting.BuildOscilloscopeWindowTitle(
                    this.thisMainWindowTitleBase,
                    UserSettings.EnableNetworkConnectedOscilloscopeTab,
                    shouldReportSessionState,
                    hasEstablishedSession);

                if (!string.Equals(window.Title, newTitle, StringComparison.Ordinal))
                {
                    window.Title = newTitle;
                }
            }

            this.NotifyOpenComponentInfoWindowsOfOscilloscopeSessionState();
        }

        // ###########################################################################################
        // Returns true once the oscilloscope has had an established session for the current target,
        // allowing other windows to decide whether a connection-status suffix should be shown.
        // ###########################################################################################
        public bool HasSeenEstablishedOscilloscopeSessionForTitleState()
        {
            return this.thisHasSeenEstablishedOscilloscopeSession;
        }

        // ###########################################################################################
        // Returns true only when an active established SCPI session currently exists so other
        // windows can mirror the same connected/disconnected title state.
        // ###########################################################################################
        public bool HasActiveEstablishedOscilloscopeSessionForTitleState()
        {
            return this.thisHasEstablishedOscilloscopeSession &&
                   this.thisConnectedScopeClient != null;
        }

        // ###########################################################################################
        // Pushes the current oscilloscope session title state into any open component info popup
        // windows owned by the main window.
        // ###########################################################################################
        private void NotifyOpenComponentInfoWindowsOfOscilloscopeSessionState()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    this.NotifyOpenComponentInfoWindowsOfOscilloscopeSessionState,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            Main? mainWindow = this.thisMainWindow ?? TopLevel.GetTopLevel(this) as Main;
            if (mainWindow != null)
            {
                mainWindow.UpdateComponentInfoWindowsOscilloscopeSessionState(
                    this.HasSeenEstablishedOscilloscopeSessionForTitleState(),
                    this.HasActiveEstablishedOscilloscopeSessionForTitleState());
            }
        }

        // ###########################################################################################
        // Persists the auto-connect checkbox state and starts or stops the background auto-connect loop.
        // ###########################################################################################
        private void OnAutoConnectOscilloscopeCheckBoxChanged()
        {
            bool isEnabled = this.AutoConnectOscilloscopeCheckBox.IsChecked == true;

            UserSettings.OscilloscopeAutoConnect = isEnabled;
            this.thisShouldAutoReconnectEstablishedOscilloscopeSession = isEnabled;

            if (isEnabled && UserSettings.EnableNetworkConnectedOscilloscopeTab)
            {
                this.StartOscilloscopeAutoConnectLoop();
            }
            else
            {
                this.StopOscilloscopeAutoConnectLoop();
            }

            this.UpdateMainWindowOscilloscopeSessionState();
        }

        // ###########################################################################################
        // Starts one background loop that continuously retries oscilloscope connection while enabled.
        // ###########################################################################################
        private void StartOscilloscopeAutoConnectLoop()
        {
            if (this.thisOscilloscopeAutoConnectTask != null &&
                !this.thisOscilloscopeAutoConnectTask.IsCompleted)
            {
                return;
            }

            this.thisOscilloscopeAutoConnectCancellationTokenSource?.Cancel();
            this.thisOscilloscopeAutoConnectCancellationTokenSource?.Dispose();
            this.thisOscilloscopeAutoConnectCancellationTokenSource = new CancellationTokenSource();
            this.thisOscilloscopeAutoConnectTask = this.RunOscilloscopeAutoConnectLoopAsync(
                this.thisOscilloscopeAutoConnectCancellationTokenSource.Token);
        }

        // ###########################################################################################
        // Stops the background oscilloscope auto-connect retry loop.
        // ###########################################################################################
        private void StopOscilloscopeAutoConnectLoop()
        {
            if (this.thisOscilloscopeAutoConnectCancellationTokenSource == null)
            {
                this.thisHasLoggedAutomaticConnectPendingMessage = false;
                return;
            }

            try
            {
                this.thisOscilloscopeAutoConnectCancellationTokenSource.Cancel();
            }
            catch
            {
                // ObjectDisposedException when the auto-connect retry loop was already
                // torn down - Cancel() on a disposed CancellationTokenSource throws, and every
                // caller here is a "stop it if it is running" path that has nothing to do
                // differently when it was not. The Dispose() below still runs.
            }

            this.thisOscilloscopeAutoConnectCancellationTokenSource.Dispose();
            this.thisOscilloscopeAutoConnectCancellationTokenSource = null;
            this.thisOscilloscopeAutoConnectTask = null;
            this.thisHasLoggedAutomaticConnectPendingMessage = false;
        }

        // ###########################################################################################
        // Continuously retries oscilloscope connection while auto-connect is enabled and no active
        // established session currently exists.
        // ###########################################################################################
        private async Task RunOscilloscopeAutoConnectLoopAsync(CancellationToken cancellationToken)
        {
            try
            {
                while (!cancellationToken.IsCancellationRequested)
                {
                    OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();

                    if (!UserSettings.OscilloscopeAutoConnect ||
                        !UserSettings.EnableNetworkConnectedOscilloscopeTab)
                    {
                        await Task.Delay(
                            GetOscilloscopeAutoConnectRetryDelay(),
                            cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    if (selectionSnapshot.HasActiveEstablishedSession ||
                        this.thisHasEstablishedOscilloscopeSession ||
                        this.thisConnectedScopeClient != null)
                    {
                        await Task.Delay(
                            GetOscilloscopeAutoConnectRetryDelay(),
                            cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    if (!this.TryValidateOscilloscopeSelectionSnapshot(selectionSnapshot, writeWarnings: false))
                    {
                        await Task.Delay(
                            GetOscilloscopeAutoConnectRetryDelay(),
                            cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    if (Interlocked.CompareExchange(ref this.thisAutoReconnectAttemptInProgress, 1, 0) != 0)
                    {
                        await Task.Delay(500, cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    try
                    {
                        this.thisLastAutoReconnectAttemptUtc = DateTime.UtcNow;

                        await this.ConnectSelectedOscilloscopeAsync(
                            cancellationToken,
                            isAutomaticReconnect: true).ConfigureAwait(false);
                    }
                    finally
                    {
                        Interlocked.Exchange(ref this.thisAutoReconnectAttemptInProgress, 0);
                    }

                    await Task.Delay(
                        GetOscilloscopeAutoConnectRetryDelay(),
                        cancellationToken).ConfigureAwait(false);
                }
            }
            catch (OperationCanceledException)
            {
            }
        }

        // ###########################################################################################
        // Initializes the oscilloscope tab against the main window so background auto-connect and
        // title updates continue even when the Oscilloscope tab is not the currently selected tab.
        // ###########################################################################################
        public void InitializeForMainWindow(Main mainWindow)
        {
            this.thisMainWindow = mainWindow;
            this.UpdateMainWindowOscilloscopeSessionState();

            if (UserSettings.OscilloscopeAutoConnect &&
                UserSettings.EnableNetworkConnectedOscilloscopeTab)
            {
                this.thisShouldAutoReconnectEstablishedOscilloscopeSession = true;
                this.StartOscilloscopeAutoConnectLoop();
            }
        }

        // ###########################################################################################
        // Applies the "Enable network connected oscilloscope tab" setting after the user has toggled
        // it. Hiding the tab must not leave oscilloscope work running behind it, so the retry loop is
        // stopped and any established session is invalidated - which also drops the ping monitor, the
        // keyboard workers and the stored SCPI client, and clears the flag the window titles read.
        //
        // Switching the tab back on restarts auto-connect if the user had it enabled. The session is
        // not restored: it was torn down, so the scope reconnects the same way it would after any
        // other disconnect.
        // ###########################################################################################
        public void ApplyOscilloscopeTabAvailability()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    this.ApplyOscilloscopeTabAvailability,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            // Before InitializeForMainWindow has run there is nothing started yet to stop, and that
            // method applies the same setting itself - so leave the startup path to do it once.
            if (this.thisMainWindow == null)
            {
                return;
            }

            if (!UserSettings.EnableNetworkConnectedOscilloscopeTab)
            {
                this.StopOscilloscopeAutoConnectLoop();
                this.InvalidateEstablishedOscilloscopeSession();
                this.UpdateMainWindowOscilloscopeSessionState();
                return;
            }

            if (UserSettings.OscilloscopeAutoConnect)
            {
                this.thisShouldAutoReconnectEstablishedOscilloscopeSession = true;
                this.StartOscilloscopeAutoConnectLoop();
            }

            this.UpdateMainWindowOscilloscopeSessionState();
        }
    }
}
