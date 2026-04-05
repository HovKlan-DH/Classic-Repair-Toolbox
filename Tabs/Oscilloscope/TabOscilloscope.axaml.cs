using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Avalonia.Platform.Storage;
using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Globalization;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;
using System.Diagnostics.CodeAnalysis;

namespace CRT
{
    public partial class TabOscilloscope : UserControl
    {
        private double? thisLastTriggerLevelVolts;
        private double? thisLastTimeDivSeconds;
        private double? thisLastVoltsDivVolts;

        private bool thisIsNormalizingPortText;
        private string thisLastValidPortText = "5025";

        private CancellationTokenSource? thisOscilloscopeMonitorCancellationTokenSource;
        private string thisMainWindowTitleBase = string.Empty;
        private bool? thisLastOscilloscopeConnectionState;
        private bool thisHasEstablishedOscilloscopeSession;
        private readonly SemaphoreSlim thisOscilloscopeImageSyncSemaphore = new(1, 1);
        private readonly SemaphoreSlim thisOscilloscopeSessionSemaphore = new(1, 1);
        private ScopeScpiClient? thisConnectedScopeClient;
        private string thisLastOscilloscopeImageSyncSignature = string.Empty;
        private readonly object thisPendingImageSyncLock = new();
        private PendingComponentImageSyncRequest? thisPendingComponentImageSyncRequest;
        private readonly SemaphoreSlim thisImageSyncSignal = new(0);
        private CancellationTokenSource? thisImageSyncWorkerCts;
        private Task? thisImageSyncWorkerTask;
        private readonly object thisPendingOutputLinesLock = new();
        private readonly List<string> thisPendingOutputLines = new();
        private bool thisOutputFlushScheduled;
        private bool thisShouldAutoReconnectEstablishedOscilloscopeSession;
        private int thisAutoReconnectAttemptInProgress;
        private DateTime thisLastAutoReconnectAttemptUtc = DateTime.MinValue;
        private bool thisHasSeenEstablishedOscilloscopeSession;
        private CancellationTokenSource? thisOscilloscopeAutoConnectCancellationTokenSource;
        private Task? thisOscilloscopeAutoConnectTask;
        private bool thisHasLoggedAutomaticConnectPendingMessage;

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

        private Main? thisMainWindow;

        public TabOscilloscope()
        {
            this.InitializeComponent();

            this.RunFullTestSuiteButton.IsEnabled = false;

            this.HostTextBox.Text = UserSettings.OscilloscopeHost;
            this.HostTextBox.TextChanged += this.OnHostTextChanged;

            this.PortTextBox.Text = (UserSettings.OscilloscopePort is >= 1 and <= 65535
                ? UserSettings.OscilloscopePort
                : 5025).ToString(CultureInfo.InvariantCulture);
            this.thisLastValidPortText = this.PortTextBox.Text;
            this.PortTextBox.TextChanged += this.OnPortTextChanged;

            this.AutoConnectOscilloscopeCheckBox.IsChecked = UserSettings.OscilloscopeAutoConnect;
            this.AutoConnectOscilloscopeCheckBox.IsCheckedChanged += (_, _) => this.OnAutoConnectOscilloscopeCheckBoxChanged();

            this.VendorComboBox.SelectionChanged += this.OnVendorSelectionChanged;
            this.SeriesOrModelComboBox.SelectionChanged += this.OnSeriesOrModelSelectionChanged;

            this.OscilloscopeImageFolderTextBox.Text = UserSettings.OscilloscopeImageFolder;
            this.UpdateOscilloscopeImageFolderUi();

            this.PopulateVendorDropDown();
        }

        // ###########################################################################################
        // Populates the vendor drop-down with distinct brand names from loaded data.
        // ###########################################################################################
        private void PopulateVendorDropDown()
        {
            var brandNames = DataManager.Oscilloscopes
                .Select(e => e.Brand)
                .Where(b => !string.IsNullOrWhiteSpace(b))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            this.VendorComboBox.ItemsSource = brandNames;

            if (brandNames.Count == 0)
            {
                this.VendorComboBox.SelectedIndex = -1;
                return;
            }

            var lastVendor = UserSettings.GetLastOscilloscopeVendor();
            var savedIndex = brandNames.FindIndex(v =>
                string.Equals(v, lastVendor, StringComparison.OrdinalIgnoreCase));

            this.VendorComboBox.SelectedIndex = savedIndex >= 0 ? savedIndex : 0;
        }

        // ###########################################################################################
        // Filters the series/model drop-down to only show items belonging to the selected vendor.
        // ###########################################################################################
        private void OnVendorSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            this.InvalidateEstablishedOscilloscopeSession();

            var selectedVendor = this.VendorComboBox.SelectedItem as string;

            var models = DataManager.Oscilloscopes
                .Where(entry => string.Equals(entry.Brand, selectedVendor, StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.SeriesOrModel)
                .Where(m => !string.IsNullOrWhiteSpace(m))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            this.SeriesOrModelComboBox.ItemsSource = models;

            if (string.IsNullOrWhiteSpace(selectedVendor) || models.Count == 0)
            {
                this.SeriesOrModelComboBox.SelectedIndex = -1;
                return;
            }

            UserSettings.SetLastOscilloscopeVendor(selectedVendor);

            var lastSeries = UserSettings.GetLastOscilloscopeSeriesForVendor(selectedVendor);
            var savedIndex = models.FindIndex(m =>
                string.Equals(m, lastSeries, StringComparison.OrdinalIgnoreCase));

            this.SeriesOrModelComboBox.SelectedIndex = savedIndex >= 0 ? savedIndex : 0;
        }

        // ###########################################################################################
        // Persists the newly selected series/model selection in user settings.
        // ###########################################################################################
        private void OnSeriesOrModelSelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            this.InvalidateEstablishedOscilloscopeSession();

            var selectedVendor = this.VendorComboBox.SelectedItem as string;
            var selectedSeries = this.SeriesOrModelComboBox.SelectedItem as string;

            if (string.IsNullOrWhiteSpace(selectedVendor) || string.IsNullOrWhiteSpace(selectedSeries))
            {
                return;
            }

            UserSettings.SetLastOscilloscopeSeriesForVendor(selectedVendor, selectedSeries);

            var selectedOscilloscope = DataManager.Oscilloscopes.FirstOrDefault(entry =>
                string.Equals(entry.Brand, selectedVendor, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(entry.SeriesOrModel, selectedSeries, StringComparison.OrdinalIgnoreCase));

            if (selectedOscilloscope == null)
            {
                return;
            }

            if (int.TryParse(selectedOscilloscope.Port, out int port) &&
                port >= 1 &&
                port <= 65535)
            {
                this.PortTextBox.Text = port.ToString(CultureInfo.InvariantCulture);
            }
        }

        // ###########################################################################################
        // Persists IP / Hostname field whenever it's updated.
        // ###########################################################################################
        private void OnHostTextChanged(object? sender, global::Avalonia.Controls.TextChangedEventArgs e)
        {
            UserSettings.OscilloscopeHost = this.HostTextBox.Text ?? string.Empty;
            this.InvalidateEstablishedOscilloscopeSession();
        }

        // ###########################################################################################
        // Persists the TCP port whenever the textbox contains a valid port value and normalizes any
        // invalid edits so only ASCII digits 0-9 remain while keeping the value within 1-65535.
        // ###########################################################################################
        private void OnPortTextChanged(object? sender, global::Avalonia.Controls.TextChangedEventArgs e)
        {
            if (this.thisIsNormalizingPortText)
            {
                return;
            }

            string originalText = this.PortTextBox.Text ?? string.Empty;
            string sanitizedText = new(originalText.Where(ch => ch is >= '0' and <= '9').Take(5).ToArray());

            if (sanitizedText.Length > 0 &&
                int.TryParse(sanitizedText, NumberStyles.None, CultureInfo.InvariantCulture, out int typedPort) &&
                typedPort > 65535)
            {
                sanitizedText = this.thisLastValidPortText;
            }

            if (!string.Equals(originalText, sanitizedText, StringComparison.Ordinal))
            {
                int caretIndex = this.PortTextBox.CaretIndex;

                this.thisIsNormalizingPortText = true;

                try
                {
                    this.PortTextBox.Text = sanitizedText;
                    this.PortTextBox.CaretIndex = Math.Min(caretIndex, sanitizedText.Length);
                }
                finally
                {
                    this.thisIsNormalizingPortText = false;
                }
            }

            if (int.TryParse(sanitizedText, NumberStyles.None, CultureInfo.InvariantCulture, out int port) &&
                port >= 1 &&
                port <= 65535)
            {
                UserSettings.OscilloscopePort = port;
                this.thisLastValidPortText = sanitizedText;
            }

            this.InvalidateEstablishedOscilloscopeSession();
        }

        // ###########################################################################################
        // Rejects non-digit text input in the TCP port field and blocks typed edits that would make
        // the resulting port value exceed 65535. The TextBox MaxLength still enforces 5 characters.
        // ###########################################################################################
        private void OnPortTextBoxTextInput(object? sender, TextInputEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Text))
            {
                return;
            }

            if (e.Text.Any(ch => ch is < '0' or > '9'))
            {
                e.Handled = true;
                return;
            }

            if (sender is not TextBox textBox)
            {
                return;
            }

            string existingText = textBox.Text ?? string.Empty;
            int selectionStart = textBox.SelectionStart;
            int selectionEnd = textBox.SelectionEnd;

            if (selectionEnd < selectionStart)
            {
                (selectionStart, selectionEnd) = (selectionEnd, selectionStart);
            }

            string resultingText =
                existingText[..selectionStart] +
                e.Text +
                existingText[selectionEnd..];

            if (resultingText.Length > 5)
            {
                e.Handled = true;
                return;
            }

            if (int.TryParse(resultingText, NumberStyles.None, CultureInfo.InvariantCulture, out int port) &&
                port > 65535)
            {
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Restores a valid TCP port value when the field loses focus.
        // ###########################################################################################
        private void OnPortTextBoxLostFocus(object? sender, RoutedEventArgs e)
        {
            string text = this.PortTextBox.Text?.Trim() ?? string.Empty;

            if (string.IsNullOrWhiteSpace(text))
            {
                int fallbackPort = UserSettings.OscilloscopePort is >= 1 and <= 65535
                    ? UserSettings.OscilloscopePort
                    : 5025;

                this.PortTextBox.Text = fallbackPort.ToString(CultureInfo.InvariantCulture);
                return;
            }

            if (!int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int port))
            {
                int fallbackPort = UserSettings.OscilloscopePort is >= 1 and <= 65535
                    ? UserSettings.OscilloscopePort
                    : 5025;

                this.PortTextBox.Text = fallbackPort.ToString(CultureInfo.InvariantCulture);
                return;
            }

            if (port < 1)
            {
                port = 1;
            }
            else if (port > 65535)
            {
                port = 65535;
            }

            this.PortTextBox.Text = port.ToString(CultureInfo.InvariantCulture);
            UserSettings.OscilloscopePort = port;
        }

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
//                this.AppendOutputLine("Debug", "---");
                this.AppendOutputLine(
                    "Info",
                    $"Connecting to {selectedOscilloscope.Brand} {selectedOscilloscope.SeriesOrModel} at {selectionSnapshot.Host}:{selectionSnapshot.Port}");
            }
            else if (!this.thisHasLoggedAutomaticConnectPendingMessage)
            {
//                this.AppendOutputLine("Debug", "---");
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
                    this.GetOscilloscopeConnectionAttemptTimeout(isAutomaticReconnect));
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
        private TimeSpan GetOscilloscopeConnectionAttemptTimeout(bool isAutomaticReconnect)
        {
            return isAutomaticReconnect
                ? TimeSpan.FromSeconds(5)
                : TimeSpan.FromSeconds(30);
        }

        // ###########################################################################################
        // Returns the delay between automatic oscilloscope reconnect attempts.
        // ###########################################################################################
        private TimeSpan GetOscilloscopeAutoConnectRetryDelay()
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
            Func<ScopeScpiClient, OscilloscopeEntry, CancellationToken, Task> runAsync,
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
            Func<ScopeScpiClient, OscilloscopeEntry, CancellationToken, Task> runAsync,
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
        // Executes the full palette sequence against the already established persistent session.
        // Dump image includes the same preparation steps shown in the sample output.
        // ###########################################################################################
        private async void OnRunFullTestSuiteClick(object? sender, RoutedEventArgs e)
        {
            await this.RunWithEstablishedOscilloscopeSessionAsync(async (scopeClient, selectedOscilloscope, cancellationToken) =>
            {
                this.thisLastTriggerLevelVolts = null;
                this.thisLastTimeDivSeconds = null;
                this.thisLastVoltsDivVolts = null;

                foreach (var palette in ScopeCommandPaletteDefinitions.GetFullCommandPaletteExecutionOrder())
                {
                    if (palette == ScopeCommandPalette.DumpImage)
                    {
                        await this.ExecuteDumpImageWorkflowAsync(
                            scopeClient,
                            selectedOscilloscope,
                            cancellationToken);
                        continue;
                    }

                    await this.ExecutePaletteAsync(
                        scopeClient,
                        selectedOscilloscope,
                        palette,
                        cancellationToken);
                }
            },
            CancellationToken.None);
        }

        // ###########################################################################################
        // Executes a named scope command palette and logs every sent and received SCPI command.
        // ###########################################################################################
        private async Task ExecutePaletteAsync(
            ScopeScpiClient scopeClient,
            OscilloscopeEntry selectedOscilloscope,
            ScopeCommandPalette palette,
            CancellationToken cancellationToken)
        {
            this.AppendOutputLine("Debug", "---");
            this.AppendOutputLine("Debug", this.GetPaletteDescription(palette));

            foreach (var command in ScopeCommandPaletteDefinitions.GetCommands(palette))
            {
                string commandText = ScopeCommandResolver.GetCommandText(selectedOscilloscope, command);

                if (string.IsNullOrWhiteSpace(commandText))
                {
                    throw new InvalidOperationException($"No SCPI command text is defined for {command}");
                }

                string effectiveCommandText = this.BuildEffectiveCommandText(command, commandText);

                if (string.IsNullOrWhiteSpace(effectiveCommandText))
                {
                    throw new InvalidOperationException($"No test value is available for {command}");
                }

                this.AppendOutputLine("Debug", $"SCPI >> {effectiveCommandText}");

                if (ScopeCommandResolver.ExpectsTextResponse(command))
                {
                    string response = await scopeClient.QueryLineAsync(effectiveCommandText, cancellationToken).ConfigureAwait(false);
                    string loggedResponse = command == ScopeCommand.Identify
                        ? this.MaskIdentifyResponseSerial(response)
                        : response;

                    this.AppendOutputLine("Debug", $"SCPI << {loggedResponse}");
                    this.ProcessTextResponse(command, response);
                }
                else
                {
                    await scopeClient.SendAsync(effectiveCommandText, cancellationToken).ConfigureAwait(false);
                }
            }

            this.AppendPaletteCompletionInfo(palette);
        }

        // ###########################################################################################
        // Executes the image dump workflow with the same preparation and cleanup pattern as shown.
        // ###########################################################################################
        private async Task ExecuteDumpImageWorkflowAsync(
            ScopeScpiClient scopeClient,
            OscilloscopeEntry selectedOscilloscope,
            CancellationToken cancellationToken)
        {
            await this.ExecutePaletteAsync(
                scopeClient,
                selectedOscilloscope,
                ScopeCommandPalette.ClearStatistics,
                cancellationToken).ConfigureAwait(false);

            this.AppendOutputLine("Info", "Statistics cleared");

            await this.ExecutePaletteAsync(
                scopeClient,
                selectedOscilloscope,
                ScopeCommandPalette.Stop,
                cancellationToken).ConfigureAwait(false);

            this.AppendOutputLine("Info", "Trigger set to STOP");

            this.AppendOutputLine("Debug", "---");
            this.AppendOutputLine("Debug", "Query dump image:");

            string dumpImageCommand = ScopeCommandResolver.GetCommandText(selectedOscilloscope, ScopeCommand.DumpImage);
            if (string.IsNullOrWhiteSpace(dumpImageCommand))
            {
                throw new InvalidOperationException("No SCPI command text is defined for DumpImage");
            }

            this.AppendOutputLine("Debug", $"SCPI >> {dumpImageCommand}");

            byte[] rawData = await scopeClient.QueryBinaryBlockAsync(dumpImageCommand, cancellationToken).ConfigureAwait(false);
            this.AppendOutputLine("Debug", $"SCPI << <{rawData.Length} bytes binary>");

            this.AppendHexDump("Dumping FIRST 64 bytes from raw data stream:", rawData, 64, fromStart: true);
            this.AppendHexDump("Dumping LAST 64 bytes from raw data stream:", rawData, 64, fromStart: false);

            if (this.TryExtractBinaryPayload(rawData, out byte[] payload) &&
                this.TryReadBmpMetadata(payload, out int width, out int height, out short bitsPerPixel))
            {
                this.AppendOutputLine(
                    "Info",
                    $"Dumped image ({width}x{height}px, {bitsPerPixel}bpp BMP, {payload.Length / 1024d:0.0} KB)");
            }
            else
            {
                this.AppendOutputLine(
                    "Info",
                    $"Dumped image payload ({rawData.Length / 1024d:0.0} KB)");
            }

            await this.ExecutePaletteAsync(
                scopeClient,
                selectedOscilloscope,
                ScopeCommandPalette.Run,
                cancellationToken).ConfigureAwait(false);

            this.AppendOutputLine("Info", "Trigger set to RUN");
        }

        // ###########################################################################################
        // Returns the currently selected oscilloscope definition from the loaded Excel data.
        // Marshals to the UI thread when called from background code.
        // ###########################################################################################
        private OscilloscopeEntry? GetSelectedOscilloscope()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                return Dispatcher.UIThread.InvokeAsync(
                    this.GetSelectedOscilloscope,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
            }

            var selectedVendor = this.VendorComboBox.SelectedItem as string;
            var selectedSeries = this.SeriesOrModelComboBox.SelectedItem as string;

            if (string.IsNullOrWhiteSpace(selectedVendor) || string.IsNullOrWhiteSpace(selectedSeries))
            {
                return null;
            }

            return DataManager.Oscilloscopes.FirstOrDefault(entry =>
                string.Equals(entry.Brand, selectedVendor, StringComparison.OrdinalIgnoreCase) &&
                string.Equals(entry.SeriesOrModel, selectedSeries, StringComparison.OrdinalIgnoreCase));
        }

        // ###########################################################################################
        // Builds the final SCPI command text, formatting value placeholders when required.
        // ###########################################################################################
        private string BuildEffectiveCommandText(ScopeCommand command, string baseCommandText)
        {
            return command switch
            {
                ScopeCommand.SetTriggerLevel => this.thisLastTriggerLevelVolts.HasValue
                    ? this.FormatParameterizedCommand(baseCommandText, this.thisLastTriggerLevelVolts.Value)
                    : string.Empty,

                ScopeCommand.SetTimeDiv => this.thisLastTimeDivSeconds.HasValue
                    ? this.FormatParameterizedCommand(baseCommandText, this.thisLastTimeDivSeconds.Value)
                    : string.Empty,

                ScopeCommand.SetVoltsDiv => this.thisLastVoltsDivVolts.HasValue
                    ? this.FormatParameterizedCommand(baseCommandText, this.thisLastVoltsDivVolts.Value)
                    : string.Empty,

                _ => baseCommandText
            };
        }

        // ###########################################################################################
        // Formats a SCPI command by replacing a "{0}" placeholder or appending the value.
        // ###########################################################################################
        private string FormatParameterizedCommand(string baseCommandText, double value)
        {
            string formattedValue = this.FormatScpiNumber(value);

            if (baseCommandText.Contains("{0}", StringComparison.Ordinal))
            {
                return baseCommandText.Replace("{0}", formattedValue, StringComparison.Ordinal);
            }

            return $"{baseCommandText} {formattedValue}";
        }

        // ###########################################################################################
        // Processes returned text from query commands and stores parsed numeric results when needed.
        // ###########################################################################################
        private void ProcessTextResponse(ScopeCommand command, string response)
        {
            switch (command)
            {
                case ScopeCommand.Identify:
                    this.AppendIdentifyResponse(response);
                    break;

                case ScopeCommand.DrainErrorQueue:
                    this.AppendOutputLine("Debug", $"System error: {response}");
                    break;

                case ScopeCommand.QueryTriggerLevel:
                    if (double.TryParse(response, NumberStyles.Float, CultureInfo.InvariantCulture, out double triggerLevel))
                    {
                        this.thisLastTriggerLevelVolts = triggerLevel;
                        this.AppendOutputLine("Info", $"Trigger level read as {this.FormatVoltage(triggerLevel)}");
                    }
                    break;

                case ScopeCommand.QueryTimeDiv:
                    if (double.TryParse(response, NumberStyles.Float, CultureInfo.InvariantCulture, out double timeDiv))
                    {
                        this.thisLastTimeDivSeconds = timeDiv;
                        this.AppendOutputLine("Info", $"TIME/DIV read as {this.FormatTime(timeDiv)}");
                    }
                    break;

                case ScopeCommand.QueryVoltsDiv:
                    if (double.TryParse(response, NumberStyles.Float, CultureInfo.InvariantCulture, out double voltsDiv))
                    {
                        this.thisLastVoltsDivVolts = voltsDiv;
                        this.AppendOutputLine("Info", $"VOLTS/DIV read as {this.FormatVoltage(voltsDiv)} per division");
                    }
                    break;
            }
        }

        // ###########################################################################################
        // Appends any additional completion info after a full command palette has finished.
        // ###########################################################################################
        private void AppendPaletteCompletionInfo(ScopeCommandPalette palette)
        {
            switch (palette)
            {
                case ScopeCommandPalette.SetTriggerLevel:
                    if (this.thisLastTriggerLevelVolts.HasValue)
                    {
                        this.AppendOutputLine("Info", $"Trigger level set to {this.FormatVoltage(this.thisLastTriggerLevelVolts.Value)}");
                    }
                    break;

                case ScopeCommandPalette.SetTimeDiv:
                    if (this.thisLastTimeDivSeconds.HasValue)
                    {
                        this.AppendOutputLine("Info", $"TIME/DIV set to {this.FormatTime(this.thisLastTimeDivSeconds.Value)}");
                    }
                    break;

                case ScopeCommandPalette.SetVoltsDiv:
                    if (this.thisLastVoltsDivVolts.HasValue)
                    {
                        this.AppendOutputLine("Info", $"VOLTS/DIV set to {this.FormatVoltage(this.thisLastVoltsDivVolts.Value)} per division");
                    }
                    break;
            }
        }

        // ###########################################################################################
        // Returns a readable debug description for the currently executed command palette.
        // ###########################################################################################
        private string GetPaletteDescription(ScopeCommandPalette palette)
        {
            return palette switch
            {
                ScopeCommandPalette.Identify => "Query identify instrument:",
                ScopeCommandPalette.DrainErrorQueue => "Query last system error:",
                ScopeCommandPalette.OperationComplete => "Query \"Operation Complete\":",
                ScopeCommandPalette.ClearStatistics => "Set \"Clear Statistics\":",
                ScopeCommandPalette.QueryActiveTrigger => "Query active trigger:",
                ScopeCommandPalette.Stop => "Set \"Stop\" mode:",
                ScopeCommandPalette.Single => "Set \"Single\" mode:",
                ScopeCommandPalette.Run => "Set \"Run\" mode:",
                ScopeCommandPalette.QueryTriggerMode => "Query trigger mode:",
                ScopeCommandPalette.QueryTriggerLevel => "Query trigger level:",
                ScopeCommandPalette.SetTriggerLevel => "Set trigger level:",
                ScopeCommandPalette.QueryTimeDiv => "Query TIME/DIV:",
                ScopeCommandPalette.SetTimeDiv => "Set TIME/DIV:",
                ScopeCommandPalette.QueryVoltsDiv => "Query VOLTS/DIV:",
                ScopeCommandPalette.SetVoltsDiv => "Set VOLTS/DIV:",
                ScopeCommandPalette.DumpImage => "Query dump image:",
                _ => palette.ToString()
            };
        }

        // ###########################################################################################
        // Appends a timestamped oscilloscope output line to a buffered UI queue and mirrors the same
        // message to the normal logfile immediately. The UI flush is batched to reduce churn.
        // ###########################################################################################
        private void AppendOutputLine(string level, string message)
        {
            string line = $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] [{level}] {message}";

            Logger.Debug("[Oscilloscope] " + message);

            bool shouldScheduleFlush = false;

            lock (this.thisPendingOutputLinesLock)
            {
                this.thisPendingOutputLines.Add(line);

                if (!this.thisOutputFlushScheduled)
                {
                    this.thisOutputFlushScheduled = true;
                    shouldScheduleFlush = true;
                }
            }

            if (shouldScheduleFlush)
            {
                _ = this.FlushPendingOutputLinesAsync();
            }
        }

        // ###########################################################################################
        // Flushes buffered oscilloscope output lines to the UI in a small batch so repeated SCPI
        // logging does not continuously rebuild the textbox contents on every single line.
        // ###########################################################################################
        private async Task FlushPendingOutputLinesAsync()
        {
            await Task.Delay(40).ConfigureAwait(false);

            await Dispatcher.UIThread.InvokeAsync(() =>
            {
                List<string> linesToAppend;

                lock (this.thisPendingOutputLinesLock)
                {
                    linesToAppend = this.thisPendingOutputLines.ToList();
                    this.thisPendingOutputLines.Clear();
                    this.thisOutputFlushScheduled = false;
                }

                if (linesToAppend.Count == 0)
                {
                    return;
                }

                string existingText = this.OutputTextBox.Text ?? string.Empty;
                string appendedText = string.Join(Environment.NewLine, linesToAppend) + Environment.NewLine;

                this.OutputTextBox.Text = existingText + appendedText;
                this.OutputTextBox.CaretIndex = this.OutputTextBox.Text.Length;

                bool shouldScheduleFlush = false;

                lock (this.thisPendingOutputLinesLock)
                {
                    if (this.thisPendingOutputLines.Count > 0 && !this.thisOutputFlushScheduled)
                    {
                        this.thisOutputFlushScheduled = true;
                        shouldScheduleFlush = true;
                    }
                }

                if (shouldScheduleFlush)
                {
                    _ = this.FlushPendingOutputLinesAsync();
                }
            },
            DispatcherPriority.Background);
        }

        // ###########################################################################################
        // Enables or disables the oscilloscope action buttons while a command is running.
        // ###########################################################################################
        private void SetOscilloscopeButtonsEnabled(bool isEnabled)
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    () => this.SetOscilloscopeButtonsEnabled(isEnabled),
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            this.ConnectToOscilloscopeButton.IsEnabled = isEnabled;
            this.RunFullTestSuiteButton.IsEnabled = isEnabled &&
                this.thisLastOscilloscopeConnectionState == true;
        }

        // ###########################################################################################
        // Parses and prints the standard *IDN? identify response in a readable form.
        // ###########################################################################################
        private void AppendIdentifyResponse(string response)
        {
            var parts = (response ?? string.Empty)
                .Split(',')
                .Select(part => part.Trim())
                .ToArray();

            string vendor = parts.Length > 0 ? parts[0] : string.Empty;
            string model = parts.Length > 1 ? parts[1] : string.Empty;
            string serial = parts.Length > 2 ? parts[2] : string.Empty;
            string firmware = parts.Length > 3 ? parts[3] : string.Empty;

            this.AppendOutputLine("Info", $"Vendor: {vendor}");
            this.AppendOutputLine("Info", $"Model: {model}");
            this.AppendOutputLine("Info", $"Serial: {this.MaskScopeSerial(serial)}");
            this.AppendOutputLine("Info", $"Firmware: {firmware}");
        }

        // ###########################################################################################
        // Formats a numeric SCPI value using invariant culture.
        // ###########################################################################################
        private string FormatScpiNumber(double value)
        {
            return value.ToString("G15", CultureInfo.InvariantCulture).Replace("E", "e", StringComparison.Ordinal);
        }

        // ###########################################################################################
        // Formats a voltage value into a compact engineering string.
        // ###########################################################################################
        private string FormatVoltage(double volts)
        {
            double absoluteValue = Math.Abs(volts);

            if (absoluteValue >= 1.0)
            {
                return $"{volts.ToString("0.###", CultureInfo.InvariantCulture)}V";
            }

            if (absoluteValue >= 0.001)
            {
                return $"{(volts * 1000.0).ToString("0.###", CultureInfo.InvariantCulture)}mV";
            }

            return $"{(volts * 1000000.0).ToString("0.###", CultureInfo.InvariantCulture)}uV";
        }

        // ###########################################################################################
        // Formats a time value into a compact engineering string.
        // ###########################################################################################
        private string FormatTime(double seconds)
        {
            double absoluteValue = Math.Abs(seconds);

            if (absoluteValue >= 1.0)
            {
                return $"{seconds.ToString("0.###", CultureInfo.InvariantCulture)}S";
            }

            if (absoluteValue >= 0.001)
            {
                return $"{(seconds * 1000.0).ToString("0.###", CultureInfo.InvariantCulture)}mS";
            }

            if (absoluteValue >= 0.000001)
            {
                return $"{(seconds * 1000000.0).ToString("0.###", CultureInfo.InvariantCulture)}uS";
            }

            return $"{(seconds * 1000000000.0).ToString("0.###", CultureInfo.InvariantCulture)}nS";
        }

        // ###########################################################################################
        // Extracts the payload bytes from a raw SCPI definite-length binary block response.
        // ###########################################################################################
        private bool TryExtractBinaryPayload(byte[] rawData, out byte[] payload)
        {
            payload = Array.Empty<byte>();

            if (rawData.Length < 3 || rawData[0] != (byte)'#')
            {
                return false;
            }

            int lengthDigits = rawData[1] - (byte)'0';
            if (lengthDigits < 1 || rawData.Length < 2 + lengthDigits)
            {
                return false;
            }

            string lengthText = System.Text.Encoding.ASCII.GetString(rawData, 2, lengthDigits);
            if (!int.TryParse(lengthText, out int payloadLength) ||
                payloadLength < 0 ||
                rawData.Length < 2 + lengthDigits + payloadLength)
            {
                return false;
            }

            payload = new byte[payloadLength];
            Buffer.BlockCopy(rawData, 2 + lengthDigits, payload, 0, payloadLength);
            return true;
        }

        // ###########################################################################################
        // Reads basic BMP metadata from the dumped image payload when the format is BMP.
        // ###########################################################################################
        private bool TryReadBmpMetadata(byte[] imageBytes, out int width, out int height, out short bitsPerPixel)
        {
            width = 0;
            height = 0;
            bitsPerPixel = 0;

            if (imageBytes.Length < 30)
            {
                return false;
            }

            if (imageBytes[0] != (byte)'B' || imageBytes[1] != (byte)'M')
            {
                return false;
            }

            width = BitConverter.ToInt32(imageBytes, 18);
            height = BitConverter.ToInt32(imageBytes, 22);
            bitsPerPixel = BitConverter.ToInt16(imageBytes, 28);
            return true;
        }

        // ###########################################################################################
        // Writes a short hex dump from either the start or end of a byte buffer to the output panel.
        // ###########################################################################################
        private void AppendHexDump(string title, byte[] rawData, int maxBytes, bool fromStart)
        {
            this.AppendOutputLine("Debug", title);

            if (rawData.Length == 0)
            {
                return;
            }

            int byteCount = Math.Min(maxBytes, rawData.Length);
            int startOffset = fromStart ? 0 : rawData.Length - byteCount;

            for (int offset = 0; offset < byteCount; offset += 16)
            {
                int lineCount = Math.Min(16, byteCount - offset);
                int absoluteOffset = startOffset + offset;

                var hexParts = new string[16];
                var asciiChars = new char[16];

                for (int i = 0; i < 16; i++)
                {
                    if (i < lineCount)
                    {
                        byte value = rawData[absoluteOffset + i];
                        hexParts[i] = value.ToString("X2", CultureInfo.InvariantCulture);
                        asciiChars[i] = value >= 32 && value <= 126 ? (char)value : '.';
                    }
                    else
                    {
                        hexParts[i] = "  ";
                        asciiChars[i] = ' ';
                    }
                }

                string hexText = string.Join(" ", hexParts);
                string asciiText = new string(asciiChars);

                this.AppendOutputLine("Debug", $"{absoluteOffset:X8}  {hexText}   {asciiText}");
            }
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
                this.thisMainWindowTitleBase = this.GetMainWindowTitleBase(window.Title ?? string.Empty);
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
        // Removes any existing oscilloscope connection suffix from the main window title.
        // ###########################################################################################
        private string GetMainWindowTitleBase(string windowTitle)
        {
            const string connectedSuffix = " (oscilloscope connected)";
            const string disconnectedSuffix = " (oscilloscope disconnected)";

            if (windowTitle.EndsWith(connectedSuffix, StringComparison.Ordinal))
            {
                return windowTitle[..^connectedSuffix.Length];
            }

            if (windowTitle.EndsWith(disconnectedSuffix, StringComparison.Ordinal))
            {
                return windowTitle[..^disconnectedSuffix.Length];
            }

            return windowTitle;
        }

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

                string signature = this.BuildComponentImageSyncSignature(
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
        private string BuildComponentImageSyncSignature(
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
        // Returns true only when the user has explicitly connected, the ping monitor reports the
        // scope as reachable, and a persistent SCPI client is currently stored for reuse.
        // ###########################################################################################
        private bool HasActiveEstablishedOscilloscopeSession()
        {
            return this.thisHasEstablishedOscilloscopeSession &&
                   this.thisLastOscilloscopeConnectionState == true &&
                   this.thisConnectedScopeClient != null;
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
            catch
            {
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
        // Container for one queued component-image sync request, pairing the image entry with the
        // oscilloscope UI snapshot that was current when the image was selected.
        // ###########################################################################################
        private sealed class PendingComponentImageSyncRequest
        {
            public ComponentImageEntry ComponentImageEntry { get; init; } = new();
            public OscilloscopeSelectionSnapshot SelectionSnapshot { get; init; } = new();
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

            bool shouldShowTitleSuffix =
                this.thisHasSeenEstablishedOscilloscopeSession ||
                UserSettings.OscilloscopeAutoConnect;

            this.RunFullTestSuiteButton.IsEnabled = hasEstablishedSession;

            Window? window = this.thisMainWindow ?? TopLevel.GetTopLevel(this) as Window;
            if (window != null)
            {
                if (string.IsNullOrWhiteSpace(this.thisMainWindowTitleBase))
                {
                    this.thisMainWindowTitleBase = this.GetMainWindowTitleBase(window.Title ?? string.Empty);
                }

                string newTitle = shouldShowTitleSuffix
                    ? this.thisMainWindowTitleBase + (hasEstablishedSession
                        ? " (oscilloscope connected)"
                        : " (oscilloscope disconnected)")
                    : this.thisMainWindowTitleBase;

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
        // Queries the current TIME/DIV, resolves the previous or next supported value from the main
        // oscilloscope Excel definition, and sends the SetTimeDiv palette over the active session.
        // ###########################################################################################
        public async Task StepTimeDivAsync(int offset, CancellationToken cancellationToken)
        {
            if (offset == 0)
            {
                return;
            }

            OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();

            await this.RunWithEstablishedOscilloscopeSessionAsync(
                selectionSnapshot,
                async (scopeClient, oscilloscopeEntry, token) =>
                {
                    this.thisLastTimeDivSeconds = null;

                    await this.ExecutePaletteAsync(
                        scopeClient,
                        oscilloscopeEntry,
                        ScopeCommandPalette.QueryTimeDiv,
                        token).ConfigureAwait(false);

                    if (!this.thisLastTimeDivSeconds.HasValue)
                    {
                        this.AppendOutputLine("Warning", "Could not read current TIME/DIV from oscilloscope");
                        return;
                    }

                    double currentTimeDivSeconds = this.thisLastTimeDivSeconds.Value;

                    if (!ScopeValueMapper.TryGetAdjacentTimeDivValue(
                        oscilloscopeEntry,
                        currentTimeDivSeconds,
                        offset,
                        out ScopeMappedValue mappedTimeDiv))
                    {
                        string directionLabel = offset < 0 ? "previous" : "next";
                        this.AppendOutputLine(
                            "Warning",
                            $"Could not resolve the {directionLabel} TIME/DIV value from the oscilloscope definition list");
                        return;
                    }

                    this.AppendOutputLine(
                        "Info",
                        $"Keyboard TIME/DIV step: {this.FormatTime(currentTimeDivSeconds)} -> {mappedTimeDiv.MatchedDisplayValue}");

                    this.thisLastTimeDivSeconds = mappedTimeDiv.NumericValue;
                    this.thisLastOscilloscopeImageSyncSignature = string.Empty;

                    await this.ExecutePaletteAsync(
                        scopeClient,
                        oscilloscopeEntry,
                        ScopeCommandPalette.SetTimeDiv,
                        token).ConfigureAwait(false);
                },
                cancellationToken,
                writeWarnings: true).ConfigureAwait(false);
        }

        // ###########################################################################################
        // Uses a cached trigger level for keyboard stepping whenever possible so repeated Up/Down
        // keypresses only send SetTriggerLevel instead of querying the oscilloscope each time.
        // ###########################################################################################
        public async Task StepTriggerLevelAsync(int direction, CancellationToken cancellationToken)
        {
            if (direction == 0)
            {
                return;
            }

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

                    double currentTriggerLevelVolts = this.thisLastTriggerLevelVolts!.Value;
                    double targetTriggerLevelVolts = this.GetNextSnappedTriggerLevelVolts(
                        currentTriggerLevelVolts,
                        direction);

                    await this.SendTriggerLevelFastAsync(
                        scopeClient,
                        oscilloscopeEntry,
                        targetTriggerLevelVolts,
                        token).ConfigureAwait(false);

                    this.thisLastOscilloscopeImageSyncSignature = string.Empty;

                    this.AppendOutputLine(
                        "Info",
                        $"Keyboard trigger level step: {this.FormatVoltage(currentTriggerLevelVolts)} -> {this.FormatVoltage(targetTriggerLevelVolts)}");
                },
                cancellationToken,
                writeWarnings: true).ConfigureAwait(false);
        }

        // ###########################################################################################
        // Snaps a trigger level to the next 0.25V boundary in the requested direction so keyboard
        // stepping always moves predictably even when the current level is slightly off-grid.
        // ###########################################################################################
        private double GetNextSnappedTriggerLevelVolts(double currentTriggerLevelVolts, int direction)
        {
            const double triggerLevelStepVolts = 0.25;
            const double stepTolerance = 1e-6;

            double scaledValue = currentTriggerLevelVolts / triggerLevelStepVolts;
            double nearestWholeStep = Math.Round(scaledValue, MidpointRounding.AwayFromZero);
            bool isNearWholeStep = Math.Abs(scaledValue - nearestWholeStep) <= stepTolerance;

            if (direction > 0)
            {
                double targetStep = isNearWholeStep
                    ? nearestWholeStep + 1.0
                    : Math.Ceiling(scaledValue);

                return targetStep * triggerLevelStepVolts;
            }

            double downTargetStep = isNearWholeStep
                ? nearestWholeStep - 1.0
                : Math.Floor(scaledValue);

            return downTargetStep * triggerLevelStepVolts;
        }

        // ###########################################################################################
        // Resolves a fixed VOLTS/DIV value from the oscilloscope definition list and sends the
        // SetVoltsDiv palette over the active session.
        // ###########################################################################################
        public async Task SetVoltsDivAsync(double targetVoltsDivVolts, CancellationToken cancellationToken)
        {
            if (targetVoltsDivVolts <= 0)
            {
                return;
            }

            OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();

            await this.RunWithEstablishedOscilloscopeSessionAsync(
                selectionSnapshot,
                async (scopeClient, oscilloscopeEntry, token) =>
                {
                    if (!ScopeValueMapper.TryGetSupportedVoltsDivValue(
                        oscilloscopeEntry,
                        targetVoltsDivVolts,
                        out ScopeMappedValue mappedVoltsDiv))
                    {
                        this.AppendOutputLine(
                            "Warning",
                            $"Could not resolve keyboard VOLTS/DIV value [{this.FormatVoltage(targetVoltsDivVolts)}] from the oscilloscope definition list");
                        return;
                    }

                    this.AppendOutputLine(
                        "Info",
                        $"Keyboard VOLTS/DIV set: {mappedVoltsDiv.MatchedDisplayValue} per division");

                    this.thisLastVoltsDivVolts = mappedVoltsDiv.NumericValue;
                    this.thisLastOscilloscopeImageSyncSignature = string.Empty;

                    await this.ExecutePaletteAsync(
                        scopeClient,
                        oscilloscopeEntry,
                        ScopeCommandPalette.SetVoltsDiv,
                        token).ConfigureAwait(false);
                },
                cancellationToken,
                writeWarnings: true).ConfigureAwait(false);
        }

        // ###########################################################################################
        // Executes one named oscilloscope command palette on the already established session so
        // keyboard shortcuts can reuse the normal SCPI logging and palette behavior.
        // ###########################################################################################
        public async Task RunPaletteAsync(ScopeCommandPalette palette, CancellationToken cancellationToken)
        {
            OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();

            await this.RunWithEstablishedOscilloscopeSessionAsync(
                selectionSnapshot,
                async (scopeClient, oscilloscopeEntry, token) =>
                {
                    this.thisLastOscilloscopeImageSyncSignature = string.Empty;

                    await this.ExecutePaletteAsync(
                        scopeClient,
                        oscilloscopeEntry,
                        palette,
                        token).ConfigureAwait(false);
                },
                cancellationToken,
                writeWarnings: true).ConfigureAwait(false);
        }

        // ###########################################################################################
        // Persists the auto-connect checkbox state and starts or stops the background auto-connect loop.
        // ###########################################################################################
        private void OnAutoConnectOscilloscopeCheckBoxChanged()
        {
            bool isEnabled = this.AutoConnectOscilloscopeCheckBox.IsChecked == true;

            UserSettings.OscilloscopeAutoConnect = isEnabled;
            this.thisShouldAutoReconnectEstablishedOscilloscopeSession = isEnabled;

            if (isEnabled)
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

                    if (!UserSettings.OscilloscopeAutoConnect)
                    {
                        await Task.Delay(
                            this.GetOscilloscopeAutoConnectRetryDelay(),
                            cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    if (selectionSnapshot.HasActiveEstablishedSession ||
                        this.thisHasEstablishedOscilloscopeSession ||
                        this.thisConnectedScopeClient != null)
                    {
                        await Task.Delay(
                            this.GetOscilloscopeAutoConnectRetryDelay(),
                            cancellationToken).ConfigureAwait(false);
                        continue;
                    }

                    if (!this.TryValidateOscilloscopeSelectionSnapshot(selectionSnapshot, writeWarnings: false))
                    {
                        await Task.Delay(
                            this.GetOscilloscopeAutoConnectRetryDelay(),
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
                        this.GetOscilloscopeAutoConnectRetryDelay(),
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

            if (UserSettings.OscilloscopeAutoConnect)
            {
                this.thisShouldAutoReconnectEstablishedOscilloscopeSession = true;
                this.StartOscilloscopeAutoConnectLoop();
            }
        }

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
            ScopeScpiClient scopeClient,
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
                                    double targetTriggerLevelVolts = this.GetNextSnappedTriggerLevelVolts(
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
                                        ? $"Keyboard trigger level step: {this.FormatVoltage(startingTriggerLevelVolts)} -> {this.FormatVoltage(currentTriggerLevelVolts)}"
                                        : $"Keyboard trigger level step: {this.FormatVoltage(startingTriggerLevelVolts)} -> {this.FormatVoltage(currentTriggerLevelVolts)} ({stepCount} steps)");
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
            ScopeScpiClient scopeClient,
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
            ScopeScpiClient scopeClient,
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
                                        ? $"Keyboard TIME/DIV step: {this.FormatTime(startingTimeDivSeconds)} -> {this.FormatTime(currentTimeDivSeconds)}"
                                        : $"Keyboard TIME/DIV step: {this.FormatTime(startingTimeDivSeconds)} -> {this.FormatTime(currentTimeDivSeconds)} ({appliedStepCount} steps)");
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
            ScopeScpiClient scopeClient,
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
                                    $"Could not resolve keyboard VOLTS/DIV value [{this.FormatVoltage(pendingTargetVoltsDivVolts.Value)}] from the oscilloscope definition list");
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
            ScopeScpiClient scopeClient,
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

        // ###########################################################################################
        // Opens the Avalonia folder picker and stores the selected oscilloscope image folder.
        // ###########################################################################################
        private async void OnSelectOscilloscopeImageFolderClick(object? sender, RoutedEventArgs e)
        {
            try
            {
                TopLevel? topLevel = TopLevel.GetTopLevel(this);
                if (topLevel?.StorageProvider == null)
                {
                    this.AppendOutputLine("Warning", "Folder picker is not available");
                    return;
                }

                if (!topLevel.StorageProvider.CanPickFolder)
                {
                    this.AppendOutputLine("Warning", "This platform does not support folder picking");
                    return;
                }

                IReadOnlyList<IStorageFolder> selectedFolders =
                    await topLevel.StorageProvider.OpenFolderPickerAsync(
                        new FolderPickerOpenOptions
                        {
                            Title = "Select oscilloscope image folder",
                            AllowMultiple = false
                        });

                IStorageFolder? selectedFolder = selectedFolders.FirstOrDefault();
                string? selectedPath = selectedFolder?.TryGetLocalPath();

                if (string.IsNullOrWhiteSpace(selectedPath))
                {
                    return;
                }

                string normalizedPath = Path.GetFullPath(selectedPath);

                UserSettings.OscilloscopeImageFolder = normalizedPath;
                this.UpdateOscilloscopeImageFolderUi();

                this.AppendOutputLine("Info", $"Oscilloscope image folder set to [{normalizedPath}]");
            }
            catch (Exception ex)
            {
                this.AppendOutputLine("Warning", $"Failed to select oscilloscope image folder: {ex.Message}");
            }
        }

        // ###########################################################################################
        // Opens the configured oscilloscope image folder in the operating system file manager.
        // ###########################################################################################
        private void OnOpenOscilloscopeImageFolderClick(object? sender, RoutedEventArgs e)
        {
            string directoryPath = UserSettings.OscilloscopeImageFolder?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(directoryPath))
            {
                this.AppendOutputLine("Warning", "Select an oscilloscope image folder first");
                this.UpdateOscilloscopeImageFolderUi();
                return;
            }

            try
            {
                this.OpenDirectoryInFileExplorer(directoryPath);
            }
            catch (Exception ex)
            {
                this.AppendOutputLine("Warning", $"Failed to open oscilloscope image folder: {ex.Message}");
            }
        }

        // ###########################################################################################
        // Refreshes the folder path textbox and the enabled state of the open-folder button.
        // ###########################################################################################
        private void UpdateOscilloscopeImageFolderUi()
        {
            if (!Dispatcher.UIThread.CheckAccess())
            {
                Dispatcher.UIThread.InvokeAsync(
                    this.UpdateOscilloscopeImageFolderUi,
                    DispatcherPriority.Background).GetAwaiter().GetResult();
                return;
            }

            string directoryPath = UserSettings.OscilloscopeImageFolder?.Trim() ?? string.Empty;

            this.OscilloscopeImageFolderTextBox.Text = directoryPath;
            this.OpenOscilloscopeImageFolderButton.IsEnabled = !string.IsNullOrWhiteSpace(directoryPath);
        }

        // ###########################################################################################
        // Creates the target directory if needed and opens it in the native file explorer.
        // ###########################################################################################
        private void OpenDirectoryInFileExplorer(string directoryPath)
        {
            string normalizedPath = Path.GetFullPath(directoryPath);
            Directory.CreateDirectory(normalizedPath);

            if (OperatingSystem.IsWindows())
            {
                Process.Start(new ProcessStartInfo("explorer.exe", $"\"{normalizedPath}\"")
                {
                    UseShellExecute = true
                });
            }
            else if (OperatingSystem.IsMacOS())
            {
                Process.Start("open", normalizedPath);
            }
            else
            {
                Process.Start("xdg-open", normalizedPath);
            }
        }

        // ###########################################################################################
        // Captures the current oscilloscope screenshot over the active SCPI session, saves it as a PNG
        // in the configured image folder, and returns the saved file path.
        // ###########################################################################################
        public async Task<string?> CaptureAndSaveOscilloscopeImageAsync(
            ComponentImageEntry componentImageEntry,
            string displayedRegion,
            CancellationToken cancellationToken)
        {
            if (componentImageEntry == null)
            {
                this.AppendOutputLine("Warning", "No active component image is selected");
                return null;
            }

            if (string.IsNullOrWhiteSpace(componentImageEntry.BoardLabel))
            {
                this.AppendOutputLine("Warning", "The selected component image has no board label");
                return null;
            }

            if (string.IsNullOrWhiteSpace(componentImageEntry.Pin) &&
                string.IsNullOrWhiteSpace(componentImageEntry.Name))
            {
                this.AppendOutputLine("Warning", "The selected component image has neither a pin value nor a name");
                return null;
            }

            string outputDirectory = UserSettings.OscilloscopeImageFolder?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(outputDirectory))
            {
                this.AppendOutputLine("Warning", "Select an oscilloscope image folder first");
                this.UpdateOscilloscopeImageFolderUi();
                return null;
            }

            OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();
            if (!this.TryValidateOscilloscopeSelectionSnapshot(selectionSnapshot, writeWarnings: true))
            {
                return null;
            }

            string outputFilePath = this.BuildCapturedOscilloscopeImageFilePath(
                componentImageEntry,
                displayedRegion,
                outputDirectory);

            bool enteredSemaphore = false;

            try
            {
                await this.thisOscilloscopeSessionSemaphore.WaitAsync(cancellationToken).ConfigureAwait(false);
                enteredSemaphore = true;

                if (!selectionSnapshot.HasActiveEstablishedSession || this.thisConnectedScopeClient == null)
                {
                    this.AppendOutputLine("Warning", "Connect to oscilloscope first");
                    return null;
                }

                await Dispatcher.UIThread.InvokeAsync(
                    () => this.SetOscilloscopeButtonsEnabled(false),
                    DispatcherPriority.Background);

                using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
                    timeoutCts.Token,
                    cancellationToken);

                byte[] rawImageData = await this.QueryDumpImagePaletteAsync(
                    this.thisConnectedScopeClient,
                    selectionSnapshot.SelectedOscilloscope!,
                    linkedCts.Token).ConfigureAwait(false);

                if (!this.TryCreateBitmapFromScopeRawImageData(rawImageData, out Bitmap? capturedBitmap))
                {
                    this.AppendOutputLine("Warning", "Could not decode the oscilloscope image");
                    return null;
                }

                using (capturedBitmap)
                {
                    Directory.CreateDirectory(outputDirectory);
                    capturedBitmap.Save(outputFilePath);
                }

                this.AppendOutputLine("Info", $"Saved oscilloscope image to [{outputFilePath}]");
                return outputFilePath;
            }
            catch (OperationCanceledException) when (cancellationToken.IsCancellationRequested)
            {
                return null;
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

            return null;
        }

        // ###########################################################################################
        // Executes only the DumpImage palette command, logs the SCPI traffic, and returns the raw
        // binary block received from the oscilloscope.
        // ###########################################################################################
        private async Task<byte[]> QueryDumpImagePaletteAsync(
            ScopeScpiClient scopeClient,
            OscilloscopeEntry selectedOscilloscope,
            CancellationToken cancellationToken)
        {
            this.AppendOutputLine("Debug", "---");
            this.AppendOutputLine("Debug", this.GetPaletteDescription(ScopeCommandPalette.DumpImage));

            string dumpImageCommand = ScopeCommandResolver.GetCommandText(selectedOscilloscope, ScopeCommand.DumpImage);
            if (string.IsNullOrWhiteSpace(dumpImageCommand))
            {
                throw new InvalidOperationException("No SCPI command text is defined for DumpImage");
            }

            this.AppendOutputLine("Debug", $"SCPI >> {dumpImageCommand}");

            byte[] rawData = await scopeClient.QueryBinaryBlockAsync(dumpImageCommand, cancellationToken).ConfigureAwait(false);

            this.AppendOutputLine("Debug", $"SCPI << <{rawData.Length} bytes binary>");
            return rawData;
        }

        // ###########################################################################################
        // Converts the raw oscilloscope DumpImage response into an Avalonia bitmap. The method first
        // strips a SCPI definite-length header when present, then falls back to the raw buffer.
        // ###########################################################################################
        private bool TryCreateBitmapFromScopeRawImageData(byte[] rawImageData, [NotNullWhen(true)] out Bitmap? bitmap)
        {
            bitmap = null;

            try
            {
                if (this.TryExtractBinaryPayload(rawImageData, out byte[] payload) && payload.Length > 0)
                {
                    using var payloadStream = new MemoryStream(payload, writable: false);
                    bitmap = new Bitmap(payloadStream);
                    return true;
                }

                using var rawStream = new MemoryStream(rawImageData, writable: false);
                bitmap = new Bitmap(rawStream);
                return true;
            }
            catch
            {
                bitmap = null;
                return false;
            }
        }

        // ###########################################################################################
        // Builds the PNG file path for one captured oscilloscope image using the selected component
        // image metadata and the popup's currently displayed region.
        // ###########################################################################################
        private string BuildCapturedOscilloscopeImageFilePath(
            ComponentImageEntry componentImageEntry,
            string displayedRegion,
            string outputDirectory)
        {
            string safeBoardLabel = this.SanitizeCapturedOscilloscopeImageFileNamePart(componentImageEntry.BoardLabel);

            string identityPart = !string.IsNullOrWhiteSpace(componentImageEntry.Pin)
                ? this.SanitizeCapturedOscilloscopeImageFileNamePart(componentImageEntry.Pin)
                : this.SanitizeCapturedOscilloscopeImageFileNamePart(componentImageEntry.Name);

            string safeRegion = string.IsNullOrWhiteSpace(displayedRegion)
                ? string.Empty
                : this.SanitizeCapturedOscilloscopeImageFileNamePart(displayedRegion);

            string fileName = string.IsNullOrWhiteSpace(safeRegion)
                ? $"{safeBoardLabel}_{identityPart}.png"
                : $"{safeBoardLabel}_{identityPart}_{safeRegion}.png";

            return Path.Combine(outputDirectory, fileName);
        }

        // ###########################################################################################
        // Replaces characters that are invalid in file names so oscilloscope captures can be written
        // safely on all supported desktop platforms.
        // ###########################################################################################
        private string SanitizeCapturedOscilloscopeImageFileNamePart(string value)
        {
            string sanitized = string.IsNullOrWhiteSpace(value) ? "Unknown" : value.Trim();

            foreach (char invalidChar in Path.GetInvalidFileNameChars())
            {
                sanitized = sanitized.Replace(invalidChar, '_');
            }

            return sanitized;
        }

        // ###########################################################################################
        // Masks a scope serial number so logs and the output panel do not expose the real value.
        // ###########################################################################################
        private string MaskScopeSerial(string serial)
        {
            return string.IsNullOrEmpty(serial)
                ? string.Empty
                : new string('*', serial.Length);
        }

        // ###########################################################################################
        // Masks only the serial field inside a standard *IDN? response while leaving the other parts
        // unchanged for debugging and display purposes.
        // ###########################################################################################
        private string MaskIdentifyResponseSerial(string response)
        {
            var parts = (response ?? string.Empty).Split(',');

            if (parts.Length > 2)
            {
                string trimmedSerial = parts[2].Trim();
                parts[2] = this.MaskScopeSerial(trimmedSerial);
            }

            return string.Join(",", parts);
        }


    }
}