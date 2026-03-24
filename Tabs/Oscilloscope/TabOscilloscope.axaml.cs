using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;
using System.Net.NetworkInformation;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    public partial class TabOscilloscope : UserControl
    {
        private double? thisLastTriggerLevelVolts;
        private double? thisLastTimeDivSeconds;
        private double? thisLastVoltsDivVolts;

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


        public TabOscilloscope()
        {
            this.InitializeComponent();

            this.RunFullTestSuiteButton.IsEnabled = false;

            this.HostTextBox.Text = UserSettings.OscilloscopeHost;
            this.HostTextBox.TextChanged += this.OnHostTextChanged;

            this.PortNumericUpDown.Value = UserSettings.OscilloscopePort;
            this.PortNumericUpDown.ValueChanged += this.OnPortValueChanged;

            this.VendorComboBox.SelectionChanged += this.OnVendorSelectionChanged;
            this.SeriesOrModelComboBox.SelectionChanged += this.OnSeriesOrModelSelectionChanged;

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
                this.PortNumericUpDown.Value = port;
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
        // Persists port numerical up/down field whenever it's updated.
        // ###########################################################################################
        private void OnPortValueChanged(object? sender, NumericUpDownValueChangedEventArgs e)
        {
            if (e.NewValue.HasValue)
            {
                UserSettings.OscilloscopePort = (int)e.NewValue.Value;
            }

            this.InvalidateEstablishedOscilloscopeSession();
        }

        // ###########################################################################################
        // Rejects non-digit text input in the TCP port field.
        // ###########################################################################################
        private void OnPortNumericUpDownTextInput(object? sender, TextInputEventArgs e)
        {
            if (string.IsNullOrEmpty(e.Text))
            {
                return;
            }

            if (e.Text.Any(ch => !char.IsDigit(ch)))
            {
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Restores a valid TCP port value when the field loses focus.
        // ###########################################################################################
        private void OnPortNumericUpDownLostFocus(object? sender, RoutedEventArgs e)
        {
            if (!this.PortNumericUpDown.Value.HasValue)
            {
                this.PortNumericUpDown.Value = UserSettings.OscilloscopePort is >= 1 and <= 65535
                    ? UserSettings.OscilloscopePort
                    : 5025;
                return;
            }

            if (this.PortNumericUpDown.Value.Value < 1)
            {
                this.PortNumericUpDown.Value = 1;
                return;
            }

            if (this.PortNumericUpDown.Value.Value > 65535)
            {
                this.PortNumericUpDown.Value = 65535;
            }
        }

        // ###########################################################################################
        // Captures the current oscilloscope UI state on the UI thread so background workers can use
        // the values without reading UI controls directly.
        // ###########################################################################################
        private OscilloscopeSelectionSnapshot CreateOscilloscopeSelectionSnapshot()
        {
            const int defaultDelayMilliseconds = 250;

            var selectedOscilloscope = this.GetSelectedOscilloscope();
            string host = this.HostTextBox.Text?.Trim() ?? string.Empty;
            int port = this.PortNumericUpDown.Value.HasValue ? (int)this.PortNumericUpDown.Value.Value : 0;

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
                    this.AppendOutputLine("Warning", "Select a vendor and series or model first.");
                }

                return false;
            }

            if (string.IsNullOrWhiteSpace(selectionSnapshot.Host))
            {
                if (writeWarnings)
                {
                    this.AppendOutputLine("Warning", "Enter an IP address or FQDN first.");
                }

                return false;
            }

            if (selectionSnapshot.Port < 1 || selectionSnapshot.Port > 65535)
            {
                if (writeWarnings)
                {
                    this.AppendOutputLine("Warning", "TCP port must be within 1-65535.");
                }

                return false;
            }

            return true;
        }

        // ###########################################################################################
        // Creates and stores one persistent SCPI client session after an explicit user connect.
        // The connection stays alive and is reused by later popup image auto-sync operations.
        // ###########################################################################################
        private async Task<bool> ConnectSelectedOscilloscopeAsync(CancellationToken externalCancellationToken)
        {
            OscilloscopeSelectionSnapshot selectionSnapshot = this.CreateOscilloscopeSelectionSnapshot();
            if (!this.TryValidateOscilloscopeSelectionSnapshot(selectionSnapshot, writeWarnings: true))
            {
                return false;
            }

            bool enteredSemaphore = false;
            ScopeScpiClient? newScopeClient = null;
            OscilloscopeEntry selectedOscilloscope = selectionSnapshot.SelectedOscilloscope!;

            this.SetOscilloscopeButtonsEnabled(false);
            this.AppendOutputLine("Debug", "---");
            this.AppendOutputLine("Info", $"Connecting to {selectedOscilloscope.Brand} {selectedOscilloscope.SeriesOrModel} at {selectionSnapshot.Host}:{selectionSnapshot.Port}.");

            try
            {
                await this.thisOscilloscopeSessionSemaphore.WaitAsync(externalCancellationToken);
                enteredSemaphore = true;

                using var timeoutCts = new CancellationTokenSource(TimeSpan.FromSeconds(30));
                using var linkedCts = CancellationTokenSource.CreateLinkedTokenSource(
                    timeoutCts.Token,
                    externalCancellationToken);

                await this.DisposeConnectedScopeClientCoreAsync();

                newScopeClient = new ScopeScpiClient();
                await newScopeClient.ConnectAsync(selectionSnapshot.Host, selectionSnapshot.Port, linkedCts.Token).ConfigureAwait(false);

                this.AppendOutputLine("Info", "Network session established.");

                await this.ExecutePaletteAsync(
                    newScopeClient,
                    selectedOscilloscope,
                    ScopeCommandPalette.Identify,
                    linkedCts.Token).ConfigureAwait(false);

                this.thisConnectedScopeClient = newScopeClient;
                newScopeClient = null;

                this.thisHasEstablishedOscilloscopeSession = true;
                this.thisLastOscilloscopeImageSyncSignature = string.Empty;

                await Dispatcher.UIThread.InvokeAsync(() =>
                {
                    this.StartOscilloscopeConnectionMonitoring();
                    this.UpdateMainWindowOscilloscopeConnectionState(true);
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
                this.AppendOutputLine("Warning", "Connection or SCPI communication timed out.");
            }
            catch (Exception ex)
            {
                this.AppendOutputLine("Critical", $"Oscilloscope communication failed: {ex.Message}");
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
                        this.AppendOutputLine("Warning", "Connect to oscilloscope first.");
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
                await this.HandleOscilloscopeSessionFailureCoreAsync("Connection or SCPI communication timed out.");
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
            await this.ConnectSelectedOscilloscopeAsync(CancellationToken.None);
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
                    throw new InvalidOperationException($"No SCPI command text is defined for {command}.");
                }

                string effectiveCommandText = this.BuildEffectiveCommandText(command, commandText);

                if (string.IsNullOrWhiteSpace(effectiveCommandText))
                {
                    throw new InvalidOperationException($"No test value is available for {command}.");
                }

                this.AppendOutputLine("Debug", $"SCPI >> {effectiveCommandText}");

                if (ScopeCommandResolver.ExpectsTextResponse(command))
                {
                    string response = await scopeClient.QueryLineAsync(effectiveCommandText, cancellationToken).ConfigureAwait(false);
                    this.AppendOutputLine("Debug", $"SCPI << {response}");
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
                throw new InvalidOperationException("No SCPI command text is defined for DumpImage.");
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
        // ###########################################################################################
        private OscilloscopeEntry? GetSelectedOscilloscope()
        {
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
            this.AppendOutputLine("Info", $"Serial: {serial}");
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
            string host = this.HostTextBox.Text?.Trim() ?? string.Empty;
            if (string.IsNullOrWhiteSpace(host))
            {
                return;
            }

            this.StopOscilloscopeConnectionMonitoring();

            if (TopLevel.GetTopLevel(this) is Window window)
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

            if (TopLevel.GetTopLevel(this) is Window window)
            {
                string baseTitle = this.GetMainWindowTitleBase(window.Title ?? string.Empty);
                if (!string.Equals(window.Title, baseTitle, StringComparison.Ordinal))
                {
                    window.Title = baseTitle;
                }
            }

            this.RunFullTestSuiteButton.IsEnabled = false;
            this.thisLastOscilloscopeConnectionState = null;
            this.thisLastOscilloscopeImageSyncSignature = string.Empty;
        }

        // ###########################################################################################
        // Repeatedly pings the oscilloscope host and updates the main window title on state changes.
        // ###########################################################################################
        private async Task RunOscilloscopeConnectionMonitoringAsync(string host, CancellationToken cancellationToken)
        {
            while (!cancellationToken.IsCancellationRequested)
            {
                bool isConnected = await this.PingOscilloscopeAsync(host);

                await Dispatcher.UIThread.InvokeAsync(
                    () => this.UpdateMainWindowOscilloscopeConnectionState(isConnected),
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
        // Updates the main window title with the current oscilloscope connection state.
        // Only applies changes when the state text actually changed.
        // ###########################################################################################
        private void UpdateMainWindowOscilloscopeConnectionState(bool isConnected)
        {
            if (this.thisLastOscilloscopeConnectionState == isConnected)
            {
                return;
            }

            if (TopLevel.GetTopLevel(this) is not Window window)
            {
                return;
            }

            if (string.IsNullOrWhiteSpace(this.thisMainWindowTitleBase))
            {
                this.thisMainWindowTitleBase = this.GetMainWindowTitleBase(window.Title ?? string.Empty);
            }

            string suffix = isConnected
                ? " (oscilloscope connected)"
                : " (oscilloscope disconnected)";

            string newTitle = this.thisMainWindowTitleBase + suffix;

            if (!string.Equals(window.Title, newTitle, StringComparison.Ordinal))
            {
                window.Title = newTitle;
            }

            this.RunFullTestSuiteButton.IsEnabled = isConnected;

            if (!isConnected)
            {
                this.thisHasEstablishedOscilloscopeSession = false;
                this.thisLastOscilloscopeImageSyncSignature = string.Empty;
                _ = this.DisposeConnectedScopeClientAsync();
            }

            this.thisLastOscilloscopeConnectionState = isConnected;
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
                        this.AppendOutputLine("Warning", $"Could not map image T/DIV value [{componentImageEntry.TimeDiv}] to a supported oscilloscope value.");
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
                        this.AppendOutputLine("Warning", $"Could not map image VOLTS/DIV value [{componentImageEntry.VoltsDiv}] to a supported oscilloscope value.");
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
                        this.AppendOutputLine("Warning", $"Could not parse image trigger level value [{componentImageEntry.TriggerLevelVolts}].");
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
                this.AppendOutputLine("Info", $"Auto-syncing oscilloscope settings from image pin [{componentImageEntry.Pin.Trim()}].");

                if (mappedTimeDiv != null)
                {
                    this.AppendOutputLine("Info", $"Image TIME/DIV [{mappedTimeDiv.RawValue}] mapped to scope value [{mappedTimeDiv.MatchedDisplayValue}].");
                }

                if (mappedVoltsDiv != null)
                {
                    this.AppendOutputLine("Info", $"Image VOLTS/DIV [{mappedVoltsDiv.RawValue}] mapped to scope value [{mappedVoltsDiv.MatchedDisplayValue}].");
                }

                if (mappedTriggerLevel != null)
                {
                    this.AppendOutputLine("Info", $"Image trigger level [{mappedTriggerLevel.RawValue}] mapped to SCPI value [{mappedTriggerLevel.ScpiValue}].");
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
            this.thisLastOscilloscopeImageSyncSignature = string.Empty;
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
        // the output pane so the next operation requires an explicit reconnect.
        // ###########################################################################################
        private async Task HandleOscilloscopeSessionFailureCoreAsync(string message)
        {
            this.AppendOutputLine("Critical", $"Oscilloscope communication failed: {message}");
            this.thisHasEstablishedOscilloscopeSession = false;
            this.thisLastOscilloscopeImageSyncSignature = string.Empty;
            this.StopOscilloscopeConnectionMonitoring();
            await this.DisposeConnectedScopeClientCoreAsync();
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

    }
}