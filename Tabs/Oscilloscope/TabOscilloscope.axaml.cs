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
        // ###########################################################################################
        // TabOscilloscope is split across partial-class files by area; this file owns construction,
        // vendor/series/host/port UI wiring, palette execution/response handling, and shared output
        // logging/button-state helpers. The others:
        //
        //   TabOscilloscope.Connection.cs - connect/disconnect flow, session lifecycle, connectivity
        //                                   monitoring, ping and auto-reconnect
        //   TabOscilloscope.ImageSync.cs  - the background worker that syncs a component image's
        //                                   oscilloscope baseline settings
        //   TabOscilloscope.Keyboard.cs   - the trigger-level / time-div / volts-div keyboard-step
        //                                   workers (three near-identical clusters)
        //   TabOscilloscope.Capture.cs    - image folder selection, dump-image workflow, saving a
        //                                   captured oscilloscope image
        //   TabOscilloscope.TestSeams.cs  - pre-existing test seam (unchanged by this split)
        // ###########################################################################################

        private double? thisLastTriggerLevelVolts;
        private double? thisLastTimeDivSeconds;
        private double? thisLastVoltsDivVolts;

        private bool thisIsNormalizingPortText;
        private string thisLastValidPortText = "5025";
        private readonly object thisPendingOutputLinesLock = new();
        private readonly List<string> thisPendingOutputLines = new();
        private bool thisOutputFlushScheduled;

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
        // Executes a named scope command palette and logs every sent and received SCPI command.
        // ###########################################################################################
        private async Task ExecutePaletteAsync(
            IScopeClient scopeClient,
            OscilloscopeEntry selectedOscilloscope,
            ScopeCommandPalette palette,
            CancellationToken cancellationToken)
        {
            this.AppendOutputLine("Debug", "---");
            this.AppendOutputLine("Debug", GetPaletteDescription(palette));

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
                        ? ScopeFormatting.MaskIdentifyResponseSerial(response)
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
                    ? FormatParameterizedCommand(baseCommandText, this.thisLastTriggerLevelVolts.Value)
                    : string.Empty,

                ScopeCommand.SetTimeDiv => this.thisLastTimeDivSeconds.HasValue
                    ? FormatParameterizedCommand(baseCommandText, this.thisLastTimeDivSeconds.Value)
                    : string.Empty,

                ScopeCommand.SetVoltsDiv => this.thisLastVoltsDivVolts.HasValue
                    ? FormatParameterizedCommand(baseCommandText, this.thisLastVoltsDivVolts.Value)
                    : string.Empty,

                _ => baseCommandText
            };
        }

        // ###########################################################################################
        // Formats a SCPI command by replacing a "{0}" placeholder or appending the value.
        // ###########################################################################################
        private static string FormatParameterizedCommand(string baseCommandText, double value)
        {
            string formattedValue = ScopeFormatting.FormatScpiNumber(value);

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
                        this.AppendOutputLine("Info", $"Trigger level read as {ScopeFormatting.FormatVoltage(triggerLevel)}");
                    }
                    break;

                case ScopeCommand.QueryTimeDiv:
                    if (double.TryParse(response, NumberStyles.Float, CultureInfo.InvariantCulture, out double timeDiv))
                    {
                        this.thisLastTimeDivSeconds = timeDiv;
                        this.AppendOutputLine("Info", $"TIME/DIV read as {ScopeFormatting.FormatTime(timeDiv)}");
                    }
                    break;

                case ScopeCommand.QueryVoltsDiv:
                    if (double.TryParse(response, NumberStyles.Float, CultureInfo.InvariantCulture, out double voltsDiv))
                    {
                        this.thisLastVoltsDivVolts = voltsDiv;
                        this.AppendOutputLine("Info", $"VOLTS/DIV read as {ScopeFormatting.FormatVoltage(voltsDiv)} per division");
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
                        this.AppendOutputLine("Info", $"Trigger level set to {ScopeFormatting.FormatVoltage(this.thisLastTriggerLevelVolts.Value)}");
                    }
                    break;

                case ScopeCommandPalette.SetTimeDiv:
                    if (this.thisLastTimeDivSeconds.HasValue)
                    {
                        this.AppendOutputLine("Info", $"TIME/DIV set to {ScopeFormatting.FormatTime(this.thisLastTimeDivSeconds.Value)}");
                    }
                    break;

                case ScopeCommandPalette.SetVoltsDiv:
                    if (this.thisLastVoltsDivVolts.HasValue)
                    {
                        this.AppendOutputLine("Info", $"VOLTS/DIV set to {ScopeFormatting.FormatVoltage(this.thisLastVoltsDivVolts.Value)} per division");
                    }
                    break;
            }
        }

        // ###########################################################################################
        // Returns a readable debug description for the currently executed command palette.
        // ###########################################################################################
        private static string GetPaletteDescription(ScopeCommandPalette palette)
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

                // The buffer above is drained and CLEARED by the flush 40ms later, so it is not
                // something a test can assert on without racing that timer. This second list is
                // never cleared and exists only when a test has opted in - see
                // TabOscilloscope.TestSeams.cs.
                this.thisRecordedOutputLinesForTests?.Add(line);

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
            this.AppendOutputLine("Info", $"Serial: {ScopeFormatting.MaskScopeSerial(serial)}");
            this.AppendOutputLine("Info", $"Firmware: {firmware}");
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
    }
}
