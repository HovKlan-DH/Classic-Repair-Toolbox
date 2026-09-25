using Avalonia.Controls;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.Platform.Storage;
using Avalonia.Threading;
using Handlers.DataHandling;
using Handlers.Oscilloscope;
using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Diagnostics.CodeAnalysis;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // Oscilloscope image capture: the output-folder picker, the dump-image SCPI workflow, and
    // saving a captured image to disk (including the "run every palette" full test suite, which
    // drives this workflow for each palette in order). See TabOscilloscope.axaml.cs for the file
    // map of the whole partial class.
    // ###########################################################################################
    public partial class TabOscilloscope
    {
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
        // Executes the image dump workflow with the same preparation and cleanup pattern as shown.
        // ###########################################################################################
        private async Task ExecuteDumpImageWorkflowAsync(
            IScopeClient scopeClient,
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

            if (ScopePayloadParser.TryExtractBinaryPayload(rawData, out byte[] payload) &&
                ScopePayloadParser.TryReadBmpMetadata(payload, out int width, out int height, out short bitsPerPixel))
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

            string outputFilePath = BuildCapturedOscilloscopeImageFilePath(
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

                if (!TryCreateBitmapFromScopeRawImageData(rawImageData, out Bitmap? capturedBitmap))
                {
                    this.AppendOutputLine("Warning", "Could not decode the oscilloscope image");
                    return null;
                }

                using (capturedBitmap)
                {
                    Directory.CreateDirectory(outputDirectory);
                    capturedBitmap.Save(outputFilePath, new PngBitmapEncoderOptions());
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
            IScopeClient scopeClient,
            OscilloscopeEntry selectedOscilloscope,
            CancellationToken cancellationToken)
        {
            this.AppendOutputLine("Debug", "---");
            this.AppendOutputLine("Debug", GetPaletteDescription(ScopeCommandPalette.DumpImage));

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
        private static bool TryCreateBitmapFromScopeRawImageData(byte[] rawImageData, [NotNullWhen(true)] out Bitmap? bitmap)
        {
            bitmap = null;

            try
            {
                if (ScopePayloadParser.TryExtractBinaryPayload(rawImageData, out byte[] payload) && payload.Length > 0)
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
        private static string BuildCapturedOscilloscopeImageFilePath(
            ComponentImageEntry componentImageEntry,
            string displayedRegion,
            string outputDirectory)
        {
            string safeBoardLabel = ScopeFormatting.SanitizeCapturedOscilloscopeImageFileNamePart(componentImageEntry.BoardLabel);

            string identityPart = !string.IsNullOrWhiteSpace(componentImageEntry.Pin)
                ? ScopeFormatting.SanitizeCapturedOscilloscopeImageFileNamePart(componentImageEntry.Pin)
                : ScopeFormatting.SanitizeCapturedOscilloscopeImageFileNamePart(componentImageEntry.Name);

            string safeRegion = string.IsNullOrWhiteSpace(displayedRegion)
                ? string.Empty
                : ScopeFormatting.SanitizeCapturedOscilloscopeImageFileNamePart(displayedRegion);

            string fileName = string.IsNullOrWhiteSpace(safeRegion)
                ? $"{safeBoardLabel}_{identityPart}.png"
                : $"{safeBoardLabel}_{identityPart}_{safeRegion}.png";

            return Path.Combine(outputDirectory, fileName);
        }
    }
}
