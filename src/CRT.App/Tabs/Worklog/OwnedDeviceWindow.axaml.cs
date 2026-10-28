using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // "Add device" and "Edit device" (CRT 4.0 step 4) - one form for both, as "Create new workbook"
    // is one form for creating and editing a workbook. Writes through OwnedDeviceStore and closes
    // with the record that reached disk (null when cancelled); a device that cannot be saved keeps
    // the form open with the reason.
    //
    // The hardware and board are picked from the boards CRT knows (the drop-down lists' own), the
    // board shown preselected for a new device. A device whose board CRT no longer lists keeps it.
    // ###########################################################################################
    public partial class OwnedDeviceWindow : Window
    {
        // How a device with no region reads in the drop-down.
        internal const string NoRegion = "Not set";

        private IReadOnlyList<HardwareBoardEntry> thisBoards = [];
        private OwnedDeviceRecord? thisEditing;

        // The record that reached disk - what the dialog closes with. For the tests, which show the
        // window without an owner and so cannot read a ShowDialog result.
        internal OwnedDeviceRecord? Saved { get; private set; }

        public OwnedDeviceWindow()
        {
            this.InitializeComponent();

            this.Opened += (_, _) =>
                Dispatcher.UIThread.Post(() => this.NameTextBox.Focus(), DispatcherPriority.Background);

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        // ###########################################################################################
        // A new device, on `boardKey` (the board shown) when CRT lists it.
        // ###########################################################################################
        public void Initialize(IReadOnlyList<HardwareBoardEntry> boards, string? boardKey)
        {
            this.thisBoards = boards ?? [];
            this.FillChoices(boardKey, region: string.Empty, OwnedDeviceStatus.Working);
            this.UpdateSaveEnabled();
        }

        // ###########################################################################################
        // An existing device: its fields filled in, "Save device" in place of "Add device".
        // ###########################################################################################
        public void InitializeForEdit(OwnedDeviceRecord device, IReadOnlyList<HardwareBoardEntry> boards)
        {
            this.thisEditing = device;
            this.thisBoards = boards ?? [];

            this.Title = "Edit device";
            this.HeaderText.Text = "Edit device";
            this.SaveButton.Content = "Save device";
            this.NameTextBox.Text = device.Name;
            this.SerialTextBox.Text = device.SerialNumber;
            this.OriginTextBox.Text = device.Origin;

            this.FillChoices(device.BoardKey, device.Region, device.Status);
            this.UpdateSaveEnabled();
        }

        private void FillChoices(string? boardKey, string region, string status)
        {
            (string hardware, string board) = Split(boardKey);

            List<string> hardwareNames = this.thisBoards
                .Select(entry => entry.HardwareName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            // A board CRT no longer lists stays the device's.
            if (hardware.Length > 0 && !hardwareNames.Contains(hardware, StringComparer.OrdinalIgnoreCase))
            {
                hardwareNames.Add(hardware);
            }

            this.HardwareComboBox.ItemsSource = hardwareNames;
            this.HardwareComboBox.SelectedItem = hardwareNames.FirstOrDefault(name => string.Equals(name, hardware, StringComparison.OrdinalIgnoreCase))
                ?? hardwareNames.FirstOrDefault();
            this.FillBoards(board);

            List<string> regions = [NoRegion, "PAL", "NTSC"];
            if (region.Trim().Length > 0 && !regions.Contains(region.Trim(), StringComparer.OrdinalIgnoreCase))
            {
                regions.Add(region.Trim());
            }

            this.RegionComboBox.ItemsSource = regions;
            this.RegionComboBox.SelectedItem = region.Trim().Length == 0
                ? NoRegion
                : regions.First(choice => string.Equals(choice, region.Trim(), StringComparison.OrdinalIgnoreCase));

            List<string> statuses = [.. OwnedDeviceStatus.All];
            string normalised = OwnedDeviceStatus.Normalise(status);
            if (!statuses.Contains(normalised))
            {
                statuses.Add(normalised);
            }

            this.StatusComboBox.ItemsSource = statuses;
            this.StatusComboBox.SelectedItem = normalised;
        }

        private void OnHardwareSelectionChanged(object? sender, SelectionChangedEventArgs e) => this.FillBoards(this.BoardComboBox.SelectedItem as string);

        // The boards of the hardware picked, keeping `keep` when that hardware has it.
        private void FillBoards(string? keep)
        {
            string hardware = this.HardwareComboBox.SelectedItem as string ?? string.Empty;

            List<string> boards = this.thisBoards
                .Where(entry => string.Equals(entry.HardwareName, hardware, StringComparison.OrdinalIgnoreCase))
                .Select(entry => entry.BoardName)
                .Where(name => !string.IsNullOrWhiteSpace(name))
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .ToList();

            // The device's own board, when CRT no longer lists it.
            if (this.thisEditing is { } device && Split(device.BoardKey) is var (ownHardware, ownBoard) &&
                string.Equals(ownHardware, hardware, StringComparison.OrdinalIgnoreCase) &&
                ownBoard.Length > 0 && !boards.Contains(ownBoard, StringComparer.OrdinalIgnoreCase))
            {
                boards.Add(ownBoard);
            }

            this.BoardComboBox.ItemsSource = boards;
            this.BoardComboBox.SelectedItem = boards.FirstOrDefault(name => string.Equals(name, keep, StringComparison.OrdinalIgnoreCase))
                ?? boards.FirstOrDefault();
        }

        // "{hardware}|{board}" in its two parts - the board after the first "|".
        private static (string Hardware, string Board) Split(string? boardKey)
        {
            string key = boardKey?.Trim() ?? string.Empty;
            int bar = key.IndexOf('|', StringComparison.Ordinal);
            return bar < 0 ? (key, string.Empty) : (key[..bar], key[(bar + 1)..]);
        }

        // What the form holds, as the store takes it.
        internal OwnedDeviceDetails Details()
        {
            string hardware = this.HardwareComboBox.SelectedItem as string ?? string.Empty;
            string board = this.BoardComboBox.SelectedItem as string ?? string.Empty;
            string region = this.RegionComboBox.SelectedItem as string ?? NoRegion;

            return new OwnedDeviceDetails(
                this.NameTextBox.Text?.Trim() ?? string.Empty,
                hardware.Length == 0 || board.Length == 0 ? string.Empty : $"{hardware}|{board}",
                region == NoRegion ? string.Empty : region,
                this.SerialTextBox.Text?.Trim() ?? string.Empty,
                this.OriginTextBox.Text?.Trim() ?? string.Empty,
                this.StatusComboBox.SelectedItem as string ?? OwnedDeviceStatus.Working);
        }

        private void OnNameTextChanged(object? sender, TextChangedEventArgs e) => this.UpdateSaveEnabled();

        private void UpdateSaveEnabled() =>
            this.SaveButton.IsEnabled = !string.IsNullOrWhiteSpace(this.NameTextBox.Text);

        // Escape cancels and Enter saves - no box here takes a line of its own. Enter on a BUTTON is
        // that button's (code review, 2026-10-28: Enter on a focused Cancel saved the form).
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close(null);
                e.Handled = true;
            }
            else if (e.Key == Key.Enter && e.Source is not Button &&
                     !(this.HardwareComboBox.IsDropDownOpen || this.BoardComboBox.IsDropDownOpen ||
                       this.RegionComboBox.IsDropDownOpen || this.StatusComboBox.IsDropDownOpen))
            {
                this.Save();
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private void OnSaveClick(object? sender, RoutedEventArgs e) => this.Save();

        private void Save()
        {
            OwnedDeviceDetails details = this.Details();

            if (details.Name.Length == 0)
            {
                this.ShowProblem(OwnedDeviceWording.NameNeeded);
                return;
            }

            OwnedDeviceRecord? saved = this.thisEditing is { } device
                ? OwnedDeviceStore.Update(device.Id, details)
                : OwnedDeviceStore.Create(details);

            if (saved is null)
            {
                // Kept open with what was typed - closing would read as cancelled.
                this.ShowProblem("Could not save the device - see the log for details.");
                return;
            }

            this.Saved = saved;
            this.Close(saved);
        }

        private void ShowProblem(string text)
        {
            this.ValidationText.Text = text;
            this.ValidationText.IsVisible = true;
            this.NameTextBox.Focus();
        }
    }
}
