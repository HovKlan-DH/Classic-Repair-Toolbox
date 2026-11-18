using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Layout;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Linq;
using Tabs.TabSchematics;

namespace CRT;

// ###########################################################################################
// THE PINS OF THE COMPONENTS SELECTED OR POINTED AT, ON A BOARD'S SCHEMATIC IMAGE (owner requests, 2026-11-17), with the
// board's KiCad pads where they are - the components whose pin 1 is marked:
//
//   - "Hold SHIFT to number pins on hover and select" (ticked until unticked): while SHIFT is held, every pad numbered on it
//     (PinNumbersOverlay), pin 1 told apart among the numbers.
//   - "Show color for pin direction" (ticked until unticked; only while the pins are numbered): each number in its pin's
//     direction's colors, the pinout's ("identical coloring as from its pinout" - ComponentPinDirections) - and A LEGEND
//     saying what they mean while they are on screen (owner question, the same day: "I think we need a legend somewhere
//     that can explain the colors?"; PinDirectionLegend): a panel right of "Important KiCad signals", one a line ("show
//     vertically and right to the "Important KiCad signals" panel"), every direction - always all of them, so it stands
//     still as the pointer goes from component to component - and pin 1's ring.
//
// "Show gates inside the ICs" - each IC's pinout drawing laid over it, pin on pad - went the same day (owner decision: "I do
// NOT want the functionaly ... It becomes too weird to look at - at least for now"); it is in git history (2026-11-17).
//
// Worked out again as the KiCad overlay is (RefreshKiCadOverlayNow), as the component pointed at changes and as SHIFT does.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    // What the legend calls pin 1's ring.
    internal const string LegendFirstPin = "Pin 1";

    // The checkboxes as stored, what a change of each does, and the legend's items - called once, by the constructor.
    private void WirePinsOnImage()
    {
        this.CheckGlobalNumberPins.IsChecked = UserSettings.SchematicsNumberPinsHoldingShift;
        this.CheckGlobalShowPinDirection.IsChecked = UserSettings.SchematicsShowPinDirection;

        this.CheckGlobalNumberPins.IsCheckedChanged += (_, _) =>
        {
            UserSettings.SchematicsNumberPinsHoldingShift = this.CheckGlobalNumberPins.IsChecked == true;

            // "Show color for pin direction" follows it.
            this.UpdateInteractiveCadTraceHoverModeUi();
            this.RefreshPinsOnImage();
        };

        this.CheckGlobalShowPinDirection.IsCheckedChanged += (_, _) =>
        {
            UserSettings.SchematicsShowPinDirection = this.CheckGlobalShowPinDirection.IsChecked == true;
            this.RefreshPinsOnImage();
        };

        // Every direction, in the order every legend of them lists them (PinDirectionColors.LegendOrder).
        foreach (PinDirectionKind kind in PinDirectionColors.LegendOrder)
        {
            var swatch = new Border { Width = 12, Height = 12, CornerRadius = new CornerRadius(2), BorderThickness = new Thickness(1) };
            ThemeResources.Paint(swatch, Border.BackgroundProperty, PinDirectionColors.FillKeyOf(kind));
            ThemeResources.Paint(swatch, Border.BorderBrushProperty, PinDirectionColors.InkKeyOf(kind));

            this.PinDirectionLegendItems.Children.Add(TabSchematics.LegendItem(swatch, PinDirectionColors.WordOf(kind)));
        }

        // Pin 1's ring as its badge has it: the first pin's orange, a black line each side of it (PinNumbersOverlay). Black in
        // both themes ON PURPOSE, never a theme's color: the badge is drawn over the board's own image, whatever the theme, so
        // its lines are black there - the swatch shows the very ring the board shows (code review, 2026-11-18, of the fixed
        // Brushes.Black). The orange and its fill are the theme's, as on the badge.
        var inner = new Border { BorderThickness = new Thickness(1) };
        ThemeResources.Paint(inner, Border.BackgroundProperty, PinDirectionColors.FillKeyOf(PinDirectionKind.Other));
        inner.BorderBrush = Avalonia.Media.Brushes.Black;

        var ring = new Border { BorderThickness = new Thickness(2), Child = inner };
        ThemeResources.Paint(ring, Border.BorderBrushProperty, "Schematics_FirstPin");

        var firstPin = new Border { Width = 14, Height = 14, BorderThickness = new Thickness(1), BorderBrush = Avalonia.Media.Brushes.Black, Child = ring };

        this.PinDirectionLegendItems.Children.Add(TabSchematics.LegendItem(firstPin, LegendFirstPin));
    }

    // A swatch and its word, side by side.
    private static StackPanel LegendItem(Control swatch, string word)
    {
        swatch.VerticalAlignment = VerticalAlignment.Center;

        var text = new TextBlock { Text = word, FontSize = 11, VerticalAlignment = VerticalAlignment.Center };
        ThemeResources.Paint(text, TextBlock.ForegroundProperty, "Schematics_Panels_Fg");

        return new StackPanel { Orientation = Orientation.Horizontal, Spacing = 4, Children = { swatch, text } };
    }

    // The legend's words as shown, after its lead - for the tests.
    internal IReadOnlyList<string> PinDirectionLegendWordsForTests =>
        [.. this.PinDirectionLegendItems.Children.OfType<StackPanel>().Select(item => item.Children.OfType<TextBlock>().Single().Text ?? string.Empty)];

    private void OnGlobalNumberPinsRowClicked(object? sender, PointerPressedEventArgs e) => this.ToggleSettingRow(this.CheckGlobalNumberPins, e);

    private void OnGlobalShowPinDirectionRowClicked(object? sender, PointerPressedEventArgs e) => this.ToggleSettingRow(this.CheckGlobalShowPinDirection, e);

    // A setting row's own check box ticked or unticked by a left press anywhere on the row - as every row of the panel is.
    private void ToggleSettingRow(CheckBox box, PointerPressedEventArgs e)
    {
        if (!box.IsEnabled || !e.GetCurrentPoint(this).Properties.IsLeftButtonPressed)
        {
            return;
        }

        box.IsChecked = box.IsChecked != true;
        e.Handled = true;
    }

    // ###########################################################################################
    // The two rows shown with the board's KiCad pads on the schematic image - "Show color for pin direction" greyed out while
    // the pins are not numbered (owner request, 2026-11-17: "make sure to enable the below "Show pin direction" checkbox, and
    // likewise disable it, if this checkbox is not checked"). From UpdateInteractiveCadTraceHoverModeUi.
    // ###########################################################################################
    private void UpdatePinsOnImageRows(bool hasPadData)
    {
        bool numbering = this.CheckGlobalNumberPins.IsChecked == true;

        TabSchematics.ShowSettingRow(this.GlobalNumberPinsRow, this.CheckGlobalNumberPins, visible: hasPadData, enabled: hasPadData);
        TabSchematics.ShowSettingRow(this.GlobalShowPinDirectionRow, this.CheckGlobalShowPinDirection, visible: hasPadData, enabled: hasPadData && numbering);
    }

    internal void UpdatePinsOnImageRowsForTests(bool hasPadData) => this.UpdatePinsOnImageRows(hasPadData);

    private static void ShowSettingRow(StackPanel row, CheckBox box, bool visible, bool enabled)
    {
        row.IsVisible = visible;
        row.IsEnabled = enabled;
        row.Opacity = enabled ? 1.0 : 0.55;
        row.Cursor = enabled ? new Cursor(StandardCursorType.Hand) : Cursor.Default;
        box.IsEnabled = enabled;
    }

    // ###########################################################################################
    // Whether the pins are numbered now: the setting on, and SHIFT held - but never while a tool works on the image, whose
    // SHIFT is its own (code review, 2026-11-18): the label editor's turns snapping off, and the numbers - over its overlay -
    // covered the very rectangle being drawn; a worklog's area and "Draw trace" are drawn there too.
    // ###########################################################################################
    private bool IsNumberingPins =>
        UserSettings.SchematicsNumberPinsHoldingShift &&
        this.thisIsInteractiveCadTraceHoverShiftPressed &&
        !this.thisIsLabelEditorMode &&
        !this.thisIsWorklogEntryMode &&
        !this.thisIsDrawingTraces;

    // SHIFT pressed or let go of wherever the keyboard is in CRT's window - Main hands it over while this tab is shown, as the
    // component filter outside the tab usually holds the keyboard.
    internal void NoteKeyModifiers(KeyModifiers modifiers) => this.UpdateInteractiveCadTraceHoverShiftState(modifiers);

    // Whether SHIFT is held as the tab last heard it.
    internal bool IsShiftHeld => this.thisIsInteractiveCadTraceHoverShiftPressed;

    // ###########################################################################################
    // The numbers, for what is selected and pointed at now - else their pin 1's orange pads framed in black (owner request,
    // 2026-11-17: "the "normal" first pin marking having this 1px 2xblack border so it is more visible"; a numbered pin 1's
    // badge has its own ring) - and the legend while any number is in its direction's colors.
    // ###########################################################################################
    private void RefreshPinsOnImage()
    {
        (IReadOnlyList<PinNumberBadge> badges, IReadOnlyList<FirstPinMark> firstPins) = this.PinsNow();

        // Asked at every component the pointer passes: nothing drawn again while there is nothing to draw.
        if (badges.Count == 0 && firstPins.Count == 0)
        {
            this.SchematicsPinNumbersOverlay.Clear();
        }
        else
        {
            this.SchematicsPinNumbersOverlay.Show(badges, firstPins);
        }

        this.PinDirectionLegend.IsVisible = badges.Any(badge => badge.Direction is not null);
    }

    // Every pad of the components selected or pointed at, numbered while the pins are - else each one's pin 1 to frame,
    // the very pad the KiCad overlay paints orange (IsPrimaryPadForPin1Highlight).
    private (IReadOnlyList<PinNumberBadge> Badges, IReadOnlyList<FirstPinMark> FirstPins) PinsNow()
    {
        HashSet<string> targets = new(this.thisSelectedKiCadReferences, StringComparer.OrdinalIgnoreCase);

        if (!string.IsNullOrWhiteSpace(this.thisHoveredComponentBoardLabel))
        {
            targets.Add(this.thisHoveredComponentBoardLabel.Trim());
        }

        // Nothing selected or pointed at, or the KiCad data being aligned: nothing to look up either - this runs at every rebuild
        // of the KiCad overlay and every SHIFT pressed anywhere in CRT's window (code review, 2026-11-18).
        if (targets.Count == 0 || this.thisIsKiCadTraceCalibrationMode || this.thisKiCadProject is null || this.currentFullResBitmap is null)
        {
            return ([], []);
        }

        // The view looked up ONCE: a side of the board, its PCB there - pin 1 is marked on every such image (IsMarkingFirstPin
        // asks the same), so its pads are always worth going through.
        KiCadProjectView? view = this.ResolveKiCadViewForCurrentSchematic();
        bool isPcb = view is not null &&
            (string.Equals(view.Type, "pcb_top", StringComparison.OrdinalIgnoreCase) ||
             string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase));
        bool numbering = this.IsNumberingPins;

        if (!isPcb || this.thisKiCadProject.Root.Pcb.ElementAtOrDefault(view!.SourceIndex) is not { } pcb)
        {
            return ([], []);
        }

        Rect contentRect = this.GetImageContentRect();
        Rect worldBounds = this.GetKiCadPcbWorldBounds(pcb);

        if (contentRect.Width <= 0 || contentRect.Height <= 0 || worldBounds.Width <= 0 || worldBounds.Height <= 0)
        {
            return ([], []);
        }

        KiCadViewCalibration calibration = this.GetKiCadViewCalibration(this.GetCurrentSchematicName());
        string layer = string.Equals(view.Type, "pcb_bottom", StringComparison.OrdinalIgnoreCase) ? "B.Cu" : "F.Cu";
        bool directions = UserSettings.SchematicsShowPinDirection;

        var badges = new List<PinNumberBadge>();
        var firstPins = new List<FirstPinMark>();

        foreach (KiCadPcbFootprint footprint in pcb.Footprints)
        {
            string reference = footprint.Reference?.Trim() ?? string.Empty;

            if (reference.Length == 0 || !targets.Contains(reference))
            {
                continue;
            }

            // Not numbered: pin 1's pad framed, sized and turned as the KiCad overlay draws it - by its very rule (PadOnImage).
            if (!numbering)
            {
                foreach (KiCadPcbPad pad in footprint.Pads.Where(pad =>
                    KiCadLayerGeometry.IsPointVisibleOnSide(pad.Layers, layer) &&
                    TabSchematics.IsPrimaryPadForPin1Highlight(footprint, pad, layer)))
                {
                    if (this.PadOnImage(pad, worldBounds, contentRect, calibration) is { } placed)
                    {
                        firstPins.Add(new FirstPinMark(placed.Rect.Center, placed.Rect.Size, placed.RotationDegrees, placed.IsRectangle));
                    }
                }

                continue;
            }

            List<KiCadPcbPad> pads = footprint.Pads
                .Where(pad => pad.AbsoluteCenter is not null && !string.IsNullOrWhiteSpace(pad.Number) && KiCadLayerGeometry.IsPointVisibleOnSide(pad.Layers, layer))
                .ToList();

            if (pads.Count == 0)
            {
                continue;
            }

            ComponentPinDirections pins = directions ? this.PinDirectionsOf(reference) : ComponentPinDirections.None;
            string firstPin = PinBadgeGeometry.FirstPinOf(pads.Select(pad => pad.Number!));

            foreach (KiCadPcbPad pad in pads)
            {
                string number = pad.Number!.Trim();
                Point center = this.MapKiCadWorldToLocal(pad.AbsoluteCenter!.X, pad.AbsoluteCenter.Y, worldBounds, contentRect, calibration);
                double size = MapKiCadWorldLengthToLocal(Math.Max(pad.Size?.X ?? 1.2, pad.Size?.Y ?? 1.2), worldBounds, contentRect, calibration);

                badges.Add(new PinNumberBadge(
                    center,
                    size,
                    number,
                    string.Equals(number, firstPin, StringComparison.OrdinalIgnoreCase),
                    directions ? pins.KindOf(number) : null));
            }
        }

        return (badges, firstPins);
    }
}
