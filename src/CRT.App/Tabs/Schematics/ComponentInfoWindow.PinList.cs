using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.Geometry;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // The pinout's list of pins: its rows - the pin of the image shown and the pin pointed at in the drawing lit - where it
    // stands beside or under the drawing, and how wide each of its columns is (PinTableGeometry). A pin clicked going to its
    // images is the main part's (GoToPin).
    // ###########################################################################################
    public partial class ComponentInfoWindow
    {
        // The wrapping columns' widths, as resources the cells' MaxWidth reads - so the rows made again as the pin shown or
        // pointed at changes take them too.
        private const string SignalWidthKey = "PopupPinSignalWidth";
        private const string NetWidthKey = "PopupPinNetWidth";
        private const string ReadingWidthKey = "PopupPinReadingWidth";

        // The markup's size of a cell's text; a heading's is read off the heading itself.
        private const double PinCellFontSize = 11;

        // The pins listed beside the drawing - InfoPinsItemsControl's, changed in place (ShowPinRows).
        private readonly System.Collections.ObjectModel.ObservableCollection<ComponentPopupPin> _pinRows = new();

        // The pin the pointer is on in the drawing - its row lit as the drawing lights it - null for none.
        private int? _pinPointedAtInDrawing;

        // The list's columns as last measured - Pin, Signal, Net name, Direction, Expected oscilloscope reading - and the
        // widths of the texts measured for them, kept while the window is open: the same pins come back as the regions and
        // components are gone through.
        private IReadOnlyList<PinTableColumn> _pinColumns = [];
        private readonly Dictionary<(string Text, double Size, bool Bold), double> _pinTextWidths = new();

        internal IReadOnlyList<PinTableColumn> PinColumnsForTests => this._pinColumns;

        // ###########################################################################################
        // THE PIN OF THE IMAGE SHOWN IS LIT, on the drawing and in the list beside it (owner request,
        // 2026-11-02: "if I am viewing a specific pin, then do highlight it on the pinout image and in the
        // right-side text"). The list keeps every row that does not change: a row is pressed to go to its
        // images, which changes the selection, and one made again under the pointer loses its hover.
        // ###########################################################################################
        private void ShowViewedPin(ComponentImageItem? shown)
        {
            string? pin = shown?.Pin;

            this.ShowPinRows(pin);
            this.InfoPinoutDrawing.MarkPin(ComponentVocabulary.PinNumberOf(pin));
            this.InfoPartDrawing.MarkPin(ComponentVocabulary.PinNumberOf(pin));
        }

        // Every row as it is now - `viewedPin` the pin of the image shown - each with its expected oscilloscope reading,
        // the one pointed at in the drawing lit.
        private void ShowPinRows(string? viewedPin) =>
            InPlaceList.KeepUnchanged(
                this._pinRows,
                ComponentPopupContent.PinsOf(
                    this._allPins,
                    this._boardLabel,
                    this._localRegion,
                    this._pinNets,
                    viewedPin,
                    this._allComponentImages,
                    this._pinPointedAtInDrawing),
                (row, wanted) => row == wanted);

        // ###########################################################################################
        // THE PIN UNDER THE POINTER IN THE DRAWING IS LIT IN THE LIST TOO (owner request, 2026-11-09: "hovering the IC
        // image pin will not highlight the table pin. This needs to be the same both places ... Not only an issue for
        // ICs"), as the list's row under the pointer lights the drawing (OnPinRowEntered). Only the rows that change are
        // replaced: the pointer is on the drawing, so no row is made again under it.
        // ###########################################################################################
        private void OnPinHoveredInDrawing(object? sender, PinHoveredEventArgs e)
        {
            if (this._pinPointedAtInDrawing != e.Pin)
            {
                this._pinPointedAtInDrawing = e.Pin;
                this.ShowPinRows((this.ThumbnailList.SelectedItem as ComponentImageItem)?.Pin);
            }
        }

        // The pin under the pointer in the list is the pin marked in the drawing.
        private void OnPinRowEntered(object? sender, PointerEventArgs e)
        {
            if (sender is Control { Tag: ComponentPopupPin pin })
            {
                this.InfoPinoutDrawing.HoverPin(ComponentVocabulary.PinNumberOf(pin.Pin));
                this.InfoPartDrawing.HoverPin(ComponentVocabulary.PinNumberOf(pin.Pin));
            }
        }

        private void OnPinRowExited(object? sender, PointerEventArgs e)
        {
            this.InfoPinoutDrawing.HoverPin(null);
            this.InfoPartDrawing.HoverPin(null);
        }

        // The pin list beside the drawing while the popup has the room for it, else under it.
        private void OnPinoutGridSizeChanged(object? sender, SizeChangedEventArgs e) => this.PlacePinList(e.NewSize.Width);

        // ###########################################################################################
        // The pin list beside the drawing when it fits there with every text whole (PinoutGeometry.ListFitsBeside), else
        // under it - and a RESISTOR'S LIST FIRST, its picture under it (owner request, 2026-11-09: "I want the pinout table
        // to be shown first and then the image and explanation afterwards"): the list across the top row, the picture in the
        // second (its place in the markup). Its columns then fit the room it has (FitPinColumns).
        // ###########################################################################################
        private void PlacePinList(double width)
        {
            double drawing = this.InfoPinoutDrawing.IsVisible ? this.InfoPinoutDrawing.LayoutForTests?.Size.Width ?? 0
                : this.InfoPartDrawing.IsVisible ? this.InfoPartDrawing.NaturalSize.Width
                : 0;
            bool first = this.InfoResistorPicture.IsVisible;
            bool beside = !first && PinoutGeometry.ListFitsBeside(width, drawing, PinTableGeometry.WidthOf(this._pinColumns));

            Grid.SetColumn(this.InfoPinsPanel, beside ? 1 : 0);
            Grid.SetColumnSpan(this.InfoPinsPanel, beside ? 1 : 2);
            Grid.SetRow(this.InfoPinsPanel, beside || first ? 0 : 1);

            this.FitPinColumns(beside && drawing > 0 ? width - drawing - PinoutGeometry.ListGap : width);
        }

        // ###########################################################################################
        // EVERY COLUMN AS WIDE AS ITS LONGEST TEXT, THE SAME SPACE BETWEEN EACH TWO (owner request, 2026-11-09: "align on
        // empty spacing between columns - it should be the same for all, and I do like the spacing between 'Pin' and
        // 'Signal'. Should be determined by longest text in each column + that spacing"). The markup sizes each column by
        // its texts and puts PinTableGeometry.Gap after every cell; what is measured here is for the room: the list's width
        // to place it, and - when the room is short - how far the signal, net and reading columns wrap. The headings say
        // which columns are shown, as before.
        // ###########################################################################################
        private void MeasurePinColumns()
        {
            List<ComponentPopupPin> rows = [.. this._pinRows];
            bool nets = rows.Any(pin => pin.HasNet);
            bool readings = rows.Any(pin => pin.HasReading);

            this.InfoPinsNetHeading.IsVisible = nets;
            this.InfoPinsReadingHeading.IsVisible = readings;

            if (this._pinTextWidths.Count > 5000)
            {
                this._pinTextWidths.Clear();
            }

            this._pinColumns = rows.Count == 0
                ? []
                :
                [
                    this.PinColumnOf(0, rows.Select(pin => pin.Pin), wraps: false),
                    this.PinColumnOf(1, rows.Select(pin => pin.Name), wraps: true),
                    nets ? this.PinColumnOf(2, rows.Select(pin => pin.Net), wraps: true) : default,
                    this.PinColumnOf(3, rows.Select(pin => pin.Direction), wraps: false),
                    readings ? this.PinColumnOf(4, rows.Select(pin => pin.Reading), wraps: true) : default,
                ];
        }

        // One column: its heading's lines and its cells' lines measured, and for a column that wraps, each word of them.
        private PinTableColumn PinColumnOf(int column, IEnumerable<string> cells, bool wraps)
        {
            TextBlock heading = this.InfoPinsHeading.Children.OfType<TextBlock>().First(block => Grid.GetColumn(block) == column);
            bool bold = heading.FontWeight >= FontWeight.SemiBold;

            List<(string Text, double Size, bool Bold)> lines =
            [
                .. (heading.Text ?? string.Empty).Split('\n').Select(line => (line, heading.FontSize, bold)),
                .. cells.SelectMany(cell => cell.Split('\n')).Distinct(StringComparer.Ordinal).Select(line => (line, PinCellFontSize, false)),
            ];

            double widest = lines.Select(this.PinTextWidth).DefaultIfEmpty(0).Max();
            double widestWord = !wraps
                ? widest
                : lines
                    .SelectMany(line => PinTableGeometry.PiecesOf(line.Text).Select(word => (word, line.Size, line.Bold)))
                    .Distinct()
                    .Select(this.PinTextWidth)
                    .DefaultIfEmpty(0)
                    .Max();

            return new PinTableColumn(widest, widestWord, wraps);
        }

        // A text's width as a cell of the list draws it - in the window's font, measured once.
        private double PinTextWidth((string Text, double Size, bool Bold) text)
        {
            if (text.Text.Length == 0)
            {
                return 0;
            }

            if (!this._pinTextWidths.TryGetValue(text, out double width))
            {
                var block = new TextBlock
                {
                    Text = text.Text,
                    FontFamily = this.FontFamily,
                    FontSize = text.Size,
                    FontWeight = text.Bold ? FontWeight.Bold : FontWeight.Normal,
                };

                block.Measure(Size.Infinity);
                width = block.DesiredSize.Width;
                this._pinTextWidths[text] = width;
            }

            return width;
        }

        // ###########################################################################################
        // The signal, net and reading columns in `room`: as wide as their texts while the list fits, narrower - wrapping -
        // only when it does not (PinTableGeometry.TextWidths). A column at its widest has no limit at all, so a text measured
        // a hair narrower than it draws never wraps.
        // ###########################################################################################
        private void FitPinColumns(double room)
        {
            IReadOnlyList<double> widths = PinTableGeometry.TextWidths(this._pinColumns, room);

            void Fit(string key, int column)
            {
                double limit = column < widths.Count && widths[column] < this._pinColumns[column].Widest
                    ? widths[column]
                    : double.PositiveInfinity;

                // Set only when it moves: every cell of the column is told of a resource set.
                if (!this.InfoPinsPanel.Resources.TryGetValue(key, out object? held) || !Equals(held, limit))
                {
                    this.InfoPinsPanel.Resources[key] = limit;
                }
            }

            Fit(SignalWidthKey, 1);
            Fit(NetWidthKey, 2);
            Fit(ReadingWidthKey, 4);
        }
    }
}
