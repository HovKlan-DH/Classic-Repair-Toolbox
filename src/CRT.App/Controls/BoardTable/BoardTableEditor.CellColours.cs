using Avalonia;
using Avalonia.Controls;
using Avalonia.Data;
using Avalonia.Data.Converters;
using Avalonia.Input;
using Avalonia.Media;
using Avalonia.Styling;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // EACH CELL'S BACKGROUND AND TOOLTIP - the per-column cell theme whose Background binds to the
    // cell's state and worst problem (the state's wash, the problem's corner mark), and the cell
    // tooltips that open at once and let the pointer through. Split out of
    // BoardTableEditor.axaml.cs (code review, 2026-10-04: the main part had passed the ~1,500-line
    // mark). The file map is in BoardTableEditor.axaml.cs.
    // ###########################################################################################
    public partial class BoardTableEditor
    {
        // Maps a cell's state and its worst problem to its background - the state's wash, with the
        // problem's corner mark (see CellBrushFor). Its brushes are kept per theme, and dropped
        // when the columns are rebuilt, so a theme switch repaints.
        private readonly IMultiValueConverter thisCellToBrush;
        private readonly Dictionary<(BoardTableCellState State, BoardProblemLevel Problem), IBrush?> thisCellBrushes = [];

        private ControlTheme BuildCellTheme(int columnIndex)
        {
            var theme = new ControlTheme(typeof(DataGridCell)) { BasedOn = BoardTableEditor.DefaultCellTheme() };

            // The grid theme's cells are taller than the table's denser rows, which clipped the
            // current cell's frame to its left and right edges. The markup's CellsWrap style gives
            // them MinRowHeight, above this.
            theme.Setters.Add(new Setter(DataGridCell.MinHeightProperty, 0d));

            theme.Setters.Add(new Setter(DataGridCell.BackgroundProperty, this.CellBackgroundBinding(columnIndex)));

            // *** A SELECTED CELL KEEPS ITS OWN COLOUR (owner request, 2026-09-24). *** The
            // grid theme fills a selected cell with the accent colour, which read as one more
            // state - a green that could be taken for "added" - and hid the orange, green or red
            // of the cell underneath. The current cell is marked by its dashed frame instead (see
            // the markup). A trigger of our own, added after the grid theme's, outranks it.
            //
            // *** AND A CELL THE SEARCH FOUND KEEPS IT TOO (2026-10-02). *** The grid theme fills a
            // matching cell with its accent colour as well; the search marks only the text it found
            // (BoardTableEditor.Search.cs).
            //
            // Its text colour as well: the grid gives a matching cell's WHOLE text the mark's text
            // colour, which in the dark theme left the unmarked rest of it dark on a dark wash (seen
            // in a render). It takes its row's - the ordinary text colour - and the found run keeps
            // the mark's colours, which the cell text gives it itself.
            foreach (string pseudoClass in new[] { ":selected", ":searchmatch", ":searchcurrent" })
            {
                var keepsItsColour = new Style(selector => selector.Nesting().Class(pseudoClass));
                keepsItsColour.Setters.Add(new Setter(DataGridCell.BackgroundProperty, this.CellBackgroundBinding(columnIndex)));

                if (pseudoClass != ":selected")
                {
                    keepsItsColour.Setters.Add(new Setter(
                        DataGridCell.ForegroundProperty,
                        new Binding(nameof(DataGridRow.Foreground))
                        {
                            RelativeSource = new RelativeSource(RelativeSourceMode.FindAncestor) { AncestorType = typeof(DataGridRow) }
                        }));
                }

                theme.Children.Add(keepsItsColour);
            }

            // A file column shows the hover card instead, which says everything the text did - see
            // BoardTableEditor.FilePreview.cs. Two popups over one cell would cover each other.
            //
            // *** EXCEPT WHAT IS WRONG WITH IT (2026-10-02). *** A file that is not there, or is
            // spelled differently, is exactly what a file cell's error is about - and the card
            // cannot say it. So a file cell's tooltip is its problems alone (null with none), which
            // opens under the cell while the card sits beside it.
            theme.Setters.Add(new Setter(
                ToolTip.TipProperty,
                new Binding(this.PreviewsFilesIn(columnIndex)
                    ? $"Cells[{columnIndex}].ProblemToolTip"
                    : $"Cells[{columnIndex}].ToolTip")));

            BoardTableEditor.AddCellToolTipSetters(theme);

            return theme;
        }

        // The background of one column's cells: its state and its worst problem, together.
        private MultiBinding CellBackgroundBinding(int columnIndex) => new()
        {
            Bindings =
            {
                new Binding($"Cells[{columnIndex}].State"),
                new Binding($"Cells[{columnIndex}].ProblemLevel")
            },
            Converter = this.thisCellToBrush
        };

        private IBrush? CachedCellBrush(BoardTableCellState state, BoardProblemLevel problem)
        {
            if (!this.thisCellBrushes.TryGetValue((state, problem), out IBrush? brush))
            {
                brush = BoardTableEditor.CellBrushFor(state, problem);
                this.thisCellBrushes[(state, problem)] = brush;
            }

            return brush;
        }

        // ###########################################################################################
        // A cell's text tooltip - the value it replaced (BoardTableDocument.BaselineLabel names it,
        // "Published value" by default), a problem the checks found, what a marker means - shows AT ONCE (owner request, 2026-09-26: "the instant-show should also work for
        // texts - not only images"), as the file hover card does. The theme's default waits 400 ms.
        // It closes the moment the pointer leaves the cell, as every tooltip does.
        //
        // *** AND THE POINTER GOES STRAIGHT THROUGH IT (owner report, 2026-10-02: "the mouse should
        // follow the cell below, and not the tooltip"). *** A cell's tooltip opens under the cell,
        // over the next row, and Avalonia keeps a tooltip open while the pointer is ON it - so moving
        // down one row landed on the tooltip, and the cell below got neither its tooltip nor the
        // click. So the cell's tooltip is drawn in the window's overlay layer rather than as a
        // window of its own (a window cannot let the pointer through), and the editor's styles make
        // the tooltip itself invisible to the pointer (BoardTableEditor.axaml). Moving down then
        // simply moves onto the next cell, whose own tooltip takes over. The row grip's tooltip does
        // the same, in its style.
        // ###########################################################################################
        internal const int CellToolTipDelay = 0;

        private static void AddCellToolTipSetters(ControlTheme theme)
        {
            theme.Setters.Add(new Setter(ToolTip.ShowDelayProperty, BoardTableEditor.CellToolTipDelay));
            theme.Setters.Add(new Setter(ToolTip.ShouldUseOverlayLayerProperty, true));
        }

        // The grid theme's own cell theme, so ours only ADDS the background and tooltip rather than
        // replacing the cell's whole template.
        private static ControlTheme? DefaultCellTheme() =>
            Application.Current?.TryGetResource(typeof(DataGridCell), Application.Current.ActualThemeVariant, out object? theme) == true
                ? theme as ControlTheme
                : null;

        // ###########################################################################################
        // *** A PROBLEM IS A TRIANGLE IN THE CELL'S TOP-LEFT CORNER (owner request, 2026-10-02) ***
        // - red for an error, amber for a warning, as a spreadsheet marks a cell - over the state's
        // own wash, so a changed cell with an error still shows it is changed. It is ONE brush:
        // a gradient running diagonally from the corner to (CornerMarkSize, CornerMarkSize), the mark's colour
        // up to half way and the wash from there on (and on past the end, which a gradient pads
        // with its last colour). That draws a right triangle with CornerMarkSize-pixel sides in
        // the corner of a cell of any size - with no change to the grid's cell template, and on an
        // empty cell too, which a mark on the TEXT could never show.
        // ###########################################################################################
        internal const double CornerMarkSize = 10;

        internal static IBrush? CellBrushFor(BoardTableCellState state, BoardProblemLevel problem)
        {
            IBrush? wash = BoardTableEditor.BrushFor(state);

            if (problem == BoardProblemLevel.None)
            {
                return wash;
            }

            Color mark = (BoardTableEditor.MarkBrushFor(problem) as ISolidColorBrush)?.Color ?? Colors.Red;
            Color under = (wash as ISolidColorBrush)?.Color ?? Colors.Transparent;

            return new LinearGradientBrush
            {
                StartPoint = new RelativePoint(0, 0, RelativeUnit.Absolute),
                EndPoint = new RelativePoint(BoardTableEditor.CornerMarkSize, BoardTableEditor.CornerMarkSize, RelativeUnit.Absolute),
                GradientStops =
                {
                    new GradientStop(mark, 0),
                    new GradientStop(mark, 0.5),
                    new GradientStop(under, 0.5),
                    new GradientStop(under, 1)
                }
            };
        }

        // The colour of a problem's mark - the corner, the sheet tab's pill, the line above the table.
        internal static IBrush? MarkBrushFor(BoardProblemLevel problem) => problem switch
        {
            BoardProblemLevel.Error => ThemeResources.Resolve<IBrush>("BoardTable_Error_Mark", Brushes.Red),
            BoardProblemLevel.Warning => ThemeResources.Resolve<IBrush>("BoardTable_Warning_Mark", Brushes.Orange),
            _ => null,
        };

        internal static IBrush? BrushFor(BoardTableCellState state) => state switch
        {
            BoardTableCellState.Added => ThemeResources.Resolve<IBrush>("BoardTable_Added_Bg", Brushes.LightGreen),
            BoardTableCellState.Modified => ThemeResources.Resolve<IBrush>("BoardTable_Modified_Bg", Brushes.Orange),
            BoardTableCellState.Deleted => ThemeResources.Resolve<IBrush>("BoardTable_Deleted_Bg", Brushes.LightPink),
            _ => Brushes.Transparent,
        };
    }
}
