using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Media;
using CRT;
using DataHandling;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text.Json;
using TabSchematics;

namespace Classic_Repair_Toolbox.TabSchematics
{
    // ###########################################################################################
    // Manages the creation, interaction, and rendering of polylines drawn on the schematics.
    // ###########################################################################################
    internal class PolylineManagement
    {
        private readonly Canvas _canvas;
        private readonly CRT.TabSchematics _parent;
        private readonly List<ManagedPolyline> _polylines = new();

        private ManagedPolyline? _activePolyline;
        private int _draggingNodeIndex = -1;

        private bool _isDrawingNew;
        private Point _drawingStartPoint;
        private ManagedPolyline? _tempDrawingLine;

        private Ellipse? _hoverMarker;

        private readonly Dictionary<Color, bool> _colorVisibilityOptions = new();
        private static readonly Dictionary<string, Stack<TraceModel>> _globalUndoStacks = new();

        private Stack<TraceModel> GetCurrentUndoStack()
        {
            var boardKey = this._parent.MainWindow?.GetCurrentBoardKey() ?? string.Empty;
            var schematicName = (this._parent.SchematicsThumbnailList.SelectedItem as SchematicThumbnail)?.Name ?? string.Empty;
            string key = $"{boardKey}|{schematicName}";

            if (!_globalUndoStacks.TryGetValue(key, out var stack))
            {
                stack = new Stack<TraceModel>();
                _globalUndoStacks[key] = stack;
            }
            return stack;
        }


        private Color _currentDrawingColor = Colors.Red;
        public Color CurrentDrawingColor
        {
            get => this._currentDrawingColor;
            set
            {
                if (this._currentDrawingColor != value)
                {
                    this._currentDrawingColor = value;
                    this.ActiveColorChanged?.Invoke(value);
                }
            }
        }

        public List<Color> PaletteColors { get; private set; } = new();

        public event Action<Color>? ActiveColorChanged;
        public event Action<List<Color>>? PaletteColorsChanged;
        public event Action<Dictionary<Color, int>>? TraceStatsChanged;
        public event Action<bool, Point>? PaletteStateChanged;
        public event Action? TracesModified; // notifies parent to save JSON
        public event Action<bool>? UndoStateChanged;

        public PolylineManagement(CRT.TabSchematics parent, Canvas canvas)
        {
            this._parent = parent;
            this._canvas = canvas;
            this.PaletteColors = this.GetDefaultPaletteColors();
            this.CurrentDrawingColor = this.PaletteColors.FirstOrDefault();
        }

        // ###########################################################################################
        // Retrieve theme-aware default colors from the parent control's resources.
        // ###########################################################################################
        private List<Color> GetDefaultPaletteColors()
        {
            Color GetColor(string key, Color fallback)
            {
                object? res = null;

                // 1. Try finding it within the logical tree context
                if (this._parent.TryFindResource(key, out var parentRes))
                {
                    res = parentRes;
                }
                // 2. Try global Application scope directly using the actively running Theme
                else if (Application.Current != null)
                {
                    var theme = Application.Current.ActualThemeVariant;
                    if (Application.Current.TryGetResource(key, theme, out var appRes))
                    {
                        res = appRes;
                    }
                }

                if (res != null)
                {
                    if (res is Color c) return c;
                    if (res is ISolidColorBrush brush) return brush.Color;
                    if (res is string s && Color.TryParse(s, out var parsed)) return parsed;
                }

                return fallback;
            }

            Color c1 = GetColor("Schematics_TracePalette_Color1", Colors.Red);
            Color c2 = GetColor("Schematics_TracePalette_Color2", Colors.DodgerBlue);
            Color c3 = GetColor("Schematics_TracePalette_Color3", Colors.LimeGreen);
            Color c4 = GetColor("Schematics_TracePalette_Color4", Colors.Yellow);

            return new List<Color> { c1, c2, c3, c4 };
        }

        public void AddOrReplacePaletteColor(Color newColor)
        {
            if (this.PaletteColors.Contains(newColor))
                return;

            var usedColors = this._polylines.Select(p => p.TraceColor).ToHashSet();

            int unusedTargetIndex = -1;
            for (int i = 0; i < this.PaletteColors.Count; i++)
            {
                if (!usedColors.Contains(this.PaletteColors[i]))
                {
                    unusedTargetIndex = i;
                    break;
                }
            }

            if (unusedTargetIndex != -1) this.PaletteColors[unusedTargetIndex] = newColor;
            else this.PaletteColors.Add(newColor);

            this.PaletteColorsChanged?.Invoke(this.PaletteColors);
        }

        public void ChangeActiveColor(Color color)
        {
            this.CurrentDrawingColor = color;

            if (this._activePolyline != null)
            {
                this._activePolyline.SetColor(color);
                this.NotifyStatsChanged();
                this.TracesModified?.Invoke(); // Trigger save
            }
        }

        private TraceModel ExportSingleTrace(ManagedPolyline p)
        {
            var tm = new TraceModel
            {
                Color = p.TraceColor.ToString(),
                Visible = this.GetColorVisibility(p.TraceColor)
            };
            for (int i = 0; i < p.NodeCount; i++)
            {
                var n = p.GetNode(i);
                tm.Nodes.Add(new PointModel { X = n.X, Y = n.Y });
            }
            return tm;
        }

        private void PushUndo(ManagedPolyline p)
        {
            var tm = this.ExportSingleTrace(p);
            var stack = this.GetCurrentUndoStack();
            stack.Push(tm);
            this.UndoStateChanged?.Invoke(true);
        }

        // ###########################################################################################
        // Pops the last deleted trace from the temporary undo stack and restores it natively.
        // ###########################################################################################
        public void UndoLastDeletion()
        {
            var stack = this.GetCurrentUndoStack();
            if (stack.Count > 0)
            {
                var tm = stack.Pop();
                Logger.Info($"Undo deletion of trace: Color [{tm.Color}], Markers [{tm.Nodes.Count}]");
                this.AddTraceFromModel(tm);
                this.NotifyStatsChanged();
                this.TracesModified?.Invoke();
                this.UndoStateChanged?.Invoke(stack.Count > 0);
            }
        }

        public void DeleteActivePolyline()
        {
            if (this._activePolyline != null)
            {
                this.PushUndo(this._activePolyline);
                Logger.Info($"Trace deleted (Palette): Color [{this._activePolyline.TraceColor}], Markers [{this._activePolyline.NodeCount}]");

                this._polylines.Remove(this._activePolyline);
                this._activePolyline.Dispose(this._canvas);
                this._activePolyline = null;

                this.PaletteStateChanged?.Invoke(false, default);
                this.NotifyStatsChanged();
                this.TracesModified?.Invoke(); // Trigger save
            }
        }

        public bool GetColorVisibility(Color color)
        {
            return this._colorVisibilityOptions.TryGetValue(color, out bool v) ? v : true;
        }

        public void SetVisibilityByColor(Color targetColor, bool visible)
        {
            this._colorVisibilityOptions[targetColor] = visible;

            foreach (var p in this._polylines.Where(x => x.TraceColor == targetColor))
                p.SetGlobalVisibility(visible);

            if (this._tempDrawingLine != null && this._tempDrawingLine.TraceColor == targetColor)
                this._tempDrawingLine.SetGlobalVisibility(visible);

            if (!visible && this._activePolyline?.TraceColor == targetColor)
            {
                this.PaletteStateChanged?.Invoke(false, default);
                this.DeselectAll();
            }

            // Fire the save event so the new 'visible' boolean gets written out to disk immediately
            this.TracesModified?.Invoke();
        }

        private void NotifyStatsChanged()
        {
            var stats = new Dictionary<Color, int>();
            foreach (var p in this._polylines)
            {
                if (!stats.ContainsKey(p.TraceColor)) stats[p.TraceColor] = 1;
                else stats[p.TraceColor]++;
            }
            this.TraceStatsChanged?.Invoke(stats);
        }

        public bool OnPointerPressed(Point containerPoint, Point localPoint, PointerPoint pointer, bool isHoveringComponent)
        {
            this.PaletteStateChanged?.Invoke(false, default);
            double scale = this._parent.schematicsMatrix.M11;
            double hitTolerance = 8.0 / scale;

            if (this._hoverMarker != null)
                this._hoverMarker.IsVisible = false;

            if (pointer.Properties.IsRightButtonPressed)
            {
                // First check if a specific node marker was right-clicked
                if (this.GetHitMarker(localPoint, hitTolerance, out var polyline, out int nodeIndex))
                {
                    if (!this.GetColorVisibility(polyline!.TraceColor)) return false;

                    if (polyline.NodeCount <= 2)
                    {
                        this.PushUndo(polyline);
                    }

                    polyline!.RemoveNode(nodeIndex);
                    if (polyline.NodeCount < 2)
                    {
                        Logger.Info($"Trace deleted (clicked on marker): Color [{polyline.TraceColor}]");
                        this._polylines.Remove(polyline);
                        polyline.Dispose(this._canvas);
                        if (this._activePolyline == polyline) this._activePolyline = null;
                    }
                    this.NotifyStatsChanged();
                    this.TracesModified?.Invoke(); // Trigger save
                    return true;
                }

                // If no marker was hit, check if any line segment was right-clicked to delete the entire line
                if (this.GetHitSegment(localPoint, hitTolerance, out var polySegment, out int segmentIndex, out _))
                {
                    if (!this.GetColorVisibility(polySegment!.TraceColor)) return false;

                    this.PushUndo(polySegment!);
                    Logger.Info($"Trace deleted (clicked on line): Color [{polySegment!.TraceColor}], Markers [{polySegment.NodeCount}]");

                    this._polylines.Remove(polySegment);
                    polySegment.Dispose(this._canvas);

                    if (this._activePolyline == polySegment)
                        this._activePolyline = null;

                    this.NotifyStatsChanged();
                    this.TracesModified?.Invoke(); // Trigger save
                    return true;
                }

                return false;
            }

            if (pointer.Properties.IsLeftButtonPressed)
            {
                if (this.GetHitMarker(localPoint, hitTolerance, out var polyMarker, out int nIndex))
                {
                    if (!this.GetColorVisibility(polyMarker!.TraceColor)) return false;

                    this.SelectPolyline(polyMarker);
                    this._activePolyline = polyMarker;
                    this._draggingNodeIndex = nIndex;
                    this.CurrentDrawingColor = polyMarker.TraceColor;
                    return true;
                }

                if (this.GetHitSegment(localPoint, hitTolerance, out var polySegment, out int segmentIndex, out Point splitPoint))
                {
                    if (!this.GetColorVisibility(polySegment!.TraceColor)) return false;

                    this.SelectPolyline(polySegment);
                    polySegment!.InsertNode(segmentIndex + 1, this.CanvasToNormalized(splitPoint));
                    this._activePolyline = polySegment;
                    this._draggingNodeIndex = segmentIndex + 1;
                    this.CurrentDrawingColor = polySegment.TraceColor;
                    return true;
                }

                this.DeselectAll();

                if (!isHoveringComponent)
                {
                    this.SetVisibilityByColor(this.CurrentDrawingColor, true);
                    this._isDrawingNew = true;

                    var normPoint = this.CanvasToNormalized(localPoint);
                    this._drawingStartPoint = normPoint;

                    this._tempDrawingLine = new ManagedPolyline(this._drawingStartPoint, normPoint, this._parent, this.CurrentDrawingColor);
                    this._tempDrawingLine.AddToCanvas(this._canvas);
                    this.SelectPolyline(this._tempDrawingLine);
                    return true;
                }
            }

            return false;
        }

        public bool OnPointerMoved(Point localPoint, bool shiftDown)
        {
            double scale = this._parent.schematicsMatrix.M11;
            double snapTolerance = 15.0 / scale;

            if (this._hoverMarker != null)
                this._hoverMarker.IsVisible = false;

            if (this._isDrawingNew && this._tempDrawingLine != null)
            {
                Point pCanvas = this.NormalizedToCanvas(this.CanvasToNormalized(localPoint));
                if (shiftDown) pCanvas = this.ApplySnappingCanvas(pCanvas, this._tempDrawingLine, 1, snapTolerance);

                this._tempDrawingLine.MoveNode(1, this.CanvasToNormalized(pCanvas));
                return true;
            }

            if (this._activePolyline != null && this._draggingNodeIndex != -1)
            {
                Point pCanvas = this.NormalizedToCanvas(this.CanvasToNormalized(localPoint));
                if (shiftDown) pCanvas = this.ApplySnappingCanvas(pCanvas, this._activePolyline, this._draggingNodeIndex, snapTolerance);

                this._activePolyline.MoveNode(this._draggingNodeIndex, this.CanvasToNormalized(pCanvas));
                return true;
            }

            // Control visual hover selection when not drawing or dragging
            if (!this._isDrawingNew && this._draggingNodeIndex == -1)
            {
                double hitTolerance = 8.0 / scale;
                ManagedPolyline? hoveredPolyline = null;

                if (this.GetHitMarker(localPoint, hitTolerance, out var hoverMarkerLine, out _) && this.GetColorVisibility(hoverMarkerLine!.TraceColor))
                {
                    hoveredPolyline = hoverMarkerLine;
                }
                else if (this.GetHitSegment(localPoint, hitTolerance, out var hoverSegmentLine, out _, out Point hoverSnapPoint) && this.GetColorVisibility(hoverSegmentLine!.TraceColor))
                {
                    hoveredPolyline = hoverSegmentLine;

                    // SHOW phantom marker preview at segment projected line location
                    if (this._hoverMarker == null)
                    {
                        this._hoverMarker = new Ellipse
                        {
                            Fill = Brushes.White,
                            IsHitTestVisible = false,
                            UseLayoutRounding = false,
                            ZIndex = 100 // Guarantees the hover marker always renders on top of shapes
                        };
                        this._canvas.Children.Add(this._hoverMarker);
                    }

                    this._hoverMarker.Width = 14.0 / scale;
                    this._hoverMarker.Height = 14.0 / scale;
                    this._hoverMarker.StrokeThickness = 2.0 / scale;
                    this._hoverMarker.Stroke = new SolidColorBrush(hoverSegmentLine.TraceColor);
                    this._hoverMarker.IsVisible = true;

                    Canvas.SetLeft(this._hoverMarker, hoverSnapPoint.X - (this._hoverMarker.Width / 2.0));
                    Canvas.SetTop(this._hoverMarker, hoverSnapPoint.Y - (this._hoverMarker.Height / 2.0));
                }

                foreach (var poly in this._polylines)
                {
                    // Allow the hovered line to light up markers, but safely retain markers for lines that are actively being edited with the floating palette.
                    bool shouldBeVisuallySelected = (poly == hoveredPolyline) || (poly == this._activePolyline);
                    poly.SetSelected(shouldBeVisuallySelected);
                }
            }

            return false;
        }

        // ###########################################################################################
        // Maps an unconstrained layout-canvas coordinate into clamped 0.0 - 1.0 image-normalized boundaries.
        // ###########################################################################################
        private Point CanvasToNormalized(Point p)
        {
            var rect = this._parent.GetImageContentRect();
            if (rect.Width <= 0 || rect.Height <= 0) return new Point(0, 0);

            double nx = (p.X - rect.X) / rect.Width;
            double ny = (p.Y - rect.Y) / rect.Height;

            return new Point(Math.Max(0.0, Math.Min(1.0, nx)), Math.Max(0.0, Math.Min(1.0, ny)));
        }

        // ###########################################################################################
        // Maps a normalized coordinate boundary back exactly into visually accurate layout canvas boundaries.
        // ###########################################################################################
        internal Point NormalizedToCanvas(Point norm)
        {
            var rect = this._parent.GetImageContentRect();
            if (rect.Width <= 0 || rect.Height <= 0) return norm;

            return new Point(rect.X + (norm.X * rect.Width), rect.Y + (norm.Y * rect.Height));
        }

        private Point ApplySnappingCanvas(Point currentCanvas, ManagedPolyline poly, int nodeIndex, double tolerance)
        {
            double snapX = currentCanvas.X;
            double snapY = currentCanvas.Y;

            double closestXDist = tolerance;
            double closestYDist = tolerance;

            var neighbors = new List<Point>();
            if (nodeIndex > 0) neighbors.Add(this.NormalizedToCanvas(poly.GetNode(nodeIndex - 1)));
            if (nodeIndex < poly.NodeCount - 1) neighbors.Add(this.NormalizedToCanvas(poly.GetNode(nodeIndex + 1)));

            foreach (var n in neighbors)
            {
                if (Math.Abs(currentCanvas.X - n.X) < closestXDist) { snapX = n.X; closestXDist = Math.Abs(currentCanvas.X - n.X); }
                if (Math.Abs(currentCanvas.Y - n.Y) < closestYDist) { snapY = n.Y; closestYDist = Math.Abs(currentCanvas.Y - n.Y); }
            }

            return new Point(snapX, snapY);
        }

        public bool OnPointerReleased(Point containerPoint, Point localPoint)
        {
            if (this._isDrawingNew)
            {
                this._isDrawingNew = false;
                if (this._tempDrawingLine != null)
                {
                    double scale = this._parent.schematicsMatrix.M11;
                    Point pStart = this.NormalizedToCanvas(this._drawingStartPoint);
                    Point pEnd = this.NormalizedToCanvas(this._tempDrawingLine.GetNode(1));

                    if (Distance(pStart, pEnd) > (3.0 / scale))
                    {
                        this._polylines.Add(this._tempDrawingLine);
                        this._activePolyline = this._tempDrawingLine;

                        this.NotifyStatsChanged();
                        this.PaletteStateChanged?.Invoke(true, containerPoint);
                        this.TracesModified?.Invoke(); // Trigger save
                    }
                    else
                    {
                        this._tempDrawingLine.Dispose(this._canvas);
                        this.SelectPolyline(null);
                    }
                    this._tempDrawingLine = null;
                }
                return true;
            }

            if (this._activePolyline != null && this._draggingNodeIndex != -1)
            {
                this._draggingNodeIndex = -1;
                this.PaletteStateChanged?.Invoke(true, containerPoint);
                this.TracesModified?.Invoke(); // Trigger save
                return true;
            }

            return false;
        }

        public void UpdateScaleFactor(double zoomScale)
        {
            this._tempDrawingLine?.UpdateScale(zoomScale);
            foreach (var polyline in this._polylines) polyline.UpdateScale(zoomScale);
        }

        // ###########################################################################################
        // Exports internal engine structures directly into portable JSON friendly models.
        // ###########################################################################################
        public List<TraceModel> ExportTraces()
        {
            var result = new List<TraceModel>();
            foreach (var p in this._polylines)
            {
                result.Add(this.ExportSingleTrace(p));
            }
            return result;
        }

        private void AddTraceFromModel(TraceModel tm)
        {
            if (Color.TryParse(tm.Color, out Color c) && tm.Nodes.Count >= 2)
            {
                // Restore exact checkbox toggle state from disk
                this._colorVisibilityOptions[c] = tm.Visible;

                bool isLegacy = tm.Nodes.Any(n => n.X > 2.0 || n.Y > 2.0);

                Point ToNormalized(PointModel nm)
                {
                    if (!isLegacy) return new Point(nm.X, nm.Y);

                    var rect = this._parent.GetImageContentRect();
                    if (rect.Width <= 0 || rect.Height <= 0) return new Point(0, 0);

                    return new Point(
                        Math.Max(0.0, Math.Min(1.0, (nm.X - rect.X) / rect.Width)),
                        Math.Max(0.0, Math.Min(1.0, (nm.Y - rect.Y) / rect.Height))
                    );
                }

                var p1 = ToNormalized(tm.Nodes[0]);
                var p2 = ToNormalized(tm.Nodes[1]);
                var poly = new ManagedPolyline(p1, p2, this._parent, c);

                for (int i = 2; i < tm.Nodes.Count; i++)
                {
                    poly.InsertNode(i, ToNormalized(tm.Nodes[i]));
                }

                poly.UpdateScale(this._parent.schematicsMatrix.M11);

                poly.AddToCanvas(this._canvas);
                poly.SetGlobalVisibility(this.GetColorVisibility(c));
                this._polylines.Add(poly);
                this.AddOrReplacePaletteColor(c); // Force this active pin back into the standard dynamic HUD palette
            }
        }

        // ###########################################################################################
        // Ingests portable JSON friendly models and reconstructs active native elements automatically.
        // Also retroactively validates and updates non-compliant legacy native canvas coordinates natively saved mappings.
        // ###########################################################################################
        public void ImportTraces(List<TraceModel> traces)
        {
            this.ResetState(); // Prepare internal board context without firing wipe triggers

            foreach (var tm in traces)
            {
                this.AddTraceFromModel(tm);
            }
            this.NotifyStatsChanged();

            var stack = this.GetCurrentUndoStack();
            this.UndoStateChanged?.Invoke(stack.Count > 0);
        }

        public void ClearAllTracesAndSave()
        {
            // Wipe the internal global undo memory specifically when users trigger absolute clear
            var stack = this.GetCurrentUndoStack();
            stack.Clear();
            this.UndoStateChanged?.Invoke(false);

            this.ResetState();
            this.NotifyStatsChanged();
            this.TracesModified?.Invoke();
        }

        public void Reset() // Used for navigation switches; avoids overwriting save context
        {
            this.ResetState();
            this.NotifyStatsChanged();
        }

        private void ResetState()
        {
            this._tempDrawingLine?.Dispose(this._canvas);
            foreach (var p in this._polylines) p.Dispose(this._canvas);

            if (this._hoverMarker != null)
            {
                this._canvas.Children.Remove(this._hoverMarker);
                this._hoverMarker = null;
            }

            this._polylines.Clear();

            // Do not clear _globalUndoStacks here, as it retains internal context across board navigations natively
            this.UndoStateChanged?.Invoke(false);

            this._colorVisibilityOptions.Clear();
            this._activePolyline = null;
            this._isDrawingNew = false;
            this._draggingNodeIndex = -1;
            this._tempDrawingLine = null;

            this.PaletteColors = this.GetDefaultPaletteColors();
            this.CurrentDrawingColor = this.PaletteColors.FirstOrDefault();
            this.PaletteColorsChanged?.Invoke(this.PaletteColors);
    }

        private void DeselectAll()
        {
            foreach (var poly in this._polylines) poly.SetSelected(false);
            this._activePolyline = null;
        }

        private void SelectPolyline(ManagedPolyline? poly)
        {
            this.DeselectAll();
            if (poly != null) poly.SetSelected(true);
        }

        private bool GetHitMarker(Point localPoint, double tolerance, out ManagedPolyline? hitPolyline, out int nodeIndex)
        {
            hitPolyline = null;
            nodeIndex = -1;
            double closestLimit = tolerance * tolerance;

            foreach (var poly in this._polylines.Concat(this._tempDrawingLine != null ? new[] { this._tempDrawingLine } : Array.Empty<ManagedPolyline>()))
            {
                for (int i = 0; i < poly.NodeCount; i++)
                {
                    double distSq = DistanceSquared(this.NormalizedToCanvas(poly.GetNode(i)), localPoint);
                    if (distSq <= closestLimit)
                    {
                        closestLimit = distSq;
                        hitPolyline = poly;
                        nodeIndex = i;
                    }
                }
            }
            return hitPolyline != null;
        }

        private bool GetHitSegment(Point localPoint, double tolerance, out ManagedPolyline? hitPolyline, out int segmentIndex, out Point splitPoint)
        {
            hitPolyline = null;
            segmentIndex = -1;
            splitPoint = localPoint;
            double closestLimit = tolerance;

            foreach (var poly in this._polylines)
            {
                for (int i = 0; i < poly.NodeCount - 1; i++)
                {
                    Point a = this.NormalizedToCanvas(poly.GetNode(i));
                    Point b = this.NormalizedToCanvas(poly.GetNode(i + 1));
                    double distToLine = DistancePointToSegment(localPoint, a, b, out Point proj);

                    if (distToLine <= closestLimit)
                    {
                        closestLimit = distToLine;
                        hitPolyline = poly;
                        segmentIndex = i;
                        splitPoint = proj; // Snap exactly to vector geometry limit
                    }
                }
            }
            return hitPolyline != null;
        }

        private static double Distance(Point a, Point b) => Math.Sqrt(DistanceSquared(a, b));
        private static double DistanceSquared(Point a, Point b) => (a.X - b.X) * (a.X - b.X) + (a.Y - b.Y) * (a.Y - b.Y);
        private static double DistancePointToSegment(Point p, Point v, Point w, out Point projection)
        {
            double l2 = DistanceSquared(v, w);
            if (l2 == 0.0)
            {
                projection = v;
                return Distance(p, v);
            }

            double t = Math.Max(0, Math.Min(1, ((p.X - v.X) * (w.X - v.X) + (p.Y - v.Y) * (w.Y - v.Y)) / l2));
            projection = new Point(v.X + t * (w.X - v.X), v.Y + t * (w.Y - v.Y));
            return Distance(p, projection);
        }
    }

    // ###########################################################################################
    // Models for portable serialization format natively converting points reliably for JSON.
    // ###########################################################################################
    public class TraceModel
    {
        public string Color { get; set; } = string.Empty;
        public bool Visible { get; set; } = true;
        public List<PointModel> Nodes { get; set; } = new();
    }

    public class PointModel
    {
        public double X { get; set; }
        public double Y { get; set; }
    }

    // ###########################################################################################
    // Global I/O static handler exclusively saving schematic trace layouts accurately to disk.
    // ###########################################################################################
    public static class TraceStorage
    {
        private static readonly string filePath;
        private static Dictionary<string, Dictionary<string, List<TraceModel>>> data = new();

        static TraceStorage()
        {
            // Matches the path logic found in App.axaml.cs / UserSettings.cs
            var appData = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData);
            var directory = System.IO.Path.Combine(appData, AppConfig.AppFolderName);
            System.IO.Directory.CreateDirectory(directory);
            filePath = System.IO.Path.Combine(directory, AppConfig.TracesFileName);

            LoadFromFile();
        }

        public static void LoadFromFile()
        {
            if (File.Exists(filePath))
            {
                try
                {
                    var json = File.ReadAllText(filePath);
                    data = JsonSerializer.Deserialize<Dictionary<string, Dictionary<string, List<TraceModel>>>>(json) ?? new();
                }
                catch { data = new(); }
            }
        }

        public static void SaveToFile()
        {
            try
            {
                var json = JsonSerializer.Serialize(data, new JsonSerializerOptions { WriteIndented = true });
                File.WriteAllText(filePath, json);
            }
            catch { }
        }

        public static List<TraceModel> GetTraces(string boardKey, string schematicName)
        {
            if (data.TryGetValue(boardKey, out var boardData) && boardData.TryGetValue(schematicName, out var traces))
                return traces;
            return new List<TraceModel>();
        }

        public static void SaveTraces(string boardKey, string schematicName, List<TraceModel> traces)
        {
            if (!data.TryGetValue(boardKey, out var boardData))
            {
                boardData = new Dictionary<string, List<TraceModel>>(StringComparer.OrdinalIgnoreCase);
                data[boardKey] = boardData;
            }
            boardData[schematicName] = traces;
            SaveToFile();
        }
    }

    // ###########################################################################################
    // Represents a single polyline geometry composed of nodes, trace links, and selection objects.
    // ###########################################################################################
    internal class ManagedPolyline
    {
        private readonly List<Point> _nodes = new();
        private readonly Polyline _shape;
        private readonly List<Ellipse> _markers = new();
        private readonly CRT.TabSchematics _parent;

        public Color TraceColor { get; private set; }
        private bool _globalVisible = true;
        private bool _selected = false;

        public int NodeCount => this._nodes.Count;

        public ManagedPolyline(Point p1Norm, Point p2Norm, CRT.TabSchematics parent, Color traceColor)
        {
            this._parent = parent;
            this.TraceColor = traceColor;

            this._nodes.Add(p1Norm);
            this._nodes.Add(p2Norm);

            this._shape = new Polyline
            {
                Stroke = new SolidColorBrush(this.TraceColor),
                StrokeJoin = PenLineJoin.Round,
                StrokeLineCap = PenLineCap.Round,
                StrokeThickness = 3.0 / this.GetCurrentScale(),
                UseLayoutRounding = false
            };

            this.SyncShape();
        }

        public Point GetNode(int index) => this._nodes[index];

        public void SetColor(Color traceColor)
        {
            this.TraceColor = traceColor;
            var targetBrush = new SolidColorBrush(traceColor);

            this._shape.Stroke = targetBrush;
            foreach (var m in this._markers) m.Stroke = targetBrush;
        }

        public void MoveNode(int index, Point pNorm) { this._nodes[index] = pNorm; this.SyncShape(); }
        public void InsertNode(int index, Point pNorm) { this._nodes.Insert(index, pNorm); this.SyncShape(); }
        public void RemoveNode(int index) { this._nodes.RemoveAt(index); this.SyncShape(); }

        public void AddToCanvas(Canvas canvas)
        {
            if (!canvas.Children.Contains(this._shape))
            {
                canvas.Children.Add(this._shape);
                foreach (var m in this._markers) canvas.Children.Add(m);
            }
        }

        public void Dispose(Canvas canvas)
        {
            canvas.Children.Remove(this._shape);
            foreach (var m in this._markers) canvas.Children.Remove(m);
        }

        public void SetGlobalVisibility(bool visible) { this._globalVisible = visible; this.SyncVisibility(); }
        public void SetSelected(bool selected) { this._selected = selected; this.SyncVisibility(); }

        private void SyncVisibility()
        {
            this._shape.IsVisible = this._globalVisible;
            bool markersVisible = this._globalVisible && this._selected;
            foreach (var m in this._markers) m.IsVisible = markersVisible;
        }

        public void UpdateScale(double scale)
        {
            this._shape.StrokeThickness = 3.0 / scale;
            foreach (var m in this._markers)
            {
                m.Width = 14.0 / scale;
                m.Height = 14.0 / scale;
                m.StrokeThickness = 2.0 / scale;
            }
            this.SyncShape();
        }

        private double GetCurrentScale()
        {
            double rawScale = this._parent.schematicsMatrix.M11;
            return rawScale > 0 ? rawScale : 1.0;
        }

        // ###########################################################################################
        // Maps native layout boundaries locally scaling standard normalized bounds strictly to runtime configurations internally
        // ###########################################################################################
        private Point NormalizedToCanvas(Point norm)
        {
            var rect = this._parent.GetImageContentRect();
            if (rect.Width <= 0 || rect.Height <= 0) return norm;

            return new Point(rect.X + (norm.X * rect.Width), rect.Y + (norm.Y * rect.Height));
        }

        private void SyncShape()
        {
            var canvasNodes = new List<Point>(this._nodes.Count);
            foreach (var n in this._nodes)
            {
                canvasNodes.Add(this.NormalizedToCanvas(n));
            }

            this._shape.Points = new Avalonia.Collections.AvaloniaList<Point>(canvasNodes);

            if (this._markers.Count != this._nodes.Count)
            {
                var panel = this._shape.Parent as Canvas;
                foreach (var marker in this._markers) panel?.Children.Remove(marker);
                this._markers.Clear();

                double scale = this.GetCurrentScale();
                bool markersVisible = this._globalVisible && this._selected;
                var colorBrush = new SolidColorBrush(this.TraceColor);

                for (int i = 0; i < this._nodes.Count; i++)
                {
                    var marker = new Ellipse
                    {
                        Width = 14.0 / scale,
                        Height = 14.0 / scale,
                        Fill = Brushes.White,
                        Stroke = colorBrush,
                        StrokeThickness = 2.0 / scale,
                        IsVisible = markersVisible,
                        IsHitTestVisible = false,
                        UseLayoutRounding = false
                    };

                    this._markers.Add(marker);
                    panel?.Children.Add(marker);
                }
            }

            for (int i = 0; i < this._nodes.Count; i++)
            {
                Point p = canvasNodes[i];
                var m = this._markers[i];

                Canvas.SetLeft(m, p.X - (m.Width / 2.0));
                Canvas.SetTop(m, p.Y - (m.Height / 2.0));
            }
        }

    }
}