using Avalonia;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Geometry
{
    // ###########################################################################################
    // WHERE A COMPONENT'S FOOTPRINT LIES ON A PICTURE ALIGNED WITH KICAD (CRT 4.0 step 3, slice 3c -
    // the plan: "Mark: on aligned PCB pictures, Mark all automatically draws each component's
    // rectangle from its footprint ... Auto-marking uses the existing calibration maths").
    //
    //   - THE BOARD'S WORLD BOUNDS AND THE MAPPING ONTO THE PICTURE are the Schematics tab's own - it
    //     draws the copper with these very two (PcbWorldBounds, WorldToImagePixel) - so a rectangle
    //     lands where the copper is drawn.
    //   - A FOOTPRINT'S RECTANGLE is its pads' extent, each pad turned by its own angle, with half a
    //     millimetre around and at least two across (an axial resistor's pads are a thin line). The
    //     body between a chip's rows is inside its pads' extent.
    //   - A FOOTPRINT IS SEEN FROM A SIDE when one of its pads is on that side's copper: a through-hole
    //     pad ("*.Cu") from both, a surface-mounted part from its own side only.
    //   - ONLY COMPONENTS THE BOARD HAS, NOT MARKED ON THIS PICTURE YET: a mark already drawn - by hand,
    //     perhaps moved to fit the photo - is never moved or doubled. A rectangle off the picture is
    //     none; one partly on it is cut to the picture, in whole pixels.
    //   - A MARK IS KiCad's WHILE IT LIES EXACTLY WHERE ITS FOOTPRINT PUTS IT (owner requests, 2026-11-01),
    //     under the alignment it was drawn from - told from the rectangle, so no board file changes.
    //     Aligned again, those move with the alignment (Realigned); under a card's thumbnail, those are
    //     the automatic ones (CountAutomatic).
    // ###########################################################################################
    // What "Mark all automatically" works with for one picture - the PCB it shows, from which side,
    // its alignment and its size - or why it cannot (Problem, said to the contributor). NotABoardSide:
    // the picture shows a schematic sheet - nothing to mark from KiCad, and the marking that follows
    // an alignment says nothing of it (2026-11-01).
    internal sealed record KiCadMarkPlan(string Picture, KiCadPcb? Pcb, bool Bottom, KiCadCalibrationValues Calibration, int Width, int Height, string? Problem, bool NotABoardSide = false)
    {
        // The marks for `board` as it is - worked out again on the draft as a change finds it.
        public List<ComponentHighlightEntry> Marks(BoardData board) =>
            this.Pcb is null ? [] : KiCadFootprintMarks.For(board, this.Picture, this.Pcb, this.Bottom, this.Calibration, this.Width, this.Height);

        // ###########################################################################################
        // Why `board` gets no marks, in the contributor's words (code review, 2026-10-26: each case read
        // "marked already"): no components yet, none with a footprint seen on the picture, or every
        // one marked already.
        // ###########################################################################################
        public string WhyNothing(BoardData board)
        {
            ArgumentNullException.ThrowIfNull(board);

            if (!board.Components.Any(row => !string.IsNullOrWhiteSpace(row.BoardLabel)))
            {
                return PictureToolWording.NoComponentsYet(this.Picture);
            }

            bool anyOnPicture = this.Pcb is not null &&
                                KiCadFootprintMarks.For(board, this.Picture, this.Pcb, this.Bottom, this.Calibration, this.Width, this.Height, skipMarked: false).Count > 0;

            return anyOnPicture
                ? PictureToolWording.MarkedAutomatically(0, this.Picture)
                : PictureToolWording.NoFootprintOnPicture(this.Picture);
        }
    }

    // Where one side of a board's footprints lie in the world - by label, capitals ignored - and that
    // PCB's world bounds: what tells a mark KiCad drew on a picture, with its alignment and its size.
    internal sealed record KiCadViewFootprints(Rect WorldBounds, IReadOnlyDictionary<string, Rect> ByLabel);

    internal static class KiCadFootprintMarks
    {
        // Around a footprint's pads, and the least a rectangle is across - in KiCad's millimetres.
        public const double MarginMm = 0.5;
        public const double SmallestMm = 2.0;

        // ###########################################################################################
        // What marking `picture` from KiCad works with: the picture as the draft has it, the KiCad view
        // its CAD name names - a side of the board, never a schematic sheet - its alignment in the
        // draft, and the image's size (`pixelSize`, from its header). The first thing missing is the
        // reason it cannot, in the contributor's words. `imageFile`, when given, picks the picture among
        // two of the same name (PictureRow) - its KiCad view is that row's.
        // ###########################################################################################
        public static KiCadMarkPlan Plan(BoardContent content, string picture, KiCadProjectRoot? root, (int Width, int Height)? pixelSize, string? imageFile = null)
        {
            ArgumentNullException.ThrowIfNull(content);

            string name = picture?.Trim() ?? string.Empty;
            KiCadMarkPlan Cannot(string problem) => new(name, null, false, default, 0, 0, problem);

            if (KiCadFootprintMarks.PictureRow(content.Board, name, imageFile) is not { } row)
            {
                return Cannot(PictureToolWording.PictureNotFound(name));
            }

            string cadName = row.CadName?.Trim() ?? string.Empty;

            if (cadName.Length == 0)
            {
                return Cannot(PictureToolWording.NoKiCadView(name));
            }

            if (content.Calibrations.FirstOrDefault(entry => string.Equals(entry.SchematicName?.Trim(), name, StringComparison.OrdinalIgnoreCase)) is not { } calibration)
            {
                return Cannot(PictureToolWording.NotAligned(name));
            }

            if (root is null)
            {
                return Cannot(PictureToolWording.KiCadNotRead);
            }

            if (root.Project.Views.FirstOrDefault(view => string.Equals(view.DisplayName?.Trim(), cadName, StringComparison.OrdinalIgnoreCase)) is not { } shown)
            {
                return Cannot(PictureToolWording.NoKiCadView(name));
            }

            KiCadViewSide side = BoardKiCadFacts.SideOf(shown);

            if (side == KiCadViewSide.Sheet)
            {
                return Cannot(PictureToolWording.NotABoardView(name)) with { NotABoardSide = true };
            }

            if (shown.SourceIndex < 0 || shown.SourceIndex >= root.Pcb.Count)
            {
                return Cannot(PictureToolWording.NoKiCadView(name));
            }

            if (pixelSize is not { Width: > 0, Height: > 0 } size)
            {
                return Cannot(PictureToolWording.ImageMissing(name));
            }

            return new KiCadMarkPlan(
                row.SchematicName!.Trim(),
                root.Pcb[shown.SourceIndex],
                side == KiCadViewSide.Bottom,
                KiCadFootprintMarks.ValuesOf(calibration),
                size.Width,
                size.Height,
                Problem: null);
        }

        public static KiCadCalibrationValues ValuesOf(KiCadCalibrationEntry calibration)
        {
            ArgumentNullException.ThrowIfNull(calibration);

            return new KiCadCalibrationValues(calibration.ScaleX, calibration.ScaleY, calibration.OffsetX, calibration.OffsetY, calibration.MirrorX, calibration.MirrorY);
        }

        // ###########################################################################################
        // The picture named `picture` - capitals ignored - and, of two sharing the name, the one whose
        // image is `imageFile` (a card on the Contribute tab is told apart so - SchematicRows.RefOf);
        // the first of the name when no file is given or none has it. Null for none of the name. Its
        // highlights and alignment are filed by NAME, so two of a name share those.
        // ###########################################################################################
        public static BoardSchematicEntry? PictureRow(BoardData board, string? picture, string? imageFile = null)
        {
            ArgumentNullException.ThrowIfNull(board);

            string name = picture?.Trim() ?? string.Empty;
            string file = imageFile?.Trim() ?? string.Empty;
            List<BoardSchematicEntry> named = board.Schematics
                .Where(row => string.Equals(row.SchematicName?.Trim(), name, StringComparison.OrdinalIgnoreCase))
                .ToList();

            return (file.Length == 0
                       ? null
                       : named.FirstOrDefault(row => string.Equals(row.SchematicImageFile?.Trim(), file, StringComparison.OrdinalIgnoreCase)))
                   ?? named.FirstOrDefault();
        }

        // ###########################################################################################
        // Whether any component is marked on `picture` (capitals ignored, as highlights are filed) - the
        // marking that follows an alignment leaves such a picture alone (2026-11-01): its marks were
        // drawn, by hand or from KiCad, and one more round is the contributor's to ask for.
        // ###########################################################################################
        public static bool HasMarks(BoardData board, string? picture)
        {
            ArgumentNullException.ThrowIfNull(board);

            string name = picture?.Trim() ?? string.Empty;
            return board.ComponentHighlights.Any(row => string.Equals(row.SchematicName?.Trim(), name, StringComparison.OrdinalIgnoreCase));
        }

        // How many components have a mark on `picture` - each once, capitals ignored.
        public static int MarkedCount(BoardData board, string? picture)
        {
            ArgumentNullException.ThrowIfNull(board);

            string name = picture?.Trim() ?? string.Empty;

            return board.ComponentHighlights
                .Where(row => string.Equals(row.SchematicName?.Trim(), name, StringComparison.OrdinalIgnoreCase))
                .Select(row => row.BoardLabel?.Trim() ?? string.Empty)
                .Distinct(StringComparer.OrdinalIgnoreCase)
                .Count();
        }

        // ###########################################################################################
        // Everything on the PCB the overlay draws - tracks, vias, arcs, pads, copper zones - as one
        // world rectangle: what a calibration maps onto the picture.
        // ###########################################################################################
        public static Rect PcbWorldBounds(KiCadPcb pcb)
        {
            ArgumentNullException.ThrowIfNull(pcb);

            bool hasValue = false;
            double minX = 0;
            double minY = 0;
            double maxX = 0;
            double maxY = 0;

            void Include(double x, double y)
            {
                if (!hasValue)
                {
                    minX = maxX = x;
                    minY = maxY = y;
                    hasValue = true;
                    return;
                }

                if (x < minX) minX = x;
                if (x > maxX) maxX = x;
                if (y < minY) minY = y;
                if (y > maxY) maxY = y;
            }

            foreach (var segment in pcb.Routing.Segments)
            {
                if (segment.Start != null) Include(segment.Start.X, segment.Start.Y);
                if (segment.End != null) Include(segment.End.X, segment.End.Y);
            }

            foreach (var via in pcb.Routing.Vias)
            {
                if (via.At != null) Include(via.At.X, via.At.Y);
            }

            foreach (var arc in pcb.Routing.Arcs)
            {
                if (arc.Start != null) Include(arc.Start.X, arc.Start.Y);
                if (arc.Mid != null) Include(arc.Mid.X, arc.Mid.Y);
                if (arc.End != null) Include(arc.End.X, arc.End.Y);
            }

            foreach (var footprint in pcb.Footprints)
            {
                foreach (var pad in footprint.Pads)
                {
                    if (pad.AbsoluteCenter == null)
                    {
                        continue;
                    }

                    double halfWidth = (pad.Size?.X ?? 0.0) / 2.0;
                    double halfHeight = (pad.Size?.Y ?? 0.0) / 2.0;

                    Include(pad.AbsoluteCenter.X - halfWidth, pad.AbsoluteCenter.Y - halfHeight);
                    Include(pad.AbsoluteCenter.X + halfWidth, pad.AbsoluteCenter.Y + halfHeight);
                }
            }

            foreach (var zone in pcb.Routing.Zones)
            {
                foreach (var polygon in zone.FilledPolygons.Count > 0 ? zone.FilledPolygons : zone.OutlinePolygons)
                {
                    foreach (var point in polygon.Points)
                    {
                        Include(point.X, point.Y);
                    }
                }
            }

            return hasValue
                ? new Rect(minX, minY, Math.Max(0.0001, maxX - minX), Math.Max(0.0001, maxY - minY))
                : default;
        }

        // ###########################################################################################
        // A KiCad world point on a picture `pixelWidth` x `pixelHeight` aligned by `calibration`: across
        // the world bounds 0..1, turned round by a mirror, scaled, then moved by the offset.
        // ###########################################################################################
        public static Point WorldToImagePixel(
            double worldX,
            double worldY,
            Rect worldBounds,
            KiCadCalibrationValues calibration,
            double pixelWidth,
            double pixelHeight)
        {
            if (worldBounds.Width <= 0 || worldBounds.Height <= 0)
            {
                return default;
            }

            double nx = (worldX - worldBounds.X) / worldBounds.Width;
            double ny = (worldY - worldBounds.Y) / worldBounds.Height;

            if (calibration.MirrorX)
            {
                nx = 1.0 - nx;
            }

            if (calibration.MirrorY)
            {
                ny = 1.0 - ny;
            }

            return new Point(
                calibration.OffsetX + (nx * calibration.ScaleX * pixelWidth),
                calibration.OffsetY + (ny * calibration.ScaleY * pixelHeight));
        }

        // ###########################################################################################
        // A KiCad world LENGTH (a trace's width) on a picture drawn `width` x `height`, aligned with
        // these scales: the mean of the two axes' scales, a mirror making no difference. The Schematics
        // tab's MapKiCadWorldLengthToLocal and a card's preview (KiCadPreviewTraces) both call it, so a
        // trace is as wide on the card as on the tab (code review, 2026-10-30: the card had a copy).
        // ###########################################################################################
        public static double WorldLengthOnPicture(double worldLength, Rect worldBounds, double scaleX, double scaleY, double width, double height)
        {
            double acrossX = (width / Math.Max(0.0001, worldBounds.Width)) * Math.Abs(scaleX);
            double acrossY = (height / Math.Max(0.0001, worldBounds.Height)) * Math.Abs(scaleY);

            return worldLength * ((acrossX + acrossY) / 2.0);
        }

        // ###########################################################################################
        // A footprint's extent in the world - its pads, each turned by its own angle, with MarginMm
        // around and at least SmallestMm across. Null for a footprint with no pad placed.
        // ###########################################################################################
        public static Rect? FootprintWorldRect(KiCadPcbFootprint footprint)
        {
            ArgumentNullException.ThrowIfNull(footprint);

            double minX = double.MaxValue;
            double minY = double.MaxValue;
            double maxX = double.MinValue;
            double maxY = double.MinValue;

            foreach (KiCadPcbPad pad in footprint.Pads.Where(pad => pad.AbsoluteCenter is not null))
            {
                double halfWidth = (pad.Size?.X ?? 0.0) / 2.0;
                double halfHeight = (pad.Size?.Y ?? 0.0) / 2.0;
                double angle = pad.RotationDegrees * Math.PI / 180.0;
                double cos = Math.Abs(Math.Cos(angle));
                double sin = Math.Abs(Math.Sin(angle));

                double extentX = (halfWidth * cos) + (halfHeight * sin);
                double extentY = (halfWidth * sin) + (halfHeight * cos);

                minX = Math.Min(minX, pad.AbsoluteCenter!.X - extentX);
                maxX = Math.Max(maxX, pad.AbsoluteCenter.X + extentX);
                minY = Math.Min(minY, pad.AbsoluteCenter.Y - extentY);
                maxY = Math.Max(maxY, pad.AbsoluteCenter.Y + extentY);
            }

            if (minX > maxX)
            {
                return null;
            }

            minX -= MarginMm;
            minY -= MarginMm;
            maxX += MarginMm;
            maxY += MarginMm;

            (minX, maxX) = KiCadFootprintMarks.AtLeast(minX, maxX, SmallestMm);
            (minY, maxY) = KiCadFootprintMarks.AtLeast(minY, maxY, SmallestMm);

            return new Rect(minX, minY, maxX - minX, maxY - minY);
        }

        // Whether a footprint is seen from the side whose copper is `copperLayer` ("F.Cu", "B.Cu").
        public static bool IsOnSide(KiCadPcbFootprint footprint, string copperLayer)
        {
            ArgumentNullException.ThrowIfNull(footprint);

            return footprint.Pads.Any(pad => pad.Layers.Any(layer =>
                string.Equals(layer, "*.Cu", StringComparison.OrdinalIgnoreCase) ||
                string.Equals(layer, copperLayer, StringComparison.OrdinalIgnoreCase)));
        }

        // ###########################################################################################
        // A world rectangle on the picture, in whole pixels, cut to the picture - null when nothing of
        // it is on the picture. A mirrored calibration turns the corners round; the rectangle is the
        // same.
        // ###########################################################################################
        public static PixelRect? OnPicture(Rect worldRect, Rect worldBounds, KiCadCalibrationValues calibration, int pixelWidth, int pixelHeight)
        {
            if (pixelWidth <= 0 || pixelHeight <= 0 || worldBounds.Width <= 0 || worldBounds.Height <= 0)
            {
                return null;
            }

            Point first = KiCadFootprintMarks.WorldToImagePixel(worldRect.Left, worldRect.Top, worldBounds, calibration, pixelWidth, pixelHeight);
            Point second = KiCadFootprintMarks.WorldToImagePixel(worldRect.Right, worldRect.Bottom, worldBounds, calibration, pixelWidth, pixelHeight);

            // A corner a hair off a whole pixel IS that pixel - the arithmetic's own noise (125.0000001)
            // must not add a pixel to a rectangle.
            double left = Math.Max(0, Math.Floor(KiCadFootprintMarks.Snapped(Math.Min(first.X, second.X))));
            double top = Math.Max(0, Math.Floor(KiCadFootprintMarks.Snapped(Math.Min(first.Y, second.Y))));
            double right = Math.Min(pixelWidth, Math.Ceiling(KiCadFootprintMarks.Snapped(Math.Max(first.X, second.X))));
            double bottom = Math.Min(pixelHeight, Math.Ceiling(KiCadFootprintMarks.Snapped(Math.Max(first.Y, second.Y))));

            return right - left >= 1 && bottom - top >= 1
                ? new PixelRect((int)left, (int)top, (int)(right - left), (int)(bottom - top))
                : null;
        }

        // ###########################################################################################
        // The marks KiCad draws on picture `schematicName`, showing `pcb` from the bottom or not: a
        // rectangle for every component of `board` with a footprint seen from that side and no mark on
        // this picture yet - under the board's own spelling of its label, in the board's order.
        // `skipMarked` false: the ones marked already too - what would be drawn on a bare picture.
        // ###########################################################################################
        public static List<ComponentHighlightEntry> For(
            BoardData board,
            string schematicName,
            KiCadPcb pcb,
            bool bottom,
            KiCadCalibrationValues calibration,
            int pixelWidth,
            int pixelHeight,
            bool skipMarked = true)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(pcb);

            string picture = schematicName?.Trim() ?? string.Empty;
            string copper = bottom ? "B.Cu" : "F.Cu";
            Rect worldBounds = KiCadFootprintMarks.PcbWorldBounds(pcb);

            var marked = new HashSet<string>(
                board.ComponentHighlights
                    .Where(row => string.Equals(row.SchematicName?.Trim(), picture, StringComparison.OrdinalIgnoreCase))
                    .Select(row => row.BoardLabel?.Trim() ?? string.Empty),
                StringComparer.OrdinalIgnoreCase);

            var footprints = new Dictionary<string, KiCadPcbFootprint>(StringComparer.OrdinalIgnoreCase);
            foreach (KiCadPcbFootprint footprint in pcb.Footprints.Where(footprint => (footprint.Reference?.Trim() ?? string.Empty).Length > 0))
            {
                footprints.TryAdd(footprint.Reference!.Trim(), footprint);
            }

            var marks = new List<ComponentHighlightEntry>();
            var done = new HashSet<string>(StringComparer.OrdinalIgnoreCase);

            foreach (string label in board.Components.Select(row => row.BoardLabel?.Trim() ?? string.Empty).Where(label => label.Length > 0))
            {
                if ((skipMarked && marked.Contains(label)) || !done.Add(label) ||
                    !footprints.TryGetValue(label, out KiCadPcbFootprint? footprint) ||
                    !KiCadFootprintMarks.IsOnSide(footprint, copper) ||
                    KiCadFootprintMarks.FootprintWorldRect(footprint) is not { } worldRect ||
                    KiCadFootprintMarks.OnPicture(worldRect, worldBounds, calibration, pixelWidth, pixelHeight) is not { } rect)
                {
                    continue;
                }

                marks.Add(new ComponentHighlightEntry
                {
                    SchematicName = picture,
                    BoardLabel = label,
                    X = rect.X.ToString(CultureInfo.InvariantCulture),
                    Y = rect.Y.ToString(CultureInfo.InvariantCulture),
                    Width = rect.Width.ToString(CultureInfo.InvariantCulture),
                    Height = rect.Height.ToString(CultureInfo.InvariantCulture),
                });
            }

            return marks;
        }

        // ###########################################################################################
        // KiCad ALIGNED AGAIN ON THE PICTURE `plan` IS OF, the alignment it had `before` (owner request,
        // 2026-11-01: "if I then again do the Align KiCad then it seems it will not readjust the components,
        // which I think it should"): every mark lying exactly where its footprint put it under `before`
        // goes where it puts it now - one now off the picture goes - and any component that had no place
        // on the picture then gets one. A mark anywhere else was moved or drawn by hand, and stays.
        // `changed`: how many marks moved, went or came - 0, and the same board, for none. A picture with
        // no mark of KiCad's is left as it is: its marks are all by hand.
        // ###########################################################################################
        public static BoardData Realigned(BoardData board, KiCadMarkPlan plan, KiCadCalibrationValues before, out int changed)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(plan);

            changed = 0;

            if (plan.Pcb is null)
            {
                return board;
            }

            Dictionary<string, ComponentHighlightEntry> was = KiCadFootprintMarks.ByLabel(
                KiCadFootprintMarks.For(board, plan.Picture, plan.Pcb, plan.Bottom, before, plan.Width, plan.Height, skipMarked: false));
            Dictionary<string, ComponentHighlightEntry> now = KiCadFootprintMarks.ByLabel(
                KiCadFootprintMarks.For(board, plan.Picture, plan.Pcb, plan.Bottom, plan.Calibration, plan.Width, plan.Height, skipMarked: false));

            var highlights = new List<ComponentHighlightEntry>(board.ComponentHighlights.Count);
            bool anyKiCads = false;

            foreach (ComponentHighlightEntry row in board.ComponentHighlights)
            {
                string label = row.BoardLabel?.Trim() ?? string.Empty;

                if (!string.Equals(row.SchematicName?.Trim(), plan.Picture, StringComparison.OrdinalIgnoreCase) ||
                    !was.TryGetValue(label, out ComponentHighlightEntry? drawn) ||
                    !KiCadFootprintMarks.SameRectangle(row, drawn))
                {
                    highlights.Add(row);
                    continue;
                }

                anyKiCads = true;

                if (!now.TryGetValue(label, out ComponentHighlightEntry? place))
                {
                    changed++;
                    continue;
                }

                if (!KiCadFootprintMarks.SameRectangle(row, place))
                {
                    changed++;
                }

                highlights.Add(row.WithRectangleOf(place));
            }

            if (!anyKiCads)
            {
                return board;
            }

            BoardData moved = board.WithComponentHighlights(highlights);

            // Only a footprint NEWLY on the picture (code review, 2026-11-01): one that lay on it all along
            // with no mark had that mark deleted on purpose, and aligning again must not bring it back.
            List<ComponentHighlightEntry> added = plan.Marks(moved).Where(mark => !was.ContainsKey(mark.BoardLabel)).ToList();

            changed += added.Count;

            return changed == 0 ? board : KiCadFootprintMarks.WithMarks(moved, added);
        }

        // ###########################################################################################
        // Where each side of the board's footprints lie, by the view's name - `For`'s rule: a reference's
        // first footprint, seen from that side. No entry for a schematic sheet.
        // ###########################################################################################
        public static Dictionary<string, KiCadViewFootprints> ViewFootprints(KiCadProjectRoot? root)
        {
            var views = new Dictionary<string, KiCadViewFootprints>(StringComparer.OrdinalIgnoreCase);

            foreach (KiCadProjectView view in root?.Project.Views ?? [])
            {
                string name = view.DisplayName?.Trim() ?? string.Empty;
                KiCadViewSide side = BoardKiCadFacts.SideOf(view);

                if (name.Length == 0 || side == KiCadViewSide.Sheet || view.SourceIndex < 0 || view.SourceIndex >= root!.Pcb.Count)
                {
                    continue;
                }

                KiCadPcb pcb = root.Pcb[view.SourceIndex];
                var seen = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                var byLabel = new Dictionary<string, Rect>(StringComparer.OrdinalIgnoreCase);

                foreach (KiCadPcbFootprint footprint in pcb.Footprints)
                {
                    string label = footprint.Reference?.Trim() ?? string.Empty;

                    if (label.Length > 0 && seen.Add(label) &&
                        KiCadFootprintMarks.IsOnSide(footprint, side == KiCadViewSide.Bottom ? "B.Cu" : "F.Cu") &&
                        KiCadFootprintMarks.FootprintWorldRect(footprint) is { } worldRect)
                    {
                        byLabel[label] = worldRect;
                    }
                }

                views[name] = new KiCadViewFootprints(KiCadFootprintMarks.PcbWorldBounds(pcb), byLabel);
            }

            return views;
        }

        // ###########################################################################################
        // How many of the labels in `marks` have a mark exactly where KiCad draws it on this picture.
        // ###########################################################################################
        public static int CountAutomatic(
            IReadOnlyDictionary<string, List<Rect>> marks,
            KiCadViewFootprints footprints,
            KiCadCalibrationValues alignment,
            int pixelWidth,
            int pixelHeight)
        {
            ArgumentNullException.ThrowIfNull(marks);
            ArgumentNullException.ThrowIfNull(footprints);

            return marks.Count(mark =>
                footprints.ByLabel.TryGetValue(mark.Key?.Trim() ?? string.Empty, out Rect worldRect) &&
                KiCadFootprintMarks.OnPicture(worldRect, footprints.WorldBounds, alignment, pixelWidth, pixelHeight) is { } place &&
                mark.Value.Any(rect =>
                    Math.Round(rect.X) == place.X && Math.Round(rect.Y) == place.Y &&
                    Math.Round(rect.Width) == place.Width && Math.Round(rect.Height) == place.Height));
        }

        private static Dictionary<string, ComponentHighlightEntry> ByLabel(IEnumerable<ComponentHighlightEntry> marks)
        {
            var byLabel = new Dictionary<string, ComponentHighlightEntry>(StringComparer.OrdinalIgnoreCase);

            foreach (ComponentHighlightEntry mark in marks)
            {
                byLabel.TryAdd(mark.BoardLabel, mark);
            }

            return byLabel;
        }

        // Whether two marks are the same rectangle - by value, so "12" and "12.0" are.
        private static bool SameRectangle(ComponentHighlightEntry first, ComponentHighlightEntry second) =>
            RectGeometry.TryParseDouble(first.X, out double x) && RectGeometry.TryParseDouble(second.X, out double otherX) && x == otherX &&
            RectGeometry.TryParseDouble(first.Y, out double y) && RectGeometry.TryParseDouble(second.Y, out double otherY) && y == otherY &&
            RectGeometry.TryParseDouble(first.Width, out double width) && RectGeometry.TryParseDouble(second.Width, out double otherWidth) && width == otherWidth &&
            RectGeometry.TryParseDouble(first.Height, out double height) && RectGeometry.TryParseDouble(second.Height, out double otherHeight) && height == otherHeight;

        // `board` with `marks` added - after the last mark of their picture, else at the end. The same
        // board when there are none.
        public static BoardData WithMarks(BoardData board, IReadOnlyList<ComponentHighlightEntry> marks)
        {
            ArgumentNullException.ThrowIfNull(board);
            ArgumentNullException.ThrowIfNull(marks);

            if (marks.Count == 0)
            {
                return board;
            }

            string picture = marks[0].SchematicName;
            List<ComponentHighlightEntry> highlights = board.ComponentHighlights.ToList();
            int last = highlights.FindLastIndex(row => string.Equals(row.SchematicName?.Trim(), picture, StringComparison.OrdinalIgnoreCase));

            highlights.InsertRange(last >= 0 ? last + 1 : highlights.Count, marks);

            return board.WithComponentHighlights(highlights);
        }

        private static double Snapped(double value) =>
            Math.Abs(value - Math.Round(value)) < 1e-6 ? Math.Round(value) : value;

        // `min`..`max` widened about its middle to at least `least`.
        private static (double Min, double Max) AtLeast(double min, double max, double least)
        {
            double missing = least - (max - min);

            return missing > 0 ? (min - (missing / 2.0), max + (missing / 2.0)) : (min, max);
        }
    }
}
