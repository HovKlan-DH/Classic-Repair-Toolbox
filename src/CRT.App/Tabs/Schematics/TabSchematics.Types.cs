using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Input.GestureRecognizers;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tabs.TabSchematics;
using Handlers.Geometry;

namespace CRT;

// ###########################################################################################
// Private data types used across the TabSchematics parts: KiCad render/hit-test cache
// records, hover candidates, label-editor undo state and the editable highlight model.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private sealed class KiCadSchematicHoverLabelCandidate
    {
        public string NormalizedNetName { get; init; } = string.Empty;
        public Point LocalPoint { get; init; }
    }

    private sealed class KiCadSchematicHoverSegmentCandidate
    {
        public string NormalizedNetName { get; init; } = string.Empty;
        public Point StartLocal { get; init; }
        public Point EndLocal { get; init; }
    }

    private sealed class KiCadSchematicHoverHitTestCache
    {
        public double CellSizeLocal { get; init; } = 24.0;
        public List<KiCadSchematicHoverLabelCandidate> LabelCandidates { get; init; } = new();
        public List<KiCadSchematicHoverSegmentCandidate> SegmentCandidates { get; init; } = new();
        public Dictionary<long, List<int>> LabelIndicesByCell { get; init; } = new();
        public Dictionary<long, List<int>> SegmentIndicesByCell { get; init; } = new();
    }

    private sealed class ImportantSignalListItem
    {
        public string DisplayName { get; init; } = string.Empty;
        public string ToolTipText { get; init; } = string.Empty;

        public override string ToString()
        {
            return this.DisplayName;
        }
    }

    private sealed class KiCadRuntimeCacheScope
    {
        public Dictionary<string, KiCadPcbNetRenderCache> NetRenderCacheByKey { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        public Dictionary<string, KiCadPcbHoverHitTestCache> HoverHitTestCacheByKey { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        public Dictionary<string, Task> NetRenderBuildTaskByKey { get; } =
            new(StringComparer.OrdinalIgnoreCase);

        public Dictionary<string, Task> HoverHitTestBuildTaskByKey { get; } =
            new(StringComparer.OrdinalIgnoreCase);
    }

    private sealed class LabelEditorUndoHighlightState
    {
        public string SchematicName { get; set; } = string.Empty;
        public string BoardLabel { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
        public bool IsSelected { get; set; }
    }

    private sealed class LabelEditorUndoState
    {
        public List<LabelEditorUndoHighlightState> Highlights { get; } = new();
        public int PrimarySelectedIndex { get; set; } = -1;
    }

    private sealed class KiCadViewCalibration
    {
        public static KiCadViewCalibration Identity { get; } = new();

        public double ScaleX { get; init; } = 1.0;
        public double ScaleY { get; init; } = 1.0;
        public double OffsetX { get; init; }
        public double OffsetY { get; init; }
        public bool MirrorX { get; init; }
        public bool MirrorY { get; init; }
    }
}
