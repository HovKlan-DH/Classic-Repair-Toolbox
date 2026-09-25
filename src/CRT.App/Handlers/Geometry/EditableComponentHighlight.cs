namespace Handlers.Geometry
{
    // ###########################################################################################
    // One component highlight rectangle as the label editor works on it: the in-memory row the
    // editor mutates while dragging, resizing and drawing, before anything is written back to the
    // board's Excel/JSON storage.
    //
    // Coordinates are BITMAP PIXELS of the schematic image the row belongs to, not screen or
    // control coordinates - the editor converts at its own boundary. That is what lets the
    // snapping maths in LabelEditorSnapGeometry be pure: it compares rows against each other in
    // one fixed space and never needs to know the zoom level or the control's size.
    //
    // This was a private nested class inside TabSchematics until the snapping logic was extracted.
    // It is a plain mutable data holder with no behaviour, deliberately: the editor edits these
    // rows in place and relies on reference identity to tell one row from another (see the
    // ReferenceEquals checks throughout LabelEditorSnapGeometry), so it must NOT become a record
    // or acquire value equality.
    // ###########################################################################################
    internal sealed class EditableComponentHighlight
    {
        public string SchematicName { get; set; } = string.Empty;
        public string BoardLabel { get; set; } = string.Empty;
        public string Category { get; set; } = string.Empty;
        public double X { get; set; }
        public double Y { get; set; }
        public double Width { get; set; }
        public double Height { get; set; }
    }
}
