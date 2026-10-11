using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System.Collections.Generic;

namespace CRT
{
    // ###########################################################################################
    // Confirmation modal for "Remove" on a schematic image in the "Schematic images" window - asked
    // only when the image has highlights drawn on it or a KiCad calibration, which go with it
    // (SchematicRemoval has why; code review, 2026-10-10). An image with nothing on it is removed at
    // once, as it always was. The Drafts tab's table asks it too, before a save that removes such
    // an image (InitializeForTableSave; code review, 2026-10-11).
    //
    // Its own window rather than a shared "confirm something" dialog, for the reason
    // DeleteWorklogWindow documents: each says exactly what it does, and changing one must not
    // change what another says.
    //
    // Returns true via ShowDialog when confirmed, or null when cancelled.
    // ###########################################################################################
    public partial class RemoveSchematicWindow : Window
    {
        public RemoveSchematicWindow()
        {
            this.InitializeComponent();

            // Tunnel, not a plain KeyDown - a focused Button handles Enter itself and marks the
            // event handled, so a bubbling handler never runs. See DiscardDraftWindow.
            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        public void Initialize(string schematicName, SchematicRemovalLoss loss)
        {
            this.QuestionText.Text = SchematicRemoval.Question(schematicName);
            this.WhatGoesText.Text = SchematicRemoval.WhatGoes(loss);
        }

        // ###########################################################################################
        // The same question from the Drafts tab's table, when a SAVE would remove images - rows
        // deleted, or renamed with their image file changed in the same save (code review,
        // 2026-10-11; BoardTableEditor.ConfirmRemovingSchematics). True saves and removes them.
        // ###########################################################################################
        public void InitializeForTableSave(IReadOnlyList<SchematicDrop> drops)
        {
            string heading = drops.Count == 1 ? "Remove schematic image" : "Remove schematic images";

            this.Title = heading;
            this.HeadingText.Text = heading;
            this.QuestionText.Text = SchematicRemoval.SaveQuestion(drops);
            this.WhatGoesText.Text = SchematicRemoval.SaveWhatGoes(drops);
            this.RemoveButton.Content = "Save and remove";
        }

        // Escape AND Enter both cancel, like every other confirmation of something that cannot be undone.
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape || e.Key == Key.Enter)
            {
                this.OnCancelClick(sender, e);
                e.Handled = true;
            }
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        private void OnRemoveClick(object? sender, RoutedEventArgs e) => this.Close(true);
    }
}
