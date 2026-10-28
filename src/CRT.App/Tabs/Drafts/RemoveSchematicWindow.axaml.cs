using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Handlers.DataHandling;

namespace CRT
{
    // ###########################################################################################
    // Confirmation modal for "Remove" on a picture's card in the Contribute tab's Pictures part (the
    // "Schematic images" window's until CRT 4.0 step 3, ComponentPane.BoardPictureList.cs) - asked
    // only when the image has highlights drawn on it or a KiCad calibration, which go with it
    // (SchematicRemoval has why; code review, 2026-10-10). An image with nothing on it is removed at
    // once, as it always was. The Drafts tab's table asked it too, before a save that removed such
    // an image (code review, 2026-10-11), until the table saved itself (CRT 4.0 step 1, 2026-10-25):
    // it now says what went with the image once it is gone, with Ctrl+Z to bring it back
    // (BoardTableSessionWording.WentWith).
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
