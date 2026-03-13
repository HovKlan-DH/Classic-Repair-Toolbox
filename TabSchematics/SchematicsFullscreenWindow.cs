using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using System;

namespace CRT
{
    public sealed class SchematicsFullscreenWindow : Window
    {
        private readonly Control thisHostedContent;
        private readonly Action<Control> thisRestoreHostedContentAction;
        private bool thisHasRestoredHostedContent;

        // ###########################################################################################
        // Hosts the existing schematics control in a separate maximized window.
        // ###########################################################################################
        public SchematicsFullscreenWindow(Control hostedContent, Action<Control> restoreHostedContentAction)
        {
            this.thisHostedContent = hostedContent;
            this.thisRestoreHostedContentAction = restoreHostedContentAction;

            this.Title = "Classic Repair Toolbox - Schematics";
            this.MinWidth = 640;
            this.MinHeight = 400;
            this.Content = hostedContent;

            this.AddHandler(
                KeyDownEvent,
                this.OnWindowKeyDown,
                RoutingStrategies.Tunnel);

            this.Closing += this.OnWindowClosingRestoreHostedContent;
        }

        // ###########################################################################################
        // Closes the fullscreen schematics window when Escape is pressed.
        // ###########################################################################################
        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                this.Close();
                e.Handled = true;
            }
        }

        // ###########################################################################################
        // Restores the hosted schematics control before the fullscreen window finishes closing.
        // ###########################################################################################
        private void OnWindowClosingRestoreHostedContent(object? sender, WindowClosingEventArgs e)
        {
            this.RestoreHostedContent();
        }

        // ###########################################################################################
        // Moves the hosted schematics control back to the main window exactly once.
        // ###########################################################################################
        private void RestoreHostedContent()
        {
            if (this.thisHasRestoredHostedContent)
                return;

            this.thisHasRestoredHostedContent = true;

            this.Content = null;
            this.thisRestoreHostedContentAction(this.thisHostedContent);
        }
    }
}