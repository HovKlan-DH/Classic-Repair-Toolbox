using System;
using System.Collections.Generic;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // ASKS FOR THE REASON BEFORE A SYSTEMS-SCREEN CHANGE IS PUBLISHED TO BETA (owner request,
    // 2026-10-03: "do ask for a change reason when clicking the 'Save changes' button, so this can go
    // along with the change, just like any normal contribution").
    //
    // Opened by SystemView.SendTableAsync after the server's check, with the files that check says
    // the publish would remove - named here, before anything is written, the way an approval shows
    // them. It decides nothing: it returns the reason, or null for Cancel.
    //
    // *** ENTER AND ESCAPE BOTH CANCEL, on the Tunnel route *** - the rule RollBackBetaWindow and
    // DeleteSystemWindow keep, for the same reason: a focused Button consumes Enter on the bubbling
    // route, and a reflexive keypress must not publish to BETA. Enter in the reason box is a new line.
    //
    // The reason is required: the button stays off until something is typed, and the server refuses a
    // blank one regardless (SystemEditFlow.NoReasonMessage).
    // ###########################################################################################
    public partial class PublishSystemChangeWindow : Window
    {
        public PublishSystemChangeWindow()
        {
            this.InitializeComponent();

            this.Title = SystemSections.ReasonTitle;
            this.SetText("ExplanationText", SystemSections.ReasonExplanation);
            this.SetText("ReasonLabelText", SystemSections.ReasonLabel);
            this.SetText("ReasonNoteText", SystemSections.ReasonNote);
            this.ReasonBox.PlaceholderText = SystemSections.ReasonHint;
            this.PublishControl.Content = SystemSections.PublishButton;

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // The reason typed, once confirmed; null while not (Cancel, Escape, Enter, the close box).
        public string? Reason { get; private set; }

        private TextBox ReasonBox => this.GetControl<TextBox>("ReasonTextBox");

        private Button PublishControl => this.GetControl<Button>("PublishButton");

        // `removals`: what the check said the publish would remove - none for nearly every change.
        // `reason`: a reason typed earlier, kept when the dialog is opened again.
        public void Initialize(string systemId, IReadOnlyList<string> removals, string? reason = null)
        {
            ArgumentNullException.ThrowIfNull(removals);

            this.SetText("HeadlineText", SystemSections.ReasonHeadline(systemId));

            StackPanel panel = this.GetControl<StackPanel>("RemovalsPanel");
            panel.Children.Clear();

            if (SystemSections.RemovalsHeading(removals) is string heading)
            {
                panel.Children.Add(new TextBlock
                {
                    Text = heading,
                    TextWrapping = Avalonia.Media.TextWrapping.Wrap,
                    FontWeight = Avalonia.Media.FontWeight.SemiBold
                });

                foreach (string path in removals)
                    panel.Children.Add(new TextBlock { Text = path, TextWrapping = Avalonia.Media.TextWrapping.Wrap });
            }

            panel.IsVisible = panel.Children.Count > 0;

            this.ReasonBox.Text = reason ?? string.Empty;
            this.ReasonBox.Focus();
            this.UpdatePublishButton();
        }

        private void SetText(string name, string text)
        {
            if (this.FindControl<TextBlock>(name) is TextBlock block)
                block.Text = text;
        }

        private void OnReasonChanged(object? sender, TextChangedEventArgs e) => this.UpdatePublishButton();

        private void UpdatePublishButton() =>
            this.PublishControl.IsEnabled = !string.IsNullOrWhiteSpace(this.ReasonBox.Text);

        private void OnPublishClick(object? sender, RoutedEventArgs e)
        {
            string reason = (this.ReasonBox.Text ?? string.Empty).Trim();

            if (reason.Length == 0)
                return;

            this.Reason = reason;
            this.Close(reason);
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close(null);

        // ###########################################################################################
        // Enter and Escape both CANCEL. Enter is NOT handled while the reason box has focus, or the
        // maintainer could not type a second line - the box accepts returns.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                this.Close(null);
                return;
            }

            if (e.Key != Key.Enter || this.ReasonBox.IsFocused)
                return;

            e.Handled = true;
            this.Close(null);
        }
    }
}
