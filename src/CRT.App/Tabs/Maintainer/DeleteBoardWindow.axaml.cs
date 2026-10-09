using System;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Markup.Xaml;
using Avalonia.Media;
using Handlers.DataHandling;
using Handlers.MaintainerHandling;

namespace CRT
{
    // ###########################################################################################
    // CONFIRMING A BOARD'S DELETION (owner request, 2026-10-03) - Account > "Delete a board".
    //
    // *** THE POINT OF THE DIALOG IS THE LIST. *** A delete cannot be undone and reaches both data
    // trees and the database, so the administrator sees, before anything happens, what goes from
    // each - and which contributors' open work goes with it. Every line is BoardDeletionWording's,
    // over the server's own plan.
    //
    // *** ENTER AND ESCAPE BOTH CANCEL, on the Tunnel route *** - the rule and the reason of
    // RollBackBetaWindow and DeleteWorkbookWindow: a focused Button consumes Enter on the bubbling
    // route, so a reflexive keypress would delete a board. Enter in the reason box is a new line.
    //
    // The reason is required only when somebody is mailed (BoardDeletionWording.NeedsReason) - here
    // AND on the server, which refuses the delete without one.
    // ###########################################################################################
    public partial class DeleteBoardWindow : Window
    {
        private bool thisNeedsReason;

        public DeleteBoardWindow()
        {
            this.InitializeComponent();

            this.AddHandler(KeyDownEvent, this.OnWindowPreviewKeyDown, RoutingStrategies.Tunnel);
        }

        private void InitializeComponent() => AvaloniaXamlLoader.Load(this);

        // True when the administrator confirmed; Reason is then what they typed (empty when none
        // was asked for).
        public bool WasConfirmed { get; private set; }

        public string Reason { get; private set; } = string.Empty;

        public void Initialize(BoardDeletePlanAnswer plan)
        {
            ArgumentNullException.ThrowIfNull(plan);

            this.Title = BoardDeletionWording.ConfirmTitle;
            this.SetText("HeadlineText", BoardDeletionWording.Headline(plan));
            this.SetText("SharedFilesText", BoardDeletionWording.SharedFilesKept);
            this.SetText("CannotBeUndoneText", BoardDeletionWording.CannotBeUndone);

            if (this.FindControl<TextBlock>("CannotBeUndoneText") is TextBlock undone)
                undone.Foreground = Brushes.IndianRed;

            if (this.FindControl<StackPanel>("WhatGoesPanel") is StackPanel whatGoes)
            {
                whatGoes.Children.Clear();

                foreach (string line in BoardDeletionWording.WhatGoes(plan))
                    whatGoes.Children.Add(new TextBlock { Text = line, TextWrapping = TextWrapping.Wrap });
            }

            this.thisNeedsReason = BoardDeletionWording.NeedsReason(plan);

            if (this.FindControl<StackPanel>("OpenPanel") is StackPanel open)
                open.IsVisible = this.thisNeedsReason;

            this.SetText("OpenHeadingText", BoardDeletionWording.OpenHeading(plan) ?? string.Empty);
            this.SetText("ReasonPromptText", BoardDeletionWording.ReasonPrompt);

            if (this.FindControl<StackPanel>("OpenListPanel") is StackPanel list)
            {
                list.Children.Clear();

                foreach (BoardDeleteOpenSubmission submission in plan.OpenSubmissions)
                    list.Children.Add(new TextBlock { Text = BoardDeletionWording.OpenLine(submission), TextWrapping = TextWrapping.Wrap });
            }

            if (this.FindControl<Button>("ConfirmButton") is Button confirm)
                confirm.Content = BoardDeletionWording.ConfirmButton;

            this.UpdateConfirm();
        }

        private void SetText(string name, string text)
        {
            if (this.FindControl<TextBlock>(name) is TextBlock block)
                block.Text = text;
        }

        private string TypedReason => (this.FindControl<TextBox>("ReasonTextBox")?.Text ?? string.Empty).Trim();

        private void UpdateConfirm()
        {
            if (this.FindControl<Button>("ConfirmButton") is Button confirm)
                confirm.IsEnabled = !this.thisNeedsReason || this.TypedReason.Length > 0;
        }

        private void OnReasonChanged(object? sender, TextChangedEventArgs e) => this.UpdateConfirm();

        private void OnConfirmClick(object? sender, RoutedEventArgs e)
        {
            string reason = this.thisNeedsReason ? this.TypedReason : string.Empty;

            if (this.thisNeedsReason && reason.Length == 0)
                return;

            this.WasConfirmed = true;
            this.Reason = reason;
            this.Close();
        }

        private void OnCancelClick(object? sender, RoutedEventArgs e) => this.Close();

        // ###########################################################################################
        // Enter and Escape both CANCEL. Enter is NOT handled while the reason box has focus, so a
        // second line of explanation can be typed - the box accepts returns.
        // ###########################################################################################
        private void OnWindowPreviewKeyDown(object? sender, KeyEventArgs e)
        {
            if (e.Key == Key.Escape)
            {
                e.Handled = true;
                this.Close();
                return;
            }

            if (e.Key != Key.Enter)
                return;

            if (this.FindControl<TextBox>("ReasonTextBox")?.IsFocused == true)
                return;

            e.Handled = true;
            this.Close();
        }
    }
}
