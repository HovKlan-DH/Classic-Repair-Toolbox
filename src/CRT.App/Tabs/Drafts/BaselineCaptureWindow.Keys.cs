using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.Media.Imaging;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.Geometry;
using Handlers.Oscilloscope;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // BaselineCaptureWindow's KEYS (see BaselineCaptureWindow.axaml.cs for the file map): the component info window's keys, the
    // numpad always the oscilloscope's here, and the number keys going to a pin.
    // ###########################################################################################
    public partial class BaselineCaptureWindow
    {
        // ---- the keys ---------------------------------------------------------------------------

        // A command of the model's being sent - one at a time.
        private bool thisCommandInFlight;

        // The number being typed (the number keys), cleared once typing pauses - the key it was last typed with, so an earlier
        // pause clears nothing.
        private string thisPinBuffer = string.Empty;
        private int thisPinTyped;

        // ###########################################################################################
        // THE COMPONENT INFO WINDOW'S KEYS, the numpad always the oscilloscope's here (owner request, 2026-11-20), Right the button that
        // was "Next pin". In a text box only Enter (it captures) and Escape are the window's; the rest are the box's. Escape while
        // closing asks about captures not yet accepted keeps capturing.
        //
        // *** NONE WHILE THE WINDOW WAITS (code review, 2026-11-21). *** A capture first waits for the oscilloscope to be set for its
        // pin, and reads the old picture, before any "please wait" is up: Left then set the oscilloscope for another pin, and the capture
        // was stored with that pin's settings; between the captures "Accept and close" puts in, Right or Enter started another save.
        // And with the close question showing, only it is answered; Space on a line's checkbox checks it.
        // ###########################################################################################
        private void OnWindowKeyDown(object? sender, KeyEventArgs e)
        {
            if (this.thisWorking || this.thisPutting)
            {
                e.Handled = true;
                return;
            }

            if ((e.KeyModifiers & (KeyModifiers.Control | KeyModifiers.Alt | KeyModifiers.Meta)) != 0)
            {
                return;
            }

            // The close question showing: it is answered first - ITS OWN buttons take Enter and Space (never "Capture" or "Accept and
            // continue to next pin", where a click left the keyboard), Escape keeps capturing, Tab goes between them, nothing else does
            // anything (code reviews, 2026-11-21: Enter took another capture, Right accepted and went on under it, the question keeping
            // its count).
            if (this.NotAcceptedPanel.IsVisible)
            {
                if (e.Key == Key.Escape)
                {
                    this.NotAcceptedPanel.IsVisible = false;
                    e.Handled = true;
                }
                else if (!(e.Key is Key.Enter or Key.Space && e.Source is Button button && this.NotAcceptedPanel.IsVisualAncestorOf(button)) && e.Key != Key.Tab)
                {
                    e.Handled = true;
                }

                return;
            }

            // Space on a line's checkbox checks it. Enter stays the capture from wherever the keyboard is - the footswitch's key, and a
            // click leaves the keyboard on the button clicked - so it never presses a focused button here.
            if (e.Key == Key.Space && e.Source is ToggleButton)
            {
                return;
            }

            bool inBox = e.Source is TextBox;

            switch (e.Key)
            {
                case Key.Escape:
                    this.Close();
                    break;

                case Key.Enter:
                    _ = this.CaptureFromKeyAsync();
                    break;

                case Key.Left when !inBox:
                    this.GoBy(-1);
                    break;

                case Key.Right when !inBox:
                    _ = this.AcceptFromKeyAsync();
                    break;

                case Key.Space when !inBox:
                    this.GoTo(this.thisTargets.ToList().FindIndex(target => target.Step.IsPicked), setScope: true);
                    break;

                case >= Key.D0 and <= Key.D9 when !inBox:
                    _ = this.TypePinDigitAsync((char)('0' + (e.Key - Key.D0)));
                    break;

                case Key.Up when !inBox:
                    this.WithScope(scope => scope.StepTriggerLevel(1));
                    break;

                case Key.Down when !inBox:
                    this.WithScope(scope => scope.StepTriggerLevel(-1));
                    break;

                case Key.Add when !inBox:
                    this.WithScope(scope => scope.StepTimeDiv(-1));
                    break;

                case Key.Subtract when !inBox:
                    this.WithScope(scope => scope.StepTimeDiv(1));
                    break;

                case Key.NumPad1 when !inBox:
                    this.WithScope(scope => scope.SetVoltsDiv(1.0));
                    break;

                case Key.NumPad2 when !inBox:
                    this.WithScope(scope => scope.SetVoltsDiv(2.0));
                    break;

                case Key.Decimal when !inBox:
                    _ = this.RunScopeCommandAsync(ScopeCommandPalette.ClearStatistics);
                    break;

                case Key.Multiply when !inBox:
                    _ = this.RunScopeCommandAsync(ScopeCommandPalette.Single);
                    break;

                case Key.Divide when !inBox:
                    _ = this.RunScopeCommandAsync(ScopeCommandPalette.Run);
                    break;

                default:
                    return;
            }

            e.Handled = true;
        }

        private async Task CaptureFromKeyAsync()
        {
            try
            {
                await this.CaptureCurrentAsync();
            }
            catch (Exception ex)
            {
                Logger.Critical($"Capturing a pin's oscilloscope baseline failed - [{ex}]");
            }
        }

        private void WithScope(Action<IBaselineScope> act)
        {
            if (this.thisConnected && this.thisPane?.Scope is { } scope)
            {
                act(scope);
            }
        }

        // A command of the model's run - one at a time, as the popup sends them.
        private async Task RunScopeCommandAsync(ScopeCommandPalette command)
        {
            if (this.thisCommandInFlight || !this.thisConnected || this.thisPane?.Scope is not { } scope)
            {
                return;
            }

            this.thisCommandInFlight = true;

            try
            {
                await scope.RunCommandAsync(command, CancellationToken.None);
            }
            catch (Exception ex)
            {
                Logger.Warning($"Sending [{command}] to the oscilloscope failed - [{ex.Message}]");
            }
            finally
            {
                this.thisCommandInFlight = false;
            }
        }

        // ###########################################################################################
        // A number key: the digits typed so far name a pin, shown at once, as the popup's go to a pin's first image - cleared once
        // typing pauses (600 ms, the popup's).
        // ###########################################################################################
        private async Task TypePinDigitAsync(char digit)
        {
            try
            {
                this.thisPinBuffer += digit;
                int typed = ++this.thisPinTyped;

                bool found = this.GoToPin(this.thisPinBuffer, pickIt: false);

                this.TypedPinText.Text = found ? $"Pin {this.thisPinBuffer}" : $"No pin {this.thisPinBuffer}";
                this.TypedPinBorder.IsVisible = true;

                await Task.Delay(600, CancellationToken.None);

                // Another digit since: the buffer goes on.
                if (typed == this.thisPinTyped)
                {
                    this.thisPinBuffer = string.Empty;
                    this.TypedPinBorder.IsVisible = false;
                }
            }
            catch (Exception ex)
            {
                Logger.Critical($"Going to a typed pin failed - [{ex}]");
            }
        }
    }
}
