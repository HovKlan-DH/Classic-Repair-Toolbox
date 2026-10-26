using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // IMPORTANT SIGNALS: SUGGESTIONS AND CLICK-TO-LIGHT-UP (CRT 4.0 step 2 - the plan: "Board >
    // Important signals: pick nets from a searchable KiCad list, with suggestions and
    // click-to-light-up").
    //
    //   - SUGGESTIONS: once the board's KiCad nets are read (UseNetChoices), every net whose name says
    //     what it is - a supply, ground, a clock, reset, a bus line - is offered as a row, named as the
    //     live boards name theirs (ImportantSignalSuggestions); a net a row has is never offered. "Add"
    //     takes one, "Add all suggestions" every one - ONE step either way.
    //   - CLICK-TO-LIGHT-UP: "Show on board" on a signal or a suggestion asks the host to show the
    //     Schematics tab with only that net lit (ShowNetRequested -> Main.ShowNetOnBoard), so a net can
    //     be seen before it is added.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        private readonly ObservableCollection<ComponentPaneSignalSuggestion> thisSignalSuggestions = new();

        // "Show on board" on a signal: the net to light up on the Schematics tab - the host's.
        internal event EventHandler<string>? ShowNetRequested;

        internal IReadOnlyList<ComponentPaneSignalSuggestion> SignalSuggestionsForTests => this.thisSignalSuggestions;

        private void InitializeSignalSuggestions()
        {
            this.SignalSuggestionsItemsControl.ItemsSource = this.thisSignalSuggestions;
            this.SignalSuggestionsIntroText.Text = ComponentPaneWording.SignalSuggestionsIntro;
        }

        // ###########################################################################################
        // The suggestions as the rows now read - after the nets are read and after every change of the
        // signals. None while the nets are being read, and on any other part.
        //
        // *** CHANGED IN PLACE (code review, 2026-10-26). *** Every box left commits the signals, and
        // a list cleared and refilled threw away the "Add" a press had just gone down on - the click
        // was lost (InPlaceList: a suggestion still suggested keeps its row, buttons and all).
        // ###########################################################################################
        private void ShowSignalSuggestions()
        {
            List<ComponentPaneSignalSuggestion> suggested =
                (this.thisBoardPart == PaneBoardPart.ImportantSignals && this.thisNetChoices is { } choices
                    ? ImportantSignalSuggestions.For(choices, this.thisSignals.Select(row => row.ToEntry()))
                    : [])
                .Select(row => new ComponentPaneSignalSuggestion(row.DisplayName, row.KiCadNetName))
                .ToList();

            InPlaceList.Sync(this.thisSignalSuggestions, suggested, (shown, wanted) => shown == wanted, wanted => wanted);

            this.SignalSuggestionsPanel.IsVisible = suggested.Count > 0;
            this.SignalSuggestionsHeading.Text = ComponentPaneWording.SignalSuggestionsHeading(suggested.Count);
        }

        private void OnAddSuggestionClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneSignalSuggestion suggestion })
            {
                this.AddSuggestions([suggestion]);
            }
        }

        private void OnAddAllSuggestionsClick(object? sender, RoutedEventArgs e) => this.AddSuggestions(null);

        // ###########################################################################################
        // `suggestions` - null for every one - added to the signals, as ONE step. What the boxes hold
        // goes in FIRST, and only what is still suggested then is added (code review, 2026-10-26: the
        // list taken before, a net just typed into a row went in twice). False when nothing was added.
        // ###########################################################################################
        internal bool AddSuggestions(IReadOnlyList<ComponentPaneSignalSuggestion>? suggestions)
        {
            if (this.thisSession is null || this.thisBoardPart != PaneBoardPart.ImportantSignals)
            {
                return false;
            }

            this.CommitHeld();

            List<ComponentPaneSignalSuggestion> adding = this.thisSignalSuggestions
                .Where(suggestion => suggestions is null || suggestions.Contains(suggestion))
                .ToList();

            if (adding.Count == 0)
            {
                return false;
            }

            foreach (ComponentPaneSignalSuggestion suggestion in adding)
            {
                var row = new ComponentPaneSignalRow { DisplayName = suggestion.DisplayName, KiCadNetName = suggestion.KiCadNetName };
                row.UseChoices(this.thisNetChoices);
                this.thisSignals.Add(row);
            }

            this.CommitSignals();
            this.ShowBoardPartLines();
            return true;
        }

        private void OnShowSignalOnBoardClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneSignalRow row })
            {
                this.ShowNetOnBoard(row.KiCadNetName);
            }
        }

        private void OnShowSuggestionOnBoardClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneSignalSuggestion suggestion })
            {
                this.ShowNetOnBoard(suggestion.KiCadNetName);
            }
        }

        // The net lit up on the Schematics tab - what the boxes hold going in first.
        internal void ShowNetOnBoard(string? net)
        {
            string wanted = net?.Trim() ?? string.Empty;

            if (wanted.Length == 0)
            {
                return;
            }

            this.CommitHeld();
            this.ShowNetRequested?.Invoke(this, wanted);
        }
    }

    // One signal suggested from the board's KiCad nets - its name, and the net as the picker spells it.
    public sealed record ComponentPaneSignalSuggestion(string DisplayName, string KiCadNetName);
}
