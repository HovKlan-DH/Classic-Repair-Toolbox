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
    //     live boards name theirs (ImportantSignalSuggestions); a net a row has is never offered. The
    //     name is a box, changed before it is added (owner request, 2026-10-30: "show the suggested
    //     names in an input text field"). "Add" takes one, "Add all suggestions" every one - ONE step
    //     either way, each under the name its box holds.
    //   - CLICK-TO-LIGHT-UP: "Show on board" on a signal or a suggestion asks the host to show the
    //     Schematics tab with only that net lit (ShowNetRequested -> Main.ShowNetOnBoard), so a net can
    //     be seen before it is added.
    //   - EVERY OTHER NET, LAST (owner requests, 2026-11-01: "all other nets not already used ... below
    //     the Suggested ones, so all nets will be visible in here", and "a proposed net name in an input
    //     field"): the nets neither a row nor a suggestion has (ImportantSignalSuggestions.OtherNets), each
    //     with the name it would be added under (StartingNameOf) in a box, "Show on board" and "Add".
    //     Changed in place for the same reason as the suggestions - a name typed stays - and realised
    //     only as it is scrolled to.
    //
    // Part of the ComponentPane partial class - see ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        private readonly ObservableCollection<ComponentPaneSignalSuggestion> thisSignalSuggestions = new();

        // "Show on board" on a signal: the net to light up on the Schematics tab - the host's.
        internal event EventHandler<string>? ShowNetRequested;

        internal IReadOnlyList<ComponentPaneSignalSuggestion> SignalSuggestionsForTests => this.thisSignalSuggestions;

        private readonly ObservableCollection<ComponentPaneOtherNet> thisOtherNets = new();

        internal IReadOnlyList<ComponentPaneOtherNet> OtherNetsForTests => this.thisOtherNets;

        private void InitializeSignalSuggestions()
        {
            this.SignalSuggestionsItemsControl.ItemsSource = this.thisSignalSuggestions;
            this.SignalSuggestionsIntroText.Text = ComponentPaneWording.SignalSuggestionsIntro;
            this.OtherNetsItemsControl.ItemsSource = this.thisOtherNets;
            this.OtherNetsIntroText.Text = ComponentPaneWording.OtherNetsIntro;
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

            // A suggestion still suggested keeps its row - and the name typed into its box.
            InPlaceList.Sync(this.thisSignalSuggestions, suggested, (shown, wanted) => shown.IsTheSameSuggestion(wanted), wanted => wanted);

            this.SignalSuggestionsPanel.IsVisible = suggested.Count > 0;
            this.SignalSuggestionsHeading.Text = ComponentPaneWording.SignalSuggestionsHeading(suggested.Count);

            // Every net left - last, so all the board's nets are to be seen here (2026-11-01).
            List<string> others = this.thisBoardPart == PaneBoardPart.ImportantSignals && this.thisNetChoices is { } nets
                ? ImportantSignalSuggestions.OtherNets(nets, this.thisSignals.Select(row => row.ToEntry()))
                : [];

            InPlaceList.Sync(
                this.thisOtherNets,
                others,
                (shown, net) => string.Equals(shown.KiCadNetName, net, StringComparison.Ordinal),
                net => new ComponentPaneOtherNet(net));

            this.OtherNetsPanel.IsVisible = others.Count > 0;
            this.OtherNetsHeading.Text = ComponentPaneWording.OtherNetsHeading(others.Count);
        }

        private void OnAddOtherNetClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneOtherNet net })
            {
                this.AddSignalForNet(net.KiCadNetName, net.NameToAdd);
            }
        }

        private void OnShowOtherNetOnBoardClick(object? sender, RoutedEventArgs e)
        {
            if (sender is Control { Tag: ComponentPaneOtherNet net })
            {
                this.ShowNetOnBoard(net.KiCadNetName);
            }
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
                var row = new ComponentPaneSignalRow { DisplayName = suggestion.NameToAdd, KiCadNetName = suggestion.KiCadNetName };
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

    // ###########################################################################################
    // A KiCad net offered to be added as a signal - the net as the picker spells it, and the name it
    // would be added under: the suggested one, until another is typed into its box (owner request,
    // 2026-10-30). A box emptied adds it under the suggested name - a signal with no name lights up
    // nothing a person can pick out on the Schematics tab's panel.
    // ###########################################################################################
    public abstract class ComponentPaneNetToAdd : ComponentPaneRowBase
    {
        private string thisDisplayName;

        protected ComponentPaneNetToAdd(string? suggestedName, string? kiCadNetName)
        {
            this.SuggestedName = suggestedName ?? string.Empty;
            this.KiCadNetName = kiCadNetName ?? string.Empty;
            this.thisDisplayName = this.SuggestedName;
        }

        public string SuggestedName { get; }

        public string KiCadNetName { get; }

        // The name box.
        public string DisplayName { get => this.thisDisplayName; set => this.Set(ref this.thisDisplayName, value); }

        // What "Add" adds it under - the box's name, or the suggested one for an empty box.
        internal string NameToAdd => this.thisDisplayName.Trim().Length > 0 ? this.thisDisplayName.Trim() : this.SuggestedName;
    }

    // One of the board's other KiCad nets, listed under the suggestions, with the name proposed for it.
    public sealed class ComponentPaneOtherNet(string kiCadNetName)
        : ComponentPaneNetToAdd(ImportantSignalSuggestions.StartingNameOf(kiCadNetName ?? string.Empty), kiCadNetName);

    // One signal suggested from the board's KiCad nets, its name saying what the net is.
    public sealed class ComponentPaneSignalSuggestion(string suggestedName, string kiCadNetName)
        : ComponentPaneNetToAdd(suggestedName, kiCadNetName)
    {
        // The same suggestion as `other` - its net and the name suggested; never what was typed.
        internal bool IsTheSameSuggestion(ComponentPaneSignalSuggestion other) =>
            string.Equals(this.KiCadNetName, other.KiCadNetName, StringComparison.Ordinal) &&
            string.Equals(this.SuggestedName, other.SuggestedName, StringComparison.Ordinal);
    }
}
