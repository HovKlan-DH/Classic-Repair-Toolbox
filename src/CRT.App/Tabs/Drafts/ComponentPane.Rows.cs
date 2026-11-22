using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Primitives;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.LogicalTree;
using Avalonia.Threading;
using Avalonia.VisualTree;
using Handlers.DataHandling;
using Handlers.Theming;
using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.ComponentModel;
using System.Globalization;
using System.Linq;
using System.Runtime.CompilerServices;
using System.Threading.Tasks;

namespace CRT
{
    // ###########################################################################################
    // The rows of a component's pins, replacement parts and links as the pane holds them (ComponentPane.axaml.cs shows and commits
    // them). Moved out of ComponentPane.axaml.cs as it passed 1,500 lines (code review, 2026-11-22).
    // ###########################################################################################

    // ###########################################################################################
    // One pin as the pane holds it. A drop-down list's empty write (a list built before its value
    // arrived) is not a change: a direction is picked, never cleared, in the pane.
    // ###########################################################################################
    public sealed class ComponentPanePinRow : INotifyPropertyChanged
    {
        private string thisPin = string.Empty;
        private string thisName = string.Empty;
        private string thisDirection = string.Empty;

        public event PropertyChangedEventHandler? PropertyChanged;

        public string Region { get; init; } = string.Empty;

        public string Pin { get => this.thisPin; set => this.Set(ref this.thisPin, value ?? string.Empty); }

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value ?? string.Empty); }

        public string? Direction
        {
            get => this.thisDirection.Length == 0 ? null : this.thisDirection;
            set
            {
                if (value is not null)
                {
                    this.Set(ref this.thisDirection, value);
                }
            }
        }

        // The vocabulary with every other direction the board's pins use (ComponentVocabulary.PinDirectionsOf - owner request,
        // 2026-11-11), and this pin's own word when it is still not among them.
        public IReadOnlyList<string> DirectionChoices =>
            ComponentVocabulary.IsKnownOrEmpty(this.thisBoardDirections, this.thisDirection)
                ? this.thisBoardDirections
                : [.. this.thisBoardDirections, this.thisDirection];

        private IReadOnlyList<string> thisBoardDirections = ComponentVocabulary.PinDirections;

        // The directions the board offers, as the pane last read them - the list changed only when its words did, as a list
        // replaced under a drop-down lets go of what it shows.
        public void UseBoardDirections(IReadOnlyList<string> directions)
        {
            ArgumentNullException.ThrowIfNull(directions);

            if (directions.SequenceEqual(this.thisBoardDirections, StringComparer.Ordinal))
            {
                return;
            }

            this.thisBoardDirections = directions;
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.DirectionChoices)));

            // A list replaced lets go of its selection and writes "nothing" back, which Direction ignores - so the box is told
            // the direction again, or it would show none (CLAUDE.md, the picture cards' drop-downs).
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.Direction)));
        }

        // The net the pin is on on this board, from its KiCad data - shown, never part of the entry.
        public string Net { get; private set; } = string.Empty;

        public bool HasNet => this.Net.Length > 0;

        public void UseNet(string? net)
        {
            string wanted = net ?? string.Empty;

            if (string.Equals(this.Net, wanted, StringComparison.Ordinal))
            {
                return;
            }

            this.Net = wanted;
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.Net)));
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.HasNet)));
        }

        // A pin the component's case does not have - an error, marked on its box (ShowPinProblems).
        public bool PinOutside => this.PinProblem.Length > 0;

        public string PinProblem { get; private set; } = string.Empty;

        // The case that has no such pin - null when it can be on it.
        public void UseCaseThatRefuses(string? package)
        {
            string wanted = package is null ? string.Empty : BoardDataChecks.PinNotOnPackage(this.thisPin, package);

            if (!string.Equals(this.PinProblem, wanted, StringComparison.Ordinal))
            {
                this.PinProblem = wanted;
                this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.PinProblem)));
                this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(nameof(this.PinOutside)));
            }
        }

        public static ComponentPanePinRow From(ComponentPinEntry pin) => new()
        {
            Region = pin.Region ?? string.Empty,
            Pin = pin.Pin ?? string.Empty,
            Name = pin.Name ?? string.Empty,
            Direction = string.IsNullOrWhiteSpace(pin.Direction) ? null : pin.Direction.Trim(),
        };

        public ComponentPinEntry ToEntry() => new()
        {
            Region = this.Region,
            Pin = this.thisPin,
            Name = this.thisName,
            Direction = this.thisDirection,
        };

        private void Set(ref string field, string value, [CallerMemberName] string? name = null)
        {
            if (string.Equals(field, value, StringComparison.Ordinal))
            {
                return;
            }

            field = value;
            this.PropertyChanged?.Invoke(this, new PropertyChangedEventArgs(name));
        }
    }

    // One replacement part as the pane holds it - the kind picked as a pin's direction is, its note of
    // several lines.
    public sealed class ComponentPaneReplacementRow : ComponentPaneRegionalRow
    {
        private string thisPartNumber = string.Empty;
        private string thisKind = string.Empty;
        private string thisNote = string.Empty;

        public string PartNumber { get => this.thisPartNumber; set => this.Set(ref this.thisPartNumber, value ?? string.Empty); }

        public string? Kind
        {
            get => this.thisKind.Length == 0 ? null : this.thisKind;
            set
            {
                if (value is not null)
                {
                    this.Set(ref this.thisKind, value);
                }
            }
        }

        public string Note { get => this.thisNote; set => this.Set(ref this.thisNote, value ?? string.Empty); }

        // ComponentPaneEdits.WithReplacements' rule for a row it leaves out.
        public override bool HoldsNothingYet =>
            string.IsNullOrWhiteSpace(this.thisPartNumber) && string.IsNullOrWhiteSpace(this.thisKind) && string.IsNullOrWhiteSpace(this.thisNote);

        public IReadOnlyList<string> KindChoices =>
            ComponentVocabulary.IsKnownOrEmpty(ComponentVocabulary.ReplacementKinds, this.thisKind)
                ? ComponentVocabulary.ReplacementKinds
                : [.. ComponentVocabulary.ReplacementKinds, this.thisKind];

        public static ComponentPaneReplacementRow From(ComponentReplacementEntry row) => new()
        {
            Region = row.Region ?? string.Empty,
            PartNumber = row.PartNumber ?? string.Empty,
            Kind = string.IsNullOrWhiteSpace(row.Kind) ? null : row.Kind.Trim(),
            Note = row.Note ?? string.Empty,
        };

        public ComponentReplacementEntry ToEntry() => new()
        {
            Region = this.Region,
            PartNumber = this.thisPartNumber,
            Kind = this.thisKind,
            Note = this.thisNote,
        };
    }

    // One link as the pane holds it.
    public sealed class ComponentPaneLinkRow : ComponentPaneRegionalRow
    {
        private string thisName = string.Empty;
        private string thisUrl = string.Empty;

        public string Name { get => this.thisName; set => this.Set(ref this.thisName, value ?? string.Empty); }

        public string Url
        {
            get => this.thisUrl;
            set
            {
                this.Set(ref this.thisUrl, value ?? string.Empty);
                this.Raise(nameof(this.UrlNotWeb));
                this.Raise(nameof(this.CanOpen));
            }
        }

        // "Open link" is on while the address is one a browser can be sent to (LinkTarget) - a board link's rule.
        public bool CanOpen => LinkTarget.For(this.thisUrl) is not null;

        // An address the checks refuse, marked on its box (ComponentPaneBoardLinkRow says why).
        public bool UrlNotWeb => BoardDataChecks.IsRefusedLink(this.thisUrl);

        public string NotWebText => ComponentPaneWording.LinkNotWeb;

        public static ComponentPaneLinkRow From(ComponentLinkEntry row) => new()
        {
            Region = row.Region ?? string.Empty,
            Name = row.Name ?? string.Empty,
            Url = row.Url ?? string.Empty,
        };

        public ComponentLinkEntry ToEntry() => new() { Region = this.Region, Name = this.thisName, Url = this.thisUrl };
    }
}
