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
    // THE COMPONENT PANE'S TYPING, LEAVING A BOX AND PICKING: a pause in typing, a box left, a drop-down list picked and Enter
    // pressed commit what the pane holds into the draft; the wheel never changes a closed list. Moved out of ComponentPane.axaml.cs
    // as it passed 1,500 lines (code review, 2026-11-22). See ComponentPane.axaml.cs for the file map.
    // ###########################################################################################
    public partial class ComponentPane
    {
        // ---- typing, leaving a box, picking ------------------------------------------------------

        private void OnAnyTextChanged(object? sender, TextChangedEventArgs e)
        {
            if (this.thisShowing || this.thisSession is null)
            {
                return;
            }

            // The new-component form is judged as it is typed in, and goes into the draft only when added.
            if (this.thisIsNew)
            {
                this.OnNewComponentTextChanged();
                return;
            }

            if (ReferenceEquals(e.Source, this.DescriptionBox))
            {
                this.ShowCounter();
            }

            // A picture's name waits for Enter or its box being left - typing it starts no pause. So do
            // a component's label and a category's name (2026-11-01). Nor does a net typed into the new
            // signal's picker: it is no part of the draft until "Add signal", and the commit a pause
            // made swapped the picker's list under it (code review, 2026-10-30).
            if (e.Source is TextBox { DataContext: ComponentPaneBoardPictureRow or ComponentPaneCategoryRow } ||
                ReferenceEquals(e.Source, this.BoardLabelBox) ||
                ReferenceEquals(e.Source, this.NewCategoryNameBox) ||
                (e.Source is Visual typedIn && ReferenceEquals(typedIn.FindAncestorOfType<AutoCompleteBox>(), this.NewSignalNetBox)))
            {
                return;
            }

            // Typed into a field several components differ in: the list of their values narrows to it, and what is typed
            // waits for Enter, a pick or the box being left - a pause would set half a word on all of them while the list
            // is still being read (ComponentPane.Several.cs).
            if (this.FollowTypingInDifferingField(e.Source))
            {
                return;
            }

            this.thisTypingTimer.Stop();
            this.thisTypingTimer.Start();
        }

        private void OnAnyLostFocus(object? sender, RoutedEventArgs e)
        {
            // ###########################################################################################
            // *** THE FOCUS GOING BETWEEN A BOX AND ITS OWN LIST IS NO BOX LEFT (code review, 2026-11-22). *** A press on a row of an
            // "Expected reading" or the package's list takes the focus to the row, and the box takes it straight back while the press
            // is still picking (Avalonia 12.1.6): each was a box left, and the half-typed word went into the draft as a step of its own,
            // the pick a second step after it - one "Undo" left a reading nobody chose. The pick is committed as it is made, after a
            // typing pause or the box left. (The keys pick no row of these lists - Down highlights none - so Enter only closes one,
            // committing once.)
            // ###########################################################################################
            if (e is FocusChangedEventArgs { NewFocusedElement: Visual to } && e.Source is Visual from &&
                ((ComponentPane.AutoCompleteBoxOf(from) is { } box && ComponentPane.IsInBox(box, to)) ||
                 (ComponentPane.AutoCompleteBoxOf(to) is { } into && ComponentPane.IsInBox(into, from))))
            {
                return;
            }

            this.thisTypingTimer.Stop();
            this.CommitHeld(leftABox: true);
            this.ShowPinout();
        }

        // A drop-down list picked (a pin's direction, a part's kind) - into the draft once the pick
        // has reached the row.
        private void OnAnySelectionChanged(object? sender, SelectionChangedEventArgs e)
        {
            // Which component "Make all like" copies is no data of the draft (code review, 2026-11-09: each refill of its list
            // posted a commit of every list of every component picked, which changed nothing).
            if (this.thisShowing || this.thisSession is null || e.Source is not ComboBox || ReferenceEquals(e.Source, this.MakeAllLikeSourceBox))
            {
                return;
            }

            // The new-component form holds what it holds until it is added - judged as it is picked in.
            if (this.thisIsNew)
            {
                this.OnNewComponentTextChanged();
                return;
            }

            Dispatcher.UIThread.Post(() => this.CommitHeld(), DispatcherPriority.Background);
        }

        // ###########################################################################################
        // THE WHEEL NEVER CHANGES A CLOSED DROP-DOWN LIST HERE (code review, 2026-11-04). A ComboBox with the
        // focus steps its selection under the wheel, so scrolling the page with the pointer over the category
        // just picked gave the component another one, a step of the draft nobody meant - and over the clone
        // panel's board, each notch read another board's file. A pick is made from the open list, as on Main's
        // device picker. The page scrolls on, by the wheel's own step (Avalonia's ScrollContentPresenter), as it
        // does over any other box; an open list keeps its wheel, and one without the focus never took it.
        // ###########################################################################################
        private const double WheelStep = 50;

        private void OnPaneWheel(object? sender, PointerWheelEventArgs e)
        {
            if (e.Source is not Visual source ||
                source.FindAncestorOfType<ComboBox>(includeSelf: true) is not { IsDropDownOpen: false, IsFocused: true } box)
            {
                return;
            }

            e.Handled = true;

            if (box.FindAncestorOfType<ScrollViewer>() is { } page)
            {
                page.Offset = page.Offset.WithY(page.Offset.Y - e.Delta.Y * ComponentPane.WheelStep);
            }
        }

        // Enter in a one-line box takes what it says into the draft at once - unless the list of what several components
        // have takes the key first (ComponentPane.Several.cs: Enter there picks the line lit).
        private void OnPaneKeyDown(object? sender, KeyEventArgs e)
        {
            if (this.TakeDifferingListKey(e) || this.AddOnEnter(e))
            {
                return;
            }

            if (e.Key == Key.Enter && e.Source is TextBox { AcceptsReturn: false })
            {
                this.thisTypingTimer.Stop();
                this.CommitHeld();
                this.ShowPinout();
            }
        }

        // The AutoCompleteBox `element` stands in - its text box, or the box itself - null for any other.
        private static AutoCompleteBox? AutoCompleteBoxOf(object? element) =>
            element is Visual visual ? visual as AutoCompleteBox ?? visual.FindAncestorOfType<AutoCompleteBox>() : null;
    }
}
