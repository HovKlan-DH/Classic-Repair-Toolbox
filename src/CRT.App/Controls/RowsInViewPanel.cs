using Avalonia;
using Avalonia.Controls;
using Avalonia.Input;
using Avalonia.Interactivity;
using Avalonia.VisualTree;
using System.Collections.Generic;
using System.Linq;

namespace CRT
{
    // ###########################################################################################
    // A LIST THAT BUILDS ONLY THE ROWS IN VIEW, ITS ROWS EDITORS (2026-11-13, measured - ComponentPane.axaml's
    // "RowsInView" style, the "Overview" tab's list). Avalonia's VirtualizingStackPanel, with what it does not do for rows
    // a person types in (Avalonia 12.1.6, read from its source):
    //
    // *** THE ROW WITH THE KEYBOARD STAYS BUILT (code review, 2026-11-13). *** The panel keeps a row scrolled out of view
    // only while it is the list's TabOnceActiveElement - and ItemsControl sets that to the control that took the focus, the
    // box inside the row, never the row. So the row being typed in was let go of as it scrolled away: its box left the
    // window, the caret was gone mid-edit. The row is made the list's TabOnceActiveElement whenever the focus comes into it,
    // after ItemsControl's own class handler (FollowFocusIntoRow, on the list), so it stands.
    //
    // *** AND STAYS BUILT WHILE THE FOCUS MOVES INSIDE IT (code review, 2026-11-14). *** Out of view, the panel lets go of
    // the row the moment the TabOnceActiveElement moves off it - and ItemsControl moves it to the next box of the row before
    // anything after it can put it back: Tab from one box to the next, out of view, lost the focus. So the row is brought
    // into view FIRST, by this panel's own handler, which the event reaches before the list (BringRowIntoView).
    //
    // *** TAB GOES BY ROW NUMBER (code review, 2026-11-14). *** Avalonia tabs in the order of the panel's children, and a
    // row built again is added at their end: scrolled down and back up, Tab from row 19 went to row 18, Shift+Tab from row
    // 20 out of the list, and rows not built were skipped. Avalonia asks an ancestor that is ICustomKeyboardNavigation for
    // the next element first. Inside a row its own way stands; from a row's last box (its first, going back) this panel
    // builds the next row by number and gives its first box (the one before's last), and out of the list after the last
    // row and before the first - as Tab went when every row was built. Coming into the list from outside: its first row,
    // or its last going back. A row's boxes are what Avalonia's TabNavigation.IsTabStop takes, in the order of the row's
    // own tree - which a row's own children keep.
    // ###########################################################################################
    public class RowsInViewPanel : VirtualizingStackPanel, ICustomKeyboardNavigation
    {
        // Rows passed over looking for one with a box the keys can reach, before Avalonia's own way is taken - every row of
        // these lists has one, and each row looked at is built.
        private const int RowsLookedAtMost = 20;

        public RowsInViewPanel()
        {
            this.AddHandler(InputElement.GotFocusEvent, this.BringRowIntoView, RoutingStrategies.Bubble, handledEventsToo: true);
        }

        protected override void OnItemsControlChanged(ItemsControl? oldValue)
        {
            base.OnItemsControlChanged(oldValue);

            oldValue?.RemoveHandler(InputElement.GotFocusEvent, this.FollowFocusIntoRow);
            this.ItemsControl?.AddHandler(InputElement.GotFocusEvent, this.FollowFocusIntoRow, RoutingStrategies.Bubble, handledEventsToo: true);
        }

        // Before ItemsControl moves its TabOnceActiveElement: a row out of view - the one kept built - into view.
        private void BringRowIntoView(object? sender, FocusChangedEventArgs e)
        {
            if (this.RowOf(e.Source as Visual) is { } row &&
                this.IndexFromContainer(row) is int index and >= 0 &&
                (index < this.FirstRealizedIndex || index > this.LastRealizedIndex))
            {
                this.ScrollIntoView(index);
            }
        }

        // After ItemsControl moved it to the box: the row it is in, which the panel keeps built.
        private void FollowFocusIntoRow(object? sender, FocusChangedEventArgs e)
        {
            if (this.ItemsControl is { } list && this.RowOf(e.Source as Visual) is { } row)
            {
                KeyboardNavigation.SetTabOnceActiveElement(list, row);
            }
        }

        public (bool handled, IInputElement? next) GetNext(IInputElement element, NavigationDirection direction)
        {
            if (direction is not (NavigationDirection.Next or NavigationDirection.Previous))
            {
                return (false, null);
            }

            bool forward = direction == NavigationDirection.Next;

            if (this.RowOf(element as Visual) is not { } row)
            {
                // Coming in from outside - Avalonia's own way led into the list: its first row going forward, its last back.
                if (element is Visual visual && this.IsVisualAncestorOf(visual))
                {
                    return (false, null);
                }

                return this.ComingIn(forward) is { } into ? (true, into) : (false, null);
            }

            List<InputElement> stops = RowsInViewPanel.TabStopsIn(row);

            // Inside the row: Avalonia's own way.
            if (stops.Count == 0 || !ReferenceEquals(element, forward ? stops[^1] : stops[0]))
            {
                return (false, null);
            }

            int index = this.IndexFromContainer(row);
            int other = index + (forward ? 1 : -1);

            if (index < 0)
            {
                return (false, null);
            }

            if (other < 0 || other >= this.Items.Count)
            {
                // Out of the list: Avalonia's next after it, or before it - unless that is in another such list.
                return (true, this.IntoTheListAfter(forward));
            }

            return this.BoxOfRowFrom(other, forward) is { } box ? (true, box) : (false, null);
        }

        private IInputElement? ComingIn(bool forward) => this.BoxOfRowFrom(forward ? 0 : this.Items.Count - 1, forward);

        // ###########################################################################################
        // Leaving the list, Avalonia goes on from what this panel answers as it stands - into a list that follows at once
        // by the order of ITS children. So the element after the list (before it, going back) is looked up the way
        // Avalonia's own GetNextTabOutside does, from the deepest of its last children or its first child; in another such
        // list it is that list's first row (last), else null: Avalonia's own.
        // ###########################################################################################
        private IInputElement? IntoTheListAfter(bool forward)
        {
            IInputElement edge = this;

            if (forward)
            {
                while (((Visual)edge).GetVisualChildren().OfType<InputElement>().LastOrDefault(child => child.IsVisible) is { } child)
                {
                    edge = child;
                }
            }
            else if (this.GetVisualChildren().OfType<InputElement>().FirstOrDefault(child => child.IsVisible) is { } first)
            {
                edge = first;
            }

            IInputElement? outside = TopLevel.GetTopLevel(this)?.FocusManager?.FindNextElement(
                forward ? NavigationDirection.Next : NavigationDirection.Previous,
                new FindNextElementOptions { FocusedElement = edge });

            return outside is Visual visual &&
                   visual.FindAncestorOfType<RowsInViewPanel>() is { } other &&
                   !ReferenceEquals(other, this)
                ? other.ComingIn(forward)
                : null;
        }

        // The first box (going forward) or the last (going back) of the row at `index`, or of the next row that has one -
        // each built for it and brought into view.
        private IInputElement? BoxOfRowFrom(int index, bool forward)
        {
            int count = this.Items.Count;

            for (int looked = 0; looked < RowsLookedAtMost && index >= 0 && index < count; looked++, index += forward ? 1 : -1)
            {
                if ((this.ScrollIntoView(index) ?? this.ContainerFromIndex(index)) is { } row &&
                    RowsInViewPanel.TabStopsIn(row) is { Count: > 0 } stops)
                {
                    return forward ? stops[0] : stops[^1];
                }
            }

            return null;
        }

        // ###########################################################################################
        // The row's boxes the keys reach, in order: what Avalonia's TabNavigation.IsTabStop takes - focusable, a tab stop,
        // visible, enabled - in the row's own tree, nothing under a control hidden or turned off.
        // ###########################################################################################
        internal static List<InputElement> TabStopsIn(Visual row)
        {
            var stops = new List<InputElement>();

            void Walk(Visual parent)
            {
                foreach (Visual child in parent.GetVisualChildren())
                {
                    if (child is InputElement element)
                    {
                        if (!element.IsVisible || !element.IsEffectivelyEnabled)
                        {
                            continue;
                        }

                        if (element.Focusable && KeyboardNavigation.GetIsTabStop(element))
                        {
                            stops.Add(element);
                        }
                    }

                    Walk(child);
                }
            }

            Walk(row);

            return stops;
        }

        // The row - this panel's child - that holds `visual`, if it is one the list knows as a row.
        private Control? RowOf(Visual? visual) =>
            visual?.GetSelfAndVisualAncestors()
                .TakeWhile(ancestor => !ReferenceEquals(ancestor, this))
                .OfType<Control>()
                .LastOrDefault() is { } child &&
            ReferenceEquals(child.GetVisualParent(), this) &&
            this.IndexFromContainer(child) >= 0
                ? child
                : null;
    }
}
