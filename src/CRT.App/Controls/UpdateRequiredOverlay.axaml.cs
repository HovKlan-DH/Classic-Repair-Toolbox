using System;
using System.Collections.Generic;
using System.Linq;
using Avalonia.Controls;
using Avalonia.Interactivity;
using Handlers.Online;

namespace CRT
{
    // ###########################################################################################
    // "CRT HAS TO BE UPDATED" OVER A WHOLE TAB - the Drafts tab and the Maintainer tab, when the
    // server turns this version of CRT away. The owner asked (2026-10-09) for a full-page modal "or
    // alike" there that cannot be closed and states that the app needs to be updated, "so it is
    // very clear for the user they must update". Which tab, and when, is AppUpdateRequirement's;
    // the words are AppUpdateRequiredWording's; Main.UpdateRequired.cs puts the two together.
    //
    // While shown:
    //   - it lies over the whole tab and takes every click meant for it, and the tab beneath it
    //     FADES and is DISABLED, so nothing in it can take focus;
    //   - it takes every key meant for the tab, on the tunnel route at the tab - a control that had
    //     focus when it appeared would otherwise still take Enter - leaving its own button's.
    //     The fading and the keys are OverlayCover's, BusyOverlay's too (code review, 2026-10-09);
    //   - the other tabs, the tab strip and the window's banners stay usable: only this tab's work
    //     needs the server, so only this tab is covered.
    //
    // *** IT CANNOT BE CLOSED - THERE IS NO Hide. *** Nothing but a newer CRT changes the server's
    // answer (AppUpdateRequirement), so its only button is the way out: install the update CRT has
    // found, or open the download page. Show again only changes what it says.
    //
    // A shared control (Controls/): it touches nothing of CRT's orchestration. The host hands it
    // the view and answers ActionClicked.
    // ###########################################################################################
    public partial class UpdateRequiredOverlay : UserControl
    {
        // The controls beneath it in the host's grid, disabled and faded while it is shown - with the
        // opacity each had, as OverlayCover keeps it.
        private readonly Dictionary<Control, double> thisCovered = [];

        // The tab's keys, taken while it is shown - for as long as CRT runs.
        private IDisposable? thisKeysTaken;

        public UpdateRequiredOverlay()
        {
            this.InitializeComponent();
        }

        // Whether it is on the tab - from the first Show for as long as CRT runs.
        public bool IsShown => this.IsVisible;

        // What it says, or null before the first Show.
        public AppUpdateRequiredView? View { get; private set; }

        // The button was pressed - View.ButtonInstalls says which it offered.
        public event EventHandler? ActionClicked;

        // ###########################################################################################
        // Covers the tab with `view`, or - already shown - says `view` instead.
        // ###########################################################################################
        public void Show(AppUpdateRequiredView view)
        {
            ArgumentNullException.ThrowIfNull(view);

            this.View = view;

            this.HeadingText.Text = view.Heading;
            this.ReasonText.Text = view.Reason;
            this.WhatItStopsText.Text = view.WhatItStops;
            this.RestOfCrtText.Text = view.RestOfCrt;
            this.ReadyText.Text = view.ReadyLine ?? string.Empty;
            this.ReadyText.IsVisible = view.ReadyLine is not null;
            this.ActionButton.Content = view.ButtonLabel;

            if (this.IsVisible)
                return;

            this.IsVisible = true;
            this.CoverHost();
        }

        // ###########################################################################################
        // Everything else in the host's grid fades and is turned off; the host's keys are taken.
        // Never undone - see the header.
        // ###########################################################################################
        private void CoverHost()
        {
            if (this.Parent is not Panel host)
                return;

            OverlayCover.FadeSiblings(this, this.thisCovered);

            foreach (Control covered in this.thisCovered.Keys)
                covered.IsEnabled = false;

            // A key for anything in the tab but this overlay goes nowhere.
            this.thisKeysTaken ??= OverlayCover.TakeKeys(host, except: this);
        }

        private void OnActionClick(object? sender, RoutedEventArgs e) =>
            this.ActionClicked?.Invoke(this, EventArgs.Empty);

        // The controls it has covered - for tests.
        internal IReadOnlyList<Control> CoveredForTests => this.thisCovered.Keys.ToList();

        // Presses the button, as a click would - for tests.
        internal void ClickActionForTests() => this.OnActionClick(this, new RoutedEventArgs());
    }
}
