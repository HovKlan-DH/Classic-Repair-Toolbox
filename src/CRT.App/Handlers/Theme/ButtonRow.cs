namespace Handlers.Theming
{
    // ###########################################################################################
    // THE GAP BETWEEN BUTTONS SIDE BY SIDE - ONE, EVERYWHERE (owner request, 2026-10-30: "Please align
    // everywhere in UI, where it shows multiple buttons on the same line next to each other, that it has
    // the same horizontal spacing, so it looks the same everywhere" - 8px read as "a little too much").
    //
    // A row of buttons in the markup is a horizontal StackPanel's Spacing or a WrapPanel's ItemSpacing of
    // 6 - or, where its buttons are spaced by their own margins, a right margin of 6 - and two buttons
    // sharing a line in a Grid have a 6px column between them. A row built in code uses this. Segmented
    // view switches, the Maintainer tab's screen tabs and link-styled buttons are no row of buttons.
    // ButtonSpacingTests holds every .axaml to it.
    // ###########################################################################################
    public static class ButtonRow
    {
        public const double Gap = 6;
    }
}
