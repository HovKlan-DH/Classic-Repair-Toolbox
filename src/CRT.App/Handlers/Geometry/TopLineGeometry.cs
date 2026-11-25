namespace Handlers.Geometry
{
    // ###########################################################################################
    // A LINE OF WORDS WITH ITS BUTTONS BESIDE IT, OR UNDER IT ON A NARROW TAB (2026-11-25, found looking for what else
    // was out of reach after the owner's report of the "Maintainer" tab's sign-in on Linux): the "Contribute" tab's top
    // line - "Contribute", what the tab is for, then "My drafts", "My submissions" and "Add a new board" beside it. On
    // CRT's narrowest window the buttons left the words some 120 wide; they ran to a dozen lines and pushed the whole
    // workspace below the window. Beside the words while that leaves them `LeastTextWidth`, else under them.
    // ###########################################################################################
    public static class TopLineGeometry
    {
        // The least room the words keep beside the buttons - about one line of the tab's top sentence.
        public const double LeastTextWidth = 320;

        // Whether the buttons, `buttonsWidth` wide, go under the words in a row `rowWidth` wide with `gap` between the
        // two. A row not laid out yet (no width) keeps them beside.
        public static bool ButtonsGoUnder(double rowWidth, double buttonsWidth, double gap)
        {
            if (!(rowWidth > 0))
            {
                return false;
            }

            return rowWidth - buttonsWidth - gap < TopLineGeometry.LeastTextWidth;
        }
    }
}
