using System;
using System.Collections.Generic;
using System.Linq;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // WHAT "ALSO ON OTHER BOARDS" DID, BOARD BY BOARD, IN WORDS (CRT 4.0 step 2; out of the tab and
    // tested on its own - code review, 2026-10-26). The host (TabDrafts.ApplyToOtherBoardsAsync) says
    // what came of each board; Said is the line under the pane's title (ComponentPaneWording.AlsoDone).
    //
    //   - A BOARD WHERE ANYTHING CHANGED is "Changed" - a component of the part, or a file.
    //   - WHAT WAS NOT CHANGED IS SAID: the board as a whole when nothing on it changed, for one reason
    //     ("Not changed on B - its draft could not be saved"); else each component not changed, by its
    //     label ("[U7] not changed on B - ...") - a part not changed never hides behind a file that went
    //     in, nor one 4164 behind another.
    //   - EACH FILE NOT PUT ON A BOARD it was ticked for is said by the file.
    // ###########################################################################################
    internal sealed class AlsoOnOtherBoardsOutcome
    {
        private readonly List<string> thisChanged = new();
        private readonly List<(string Board, string Why)> thisNotChanged = new();
        private readonly List<(string Label, string Board, string Why)> thisComponentsNotChanged = new();
        private readonly List<(string File, string Board, string Why)> thisFilesNotPut = new();

        // A board of which no draft could be made - nothing on it was tried.
        public void NoDraft(string board) => this.thisNotChanged.Add((board, ComponentPaneWording.AlsoNoDraft));

        // A file ticked for `board` that did not go in, and why.
        public void FileNotPut(string file, string board, string why) => this.thisFilesNotPut.Add((file, board, why));

        // ###########################################################################################
        // One board, once everything ticked for it was tried: whether anything there changed, and
        // each ticked component of it that was not changed, with why.
        // ###########################################################################################
        public void Board(string board, bool changed, IReadOnlyList<(string Label, string Why)> componentsNotChanged)
        {
            ArgumentNullException.ThrowIfNull(componentsNotChanged);

            if (changed)
            {
                this.thisChanged.Add(board);
            }

            if (componentsNotChanged.Count == 0)
            {
                return;
            }

            if (!changed && componentsNotChanged.Select(component => component.Why).Distinct(StringComparer.Ordinal).Count() == 1)
            {
                this.thisNotChanged.Add((board, componentsNotChanged[0].Why));
            }
            else
            {
                this.thisComponentsNotChanged.AddRange(componentsNotChanged.Select(component => (component.Label, board, component.Why)));
            }
        }

        public string Said => ComponentPaneWording.AlsoDone(this.thisChanged, this.thisNotChanged, this.thisFilesNotPut, this.thisComponentsNotChanged);
    }
}
