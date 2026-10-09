using System.Collections.Generic;
using System.Linq;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // A PERSON, IN BOLD (owner request, 2026-10-09: "do 'bold' all maintainers and contributors in
    // the UI, so it gets more visibility who we are talking about").
    //
    // Every line on the Maintainer tab that names a maintainer or a contributor names them through
    // here, as runs the tab draws with TabMaintainer.ShowCounts - the person bold (ReviewNoteRun's
    // IsCount is what is drawn bold), the rest of the sentence plain. A name comes first; an address
    // beside it stays plain, in brackets, and is the person itself only when there is no name.
    // ###########################################################################################
    public static class PersonRuns
    {
        public static ReviewNoteRun Person(string person) => new(person, IsCount: true);

        public static ReviewNoteRun Plain(string text) => new(text, IsCount: false);

        // "**Anna** (anna@example.com)", "**anna@example.com**" with no name, "**Anna**" with no
        // address - and `nobody`, plain, with neither.
        public static IReadOnlyList<ReviewNoteRun> NameAndAddress(string? name, string? address, string nobody)
        {
            string shownName = name?.Trim() ?? string.Empty;
            string shownAddress = address?.Trim() ?? string.Empty;

            if (shownName.Length == 0)
                return shownAddress.Length == 0 ? [PersonRuns.Plain(nobody)] : [PersonRuns.Person(shownAddress)];

            return shownAddress.Length == 0
                ? [PersonRuns.Person(shownName)]
                : [PersonRuns.Person(shownName), PersonRuns.Plain($" ({shownAddress})")];
        }

        // The person a sentence names - bold - or `nobody`, plain, when it names nobody.
        public static ReviewNoteRun PersonOr(string? person, string nobody) =>
            string.IsNullOrWhiteSpace(person) ? PersonRuns.Plain(nobody) : PersonRuns.Person(person.Trim());

        // What the runs say, as one text.
        public static string Text(IEnumerable<ReviewNoteRun> runs) => string.Concat(runs.Select(run => run.Text));
    }
}
