using System.Text.RegularExpressions;
using System.Xml.Linq;
using Handlers.DataHandling;

namespace CRT.Server.Tests.ApiCompatibility
{
    // ###########################################################################################
    // *** EVERY CRT EVER RELEASED KEEPS WORKING (owner request, 2026-10-04: "It is important that all
    // older versions will continue to work"). ***
    //
    // The compiler cannot see this. CRT and the server share CRT.Data's records, so renaming a field
    // moves BOTH ends of the code at once - it compiles, ReviewWireContractTests passes - while the
    // CRTs already installed still send and read the OLD name. Only a copy of what a released CRT was
    // built against can catch that, so this keeps one:
    //
    //   - When CRT.App.csproj's InformationalVersion is a RELEASE (no "-", e.g. "3.0.0"), its API
    //     surface (ApiSurface) must be frozen in this folder as crt-<version>.txt. The first run
    //     without it WRITES it and fails once, saying so - commit the file with the release. A
    //     pre-release (3.1.0-beta.2) is not frozen: betas update themselves, the promise is for
    //     releases.
    //   - Every frozen file is then compared with the surface as it is NOW, on every run. A route,
    //     field, enum member or wire name a released CRT relied on that has gone - or changed type -
    //     fails here, named, with the version it breaks.
    //
    // A FAILURE HERE IS A BREAK FOR INSTALLED CRTs, not a test to update. Put the field or route
    // back (add beside it, never rename), or - only when the owner has decided that this break is
    // worth it, and CRTs below a minimum are told to update (ClientVersionPolicy) - list the line in
    // allowed-breaks.txt with " || " and the reason. NEVER edit a crt-*.txt file: it is what a
    // released CRT is.
    //
    // The same surface also keeps the API REVISION honest (2026-10-04) - pre-releases included, where
    // the API may change freely: api-revision-<N>.txt is what a CRT built for revision N may use, and
    // a change it cannot use fails until ClientVersionContract.ApiRevision is raised. See the test.
    // ###########################################################################################
    public sealed class ApiCompatibilityTests
    {
        private const string AllowedBreaksFile = "allowed-breaks.txt";

        private static string Folder() =>
            Path.Combine(ServerRouteTable.RepositoryRoot(), "tests", "CRT.Server.Tests", "ApiCompatibility");

        // -----------------------------------------------------------------------------------
        // The check itself
        // -----------------------------------------------------------------------------------

        [Fact]
        public void Every_released_CRT_still_finds_everything_it_was_released_against()
        {
            SortedSet<string> current = ApiSurface.Build();
            string[] allowed = File.ReadAllLines(Path.Combine(ApiCompatibilityTests.Folder(), ApiCompatibilityTests.AllowedBreaksFile));

            var report = new List<string>();

            foreach (string file in ApiCompatibilityTests.FrozenFiles())
            {
                IReadOnlyList<string> breaks = ApiSurface.Breaks(File.ReadAllLines(file), current, allowed);

                if (breaks.Count > 0)
                    report.Add($"{Path.GetFileName(file)} - installed CRTs of this version would break on:\n  " + string.Join("\n  ", breaks));
            }

            Assert.True(
                report.Count == 0,
                "The API no longer offers what a released CRT relies on. Add beside it instead of renaming or " +
                "removing it - see ApiCompatibilityTests' header and CLAUDE.md, \"Installed CRTs keep working\".\n" +
                string.Join("\n", report));
        }

        [Fact]
        public void A_release_of_CRT_has_its_API_surface_frozen()
        {
            string version = ApiCompatibilityTests.CrtVersionBeingBuilt();

            if (version.Contains('-', StringComparison.Ordinal))
                return;

            string path = Path.Combine(ApiCompatibilityTests.Folder(), $"crt-{version}.txt");

            if (File.Exists(path))
                return;

            var text = new List<string>
            {
                $"# The API surface CRT {version} was released against, frozen by ApiCompatibilityTests on {DateTime.UtcNow:yyyy-MM-dd}.",
                "# NEVER EDIT THIS FILE - every installed CRT of this version depends on each line. See CLAUDE.md, \"Installed CRTs keep working\"."
            };

            text.AddRange(ApiSurface.Build());
            File.WriteAllLines(path, text);

            Assert.Fail(
                $"CRT {version} is a release, so its API surface has now been frozen into {path}. " +
                "Commit that file with the release - the next run passes.");
        }

        // -----------------------------------------------------------------------------------
        // The API revision
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** THE API REVISION GOES UP WHEN THE API BREAKS FOR A CRT BUILT FOR IT (owner request,
        // 2026-10-04: "how do I know which version of app uses which version of server? ... I do not
        // have an overview on what kind of specific changes are done to API and if those changes are
        // compatible or not"). ***
        //
        // api-revision-<N>.txt holds everything a CRT built for revision N may use. It is written the
        // first time N is the current revision - that run fails once, so the file is committed with
        // the change - and GROWN as the API grows: something added while N is current is something a
        // later build of N relies on, so taking it away again has to raise the revision too.
        //
        // A break of it fails here until ClientVersionContract.ApiRevision goes up by one, and the
        // server then turns CRTs of an older revision away from the submission and review routes with
        // "update CRT" (ClientVersionPolicy). There is no allowed-breaks list here: breaking a
        // revision costs nothing but the number - a released CRT is held by the crt-*.txt check
        // above. What this cannot see - a route given a new meaning with the same shape, a request
        // field the server starts to require - is raised by hand.
        // ###########################################################################################
        [Fact]
        public void The_API_revision_goes_up_whenever_the_API_breaks_for_a_CRT_built_for_it()
        {
            int revision = ClientVersionContract.ApiRevision;
            string path = ApiCompatibilityTests.RevisionFile(revision);
            SortedSet<string> current = ApiSurface.Build();

            if (!File.Exists(path))
            {
                ApiCompatibilityTests.WriteRevisionFile(path, revision, current);

                Assert.Fail(
                    $"API revision {revision} is new, so the API it stands for has been recorded in {path}. " +
                    "Commit that file with the change - the next run passes.");
            }

            string[] recorded = File.ReadAllLines(path);
            IReadOnlyList<string> breaks = ApiSurface.Breaks(recorded, current, []);

            Assert.True(
                breaks.Count == 0,
                $"The API changed in a way a CRT built for API revision {revision} cannot use. Raise " +
                $"ClientVersionContract.ApiRevision (CRT.Data) to {revision + 1}: the server then tells CRTs of an " +
                $"older revision to update, and the next run records revision {revision + 1}.\n  " +
                string.Join("\n  ", breaks));

            // Grown to everything the API offers now - see the header.
            HashSet<string> recordedLines = recorded
                .Select(line => line.Trim())
                .Where(line => line.Length > 0 && !line.StartsWith('#'))
                .ToHashSet(StringComparer.Ordinal);

            if (!current.SetEquals(recordedLines))
                ApiCompatibilityTests.WriteRevisionFile(path, revision, current);
        }

        // Lowering the number would hand a newer shape an older revision's name - and serve the CRTs
        // that were told to update. A revision file above the current revision means it went back.
        [Fact]
        public void The_API_revision_never_goes_back()
        {
            int[] recorded = Directory.EnumerateFiles(ApiCompatibilityTests.Folder(), "api-revision-*.txt")
                .Select(file => Path.GetFileNameWithoutExtension(file)["api-revision-".Length..])
                .Select(number => int.TryParse(number, out int value) ? value : int.MaxValue)
                .ToArray();

            Assert.True(ClientVersionContract.ApiRevision >= 1);
            Assert.All(recorded, number => Assert.True(
                number <= ClientVersionContract.ApiRevision,
                $"api-revision-{number}.txt is above ClientVersionContract.ApiRevision ({ClientVersionContract.ApiRevision}) - the revision went back."));
        }

        private static string RevisionFile(int revision) =>
            Path.Combine(ApiCompatibilityTests.Folder(), $"api-revision-{revision}.txt");

        private static void WriteRevisionFile(string path, int revision, IEnumerable<string> surface)
        {
            var text = new List<string>
            {
                $"# The API a CRT built for API revision {revision} may use - recorded by ApiCompatibilityTests, which adds to it as the API grows.",
                "# Do not edit by hand: a change that breaks a line here raises ClientVersionContract.ApiRevision instead. See CLAUDE.md, \"Installed CRTs keep working\"."
            };

            text.AddRange(surface);
            File.WriteAllLines(path, text);
        }

        // -----------------------------------------------------------------------------------
        // The surface is real
        // -----------------------------------------------------------------------------------

        // A builder that silently found nothing would freeze an empty file and guard nothing.
        [Theory]
        [InlineData("route POST /api/submissions")]
        [InlineData("route POST /api/usage/check-in")]
        [InlineData("body POST /api/accounts/login | .email string")]
        [InlineData("body POST /api/submissions | .files[].sha256 string")]
        [InlineData("body POST /api/review/submissions/{submissionId:long}/approve | .expectedRemovals array of string")]
        [InlineData("answer SessionAnswer | .refreshToken string")]
        [InlineData("answer SessionAnswer | .account.id number")]
        [InlineData("answer SubmissionDetailAnswer | .submission.awaitsYou boolean")]
        [InlineData("answer BlobUploadAnswer | .uploaded number")]
        [InlineData("answer ClientOutdatedAnswer | .minimumVersion string")]
        [InlineData("answer HealthStatus | .version string")]
        [InlineData("answer HealthStatus | .apiRevision number")]
        [InlineData("const ClientVersionContract.ApiRevisionHeader = \"X-CRT-Api-Revision\"")]
        [InlineData("enum SubmissionFileScope | Own -> \"Own\"")]
        [InlineData("const CheckInContract.ControlField = \"control\"")]
        [InlineData("const FeedbackContract.SuccessAnswer = \"Success\"")]
        [InlineData("const SubmissionFormat.UploadTokenHeader = \"X-Submission-Token\"")]
        public void The_surface_holds_what_installed_CRTs_use(string line)
        {
            Assert.Contains(line, ApiSurface.Build());
        }

        // -----------------------------------------------------------------------------------
        // The comparison
        // -----------------------------------------------------------------------------------

        private static readonly string[] Frozen =
        [
            "# a comment",
            "route POST /api/submissions",
            "answer SessionAnswer | .refreshToken string",
            "answer SessionAnswer | .expiresUtc string",
            "enum ApproverRole | Maintainer -> \"Maintainer\""
        ];

        [Fact]
        public void A_surface_that_only_grew_breaks_nothing()
        {
            var current = new SortedSet<string>(ApiCompatibilityTests.Frozen.Skip(1), StringComparer.Ordinal)
            {
                "route POST /api/something-new",
                "answer SessionAnswer | .addedLater string",
                "enum NewEnum | A -> \"A\""
            };

            Assert.Empty(ApiSurface.Breaks(ApiCompatibilityTests.Frozen, current, []));
        }

        [Fact]
        public void A_renamed_field_a_retyped_field_and_a_removed_route_are_each_a_break()
        {
            var current = new SortedSet<string>(StringComparer.Ordinal)
            {
                "answer SessionAnswer | .token string",
                "answer SessionAnswer | .expiresUtc number",
                "enum ApproverRole | Maintainer -> \"Maintainer\""
            };

            Assert.Equal(
                [
                    "gone:  answer SessionAnswer | .expiresUtc string",
                    "gone:  answer SessionAnswer | .refreshToken string",
                    "gone:  route POST /api/submissions"
                ],
                ApiSurface.Breaks(ApiCompatibilityTests.Frozen, current, []));
        }

        // An installed CRT reads an enum by its members; one it has never seen fails its parse.
        [Fact]
        public void A_member_added_to_an_enum_a_released_CRT_reads_is_a_break()
        {
            var current = new SortedSet<string>(ApiCompatibilityTests.Frozen.Skip(1), StringComparer.Ordinal)
            {
                "enum ApproverRole | Reviewer -> \"Reviewer\""
            };

            Assert.Equal(
                ["added: enum ApproverRole | Reviewer -> \"Reviewer\""],
                ApiSurface.Breaks(ApiCompatibilityTests.Frozen, current, []));
        }

        // A break the owner decided on is listed once, with its reason, and not reported again.
        [Fact]
        public void A_break_listed_in_the_allowed_file_is_not_reported_again()
        {
            var current = new SortedSet<string>(StringComparer.Ordinal)
            {
                "answer SessionAnswer | .refreshToken string",
                "answer SessionAnswer | .expiresUtc string",
                "enum ApproverRole | Maintainer -> \"Maintainer\""
            };

            string[] allowed =
            [
                "# owner decisions",
                "route POST /api/submissions || replaced by /api/v2/submissions; CRT < 3.4.0 told to update (2027-01-01)"
            ];

            Assert.Empty(ApiSurface.Breaks(ApiCompatibilityTests.Frozen, current, allowed));
        }

        // -----------------------------------------------------------------------------------
        // Every answer is a named record
        // -----------------------------------------------------------------------------------

        // ###########################################################################################
        // *** THE SURFACE ONLY SEES NAMED RECORDS. *** An anonymous object written as an answer is
        // invisible to it - which is how the sign-in answer and a submission's detail went unchecked
        // until 2026-10-04. So an endpoint may write an anonymous object only to carry a sentence
        // (`message`, `error`) or a findings list (`errors`); anything with data in it is a CRT.Data
        // record, which the surface then covers.
        // ###########################################################################################
        [Fact]
        public void Every_answer_with_data_in_it_is_a_named_record()
        {
            string handlers = Path.Combine(ServerRouteTable.RepositoryRoot(), "src", "CRT.Server", "Handlers");
            var offenders = new List<string>();

            foreach (string file in Directory.EnumerateFiles(handlers, "*Endpoints.cs", SearchOption.AllDirectories))
            {
                string text = File.ReadAllText(file);

                foreach (Match match in Regex.Matches(text, @"new\s*\{(?<body>[^{}]*)\}"))
                {
                    string[] names = ApiCompatibilityTests.MemberNames(match.Groups["body"].Value);

                    if (names.Length == 0 || names.Any(name => name is not ("message" or "error" or "errors")))
                    {
                        int line = text.Take(match.Index).Count(character => character == '\n') + 1;
                        offenders.Add($"{Path.GetFileName(file)}:{line}  new {{ {string.Join(", ", names)} }}");
                    }
                }
            }

            Assert.True(
                offenders.Count == 0,
                "These answers are anonymous objects carrying data, which the API compatibility check cannot see - " +
                "make each a record in CRT.Data (ReviewApiContract.cs or SubmissionContract.cs):\n" + string.Join("\n", offenders));
        }

        [Theory]
        [InlineData(" message = outcome.Error ", new[] { "message" })]
        [InlineData(" error ", new[] { "error" })]
        [InlineData(" errors = outcome.Findings ", new[] { "errors" })]
        [InlineData(" message = string.Join(\", \", parts), resumeFrom = 3 ", new[] { "message", "resumeFrom" })]
        [InlineData("\n // a comment, with a comma\n message = \"x\" ", new[] { "message" })]
        [InlineData(" outcome.Answer ", new[] { "Answer" })]
        public void The_member_names_of_an_anonymous_object_are_read(string body, string[] expected)
        {
            Assert.Equal(expected, ApiCompatibilityTests.MemberNames(body));
        }

        // `name = value` gives name; a bare `name` or `x.name` (a projection) gives name. Commas inside
        // brackets and quotes do not split, and // comments are dropped.
        private static string[] MemberNames(string body)
        {
            string code = Regex.Replace(body, @"//[^\n]*", string.Empty);
            code = Regex.Replace(code, "\"(?:[^\"\\\\]|\\\\.)*\"", "\"\"");

            var parts = new List<string>();
            int depth = 0;
            int start = 0;

            for (int index = 0; index < code.Length; index++)
            {
                char character = code[index];

                if (character is '(' or '[')
                    depth++;
                else if (character is ')' or ']')
                    depth--;
                else if (character == ',' && depth == 0)
                {
                    parts.Add(code[start..index]);
                    start = index + 1;
                }
            }

            parts.Add(code[start..]);

            return parts
                .Select(part => part.Trim())
                .Where(part => part.Length > 0)
                .Select(part =>
                {
                    Match assigned = Regex.Match(part, @"^(\w+)\s*=(?!=)");

                    if (assigned.Success)
                        return assigned.Groups[1].Value;

                    Match projected = Regex.Match(part, @"(\w+)$");
                    return projected.Success ? projected.Groups[1].Value : part;
                })
                .ToArray();
        }

        // -----------------------------------------------------------------------------------

        private static IEnumerable<string> FrozenFiles() =>
            Directory.EnumerateFiles(ApiCompatibilityTests.Folder(), "crt-*.txt")
                .OrderBy(file => file, StringComparer.Ordinal);

        // ###########################################################################################
        // The version CRT.App.csproj builds - its InformationalVersion, the one place a release's
        // version is entered (CLAUDE.md, "Release process"). An InformationalVersion that is
        // "$(AssemblyVersion)" means the AssemblyVersion element's.
        // ###########################################################################################
        private static string CrtVersionBeingBuilt()
        {
            XDocument project = XDocument.Load(
                Path.Combine(ServerRouteTable.RepositoryRoot(), "src", "CRT.App", "CRT.App.csproj"));

            string? informational = project.Descendants("InformationalVersion").Select(element => element.Value.Trim()).LastOrDefault();
            string? assembly = project.Descendants("AssemblyVersion").Select(element => element.Value.Trim()).LastOrDefault();

            string? version = informational is null || informational.Contains("$(", StringComparison.Ordinal)
                ? assembly
                : informational;

            Assert.False(string.IsNullOrWhiteSpace(version), "CRT.App.csproj declares no version.");
            return version!;
        }
    }
}
