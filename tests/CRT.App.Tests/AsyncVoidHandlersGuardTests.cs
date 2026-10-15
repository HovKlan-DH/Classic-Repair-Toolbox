using System.Text.RegularExpressions;

namespace ClassicRepairToolbox.Tests;

// ###########################################################################################
// *** AN "async void" METHOD IS GUARDED AS A WHOLE (code review, 2026-10-13). ***
//
// An async void handler has no caller: an exception after its first await is rethrown on the
// dispatcher, and App's handler lets that crash go on, on purpose. A thumbnail click, a worklog
// pill or a Configuration checkbox could end CRT that way. Each one's whole body is now a try with
// a catch (Exception ...) that logs and keeps CRT running - Main.OnBoardSelectionChanged's pattern.
//
// This reads every such method in the areas that review covered and fails on one that awaits
// anywhere but inside a top-level try with a catch-all (comments and strings set aside). An
// expression-bodied one must call a method that guards itself, named below with the reason.
// ###########################################################################################
public sealed class AsyncVoidHandlersGuardTests
{
    private static readonly string[] Folders =
    [
        Path.Combine("Main"),
        Path.Combine("Tabs", "Schematics"),
        Path.Combine("Tabs", "Oscilloscope"),
        Path.Combine("Tabs", "Workbooks"),
        Path.Combine("Tabs", "Worklog"),
        Path.Combine("Tabs", "Overview"),
        Path.Combine("Tabs", "Resources"),
        Path.Combine("Tabs", "Configuration"),
        Path.Combine("Tabs", "Feedback"),
        Path.Combine("Tabs", "About"),

        // The contribution side and the shared controls (code review, 2026-10-13).
        Path.Combine("Tabs", "Drafts"),
        Path.Combine("Tabs", "Contribute"),
        Path.Combine("Controls"),
    ];

    // The Maintainer tab guards its handlers through HandlerGuard instead (HandlerGuardTests), and
    // Main.Maintainer.cs is its part of Main.
    private static readonly string[] FilesNotCovered =
    [
        "Main.Maintainer.cs",
    ];

    // Expression-bodied handlers, each calling a method that is itself guarded as a whole.
    private static readonly Dictionary<string, string> ExpressionBodiedAllowed = new()
    {
        ["OnInstallUpdateClick"] = "InstallPendingUpdateAsync is one try with a catch-all",
    };

    private static readonly Regex CommentOrLiteral = new(
        @"//[^\n]*|/\*.*?\*/|@""(?:""""|[^""])*""|""(?:\\.|[^""\\\n])*""|'(?:\\.|[^'\\\n])+'",
        RegexOptions.Singleline | RegexOptions.CultureInvariant);

    private static readonly Regex AsyncVoid = new(
        @"async\s+void\s+(\w+)\s*\([^)]*\)\s*(=>|\{)",
        RegexOptions.CultureInvariant);

    // ###########################################################################################
    // The async void methods in `source` that are not guarded as a whole, by name.
    // ###########################################################################################
    internal static IReadOnlyList<string> UnguardedIn(string source)
    {
        string code = CommentOrLiteral.Replace(source, match => new string(' ', match.Length));
        var unguarded = new List<string>();

        foreach (Match method in AsyncVoid.Matches(code))
        {
            string name = method.Groups[1].Value;

            if (method.Groups[2].Value == "=>")
            {
                if (!ExpressionBodiedAllowed.ContainsKey(name))
                    unguarded.Add(name);

                continue;
            }

            int open = method.Index + method.Length - 1;
            int close = MatchingBrace(code, open);

            if (!IsGuarded(code[(open + 1)..close]))
                unguarded.Add(name);
        }

        return unguarded;
    }

    // ###########################################################################################
    // Whether a method body awaits only inside top-level try blocks that catch everything. Plain
    // statements before the first await (an early "return" on a flag) may sit outside: what
    // reaches App's handler is what is thrown after an await, where the method has no caller.
    // ###########################################################################################
    private static bool IsGuarded(string body)
    {
        var outside = new System.Text.StringBuilder();
        bool sawGuardedTry = false;
        int position = 0;

        while (position < body.Length)
        {
            var nextTry = Regex.Match(body[position..], @"(?<![\w.])try\s*\{");

            if (!nextTry.Success)
            {
                outside.Append(body[position..]);
                break;
            }

            int tryStart = position + nextTry.Index;
            string before = body[position..tryStart];

            // A try nested in another block is the body's own business, not the guard.
            if (Depth(outside.ToString() + before) != 0)
            {
                outside.Append(body[position..(tryStart + nextTry.Length)]);
                position = tryStart + nextTry.Length;
                continue;
            }

            outside.Append(before);

            int tryClose = MatchingBrace(body, tryStart + nextTry.Length - 1);
            int after = tryClose + 1;
            bool catchesAll = false;

            while (true)
            {
                var clause = Regex.Match(body[after..], @"^\s*(catch|finally)\b");

                if (!clause.Success)
                    break;

                catchesAll |= clause.Groups[1].Value == "catch" &&
                              Regex.IsMatch(body[(after + clause.Length)..], @"^\s*(\{|\(\s*(System\.)?Exception\b)");

                int clauseOpen = body.IndexOf('{', after + clause.Length);
                after = MatchingBrace(body, clauseOpen) + 1;
            }

            if (!catchesAll)
                return false;

            sawGuardedTry = true;
            position = after;
        }

        return sawGuardedTry && !Regex.IsMatch(outside.ToString(), @"\bawait\b");
    }

    private static int Depth(string text) => text.Count(c => c == '{') - text.Count(c => c == '}');

    private static int MatchingBrace(string text, int open)
    {
        int depth = 0;

        for (int i = open; i < text.Length; i++)
        {
            if (text[i] == '{')
                depth++;
            else if (text[i] == '}' && --depth == 0)
                return i;
        }

        throw new InvalidOperationException("Unbalanced braces");
    }

    [Fact]
    public void Every_async_void_method_in_the_covered_areas_is_one_try_with_a_catch_all()
    {
        string app = ResolveAppFolder();

        var files = Folders
            .Select(folder => Path.Combine(app, folder))
            .Where(Directory.Exists)
            .SelectMany(folder => Directory.GetFiles(folder, "*.cs", SearchOption.AllDirectories))
            .Where(file => !FilesNotCovered.Contains(Path.GetFileName(file)))
            .ToList();

        Assert.NotEmpty(files);

        var offences = files
            .SelectMany(file => UnguardedIn(File.ReadAllText(file))
                .Select(name => $"{Path.GetRelativePath(app, file)}: {name}"))
            .ToList();

        Assert.True(offences.Count == 0,
            "An async void method must be guarded as a whole (try { ... } catch (Exception ex) { log }):" +
            Environment.NewLine + string.Join(Environment.NewLine, offences));
    }

    // The scanner itself, on small pieces of C#.
    [Theory]
    [InlineData("async void A() { try { await X(); } catch (Exception ex) { Log(ex); } }", true)]
    [InlineData("async void A() { try { await X(); } catch (OperationCanceledException) { } catch (Exception) { } finally { Y(); } }", true)]
    [InlineData("async void A() { try { await X(); } catch { } }", true)]
    [InlineData("async void A() { await X(); }", false)]
    [InlineData("async void A() { try { await X(); } catch (OperationCanceledException) { } }", false)]
    [InlineData("async void A() { try { await X(); } catch (Exception) { } await Y(); }", false)]
    [InlineData("async void A() { if (b) return; try { await X(); } catch (Exception) { } }", true)]
    [InlineData("async void A() { await Y(); try { await X(); } catch (Exception) { } }", false)]
    [InlineData("async void A() { if (b) { try { await X(); } catch (Exception) { } } }", false)]
    [InlineData("async void A() => await X();", false)]
    [InlineData("async void OnInstallUpdateClick() => await X();", true)]
    [InlineData("async void A() { // await here\n try { await X(); } catch (Exception) { Log(\"} oops {\"); } }", true)]
    public void The_scanner_tells_a_guarded_body_from_an_unguarded_one(string source, bool guarded)
    {
        Assert.Equal(guarded, UnguardedIn(source).Count == 0);
    }

    private static string ResolveAppFolder()
    {
        string relative = Path.Combine("src", "CRT.App");
        var directory = new DirectoryInfo(AppContext.BaseDirectory);

        while (directory is not null && !Directory.Exists(Path.Combine(directory.FullName, relative, "Main")))
            directory = directory.Parent;

        Assert.NotNull(directory);

        return Path.Combine(directory.FullName, relative);
    }
}
