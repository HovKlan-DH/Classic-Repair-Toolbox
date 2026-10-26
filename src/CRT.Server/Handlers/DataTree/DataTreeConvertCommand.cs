using System.Globalization;
using Handlers.DataHandling;

namespace CRT.Server.Handlers.DataTree
{
    // ###########################################################################################
    // `CRT.Server convert-data-tree` - CONVERTS AN EXCEL-ERA DATA TREE INTO THE CRT 4.0 TREE
    // (CRT.Data's DataTreeConversion; owner decisions, 2026-10-24).
    //
    //   dotnet CRT.Server.dll convert-data-tree --source=<the Excel tree's Data folder>
    //                                           --target=<the 4.0 tree's folder>
    //                                           [--public-url=<the 4.0 tree's address, ending /Data>]
    //                                           [--plan-only]
    //
    // It runs INSTEAD of the service: no web host, no configuration, no database - so it can be
    // run on the server beside the running service, or on any PC with .NET against a copy of the
    // data. It only reads the source. It writes only the target, and only when the conversion
    // has no problem; `--plan-only` reports what it would do and writes nothing. Given the
    // public address, it writes the target's manifest too, which a tree CRTs sync from needs.
    //
    // Exit codes: 0 done, 1 refused (a problem, or a target it may not write - the report says
    // which), 2 the command line was wrong.
    // ###########################################################################################
    public static class DataTreeConvertCommand
    {
        public const string Name = "convert-data-tree";

        public const int Done = 0;
        public const int Refused = 1;
        public const int WrongCommandLine = 2;

        // How many left-out files the report names before it only counts the rest - enough for a
        // whole tree's (about 55 today), which is what is read before release day.
        private const int LeftOutNamed = 500;

        public const string UsageText =
            "Usage: CRT.Server convert-data-tree --source=<Excel tree's Data folder> --target=<4.0 tree folder> " +
            "[--public-url=<4.0 tree's address ending /Data>] [--plan-only]";

        public sealed record Options(string Source, string Target, string? PublicUrl, bool PlanOnly);

        public static bool IsCommand(IReadOnlyList<string> args) =>
            args.Count > 0 && string.Equals(args[0], DataTreeConvertCommand.Name, StringComparison.Ordinal);

        // ###########################################################################################
        // The options after the command's name, or null with the reason. Each value option is
        // given once, as --name=value (surrounding quotes stripped); anything else is refused, so
        // a mistyped option never quietly converts into the wrong place.
        // ###########################################################################################
        public static Options? Parse(IReadOnlyList<string> args, out string why)
        {
            var values = new Dictionary<string, string>(StringComparer.Ordinal);
            bool planOnly = false;

            foreach (string arg in args)
            {
                if (string.Equals(arg, "--plan-only", StringComparison.Ordinal))
                {
                    planOnly = true;
                    continue;
                }

                int equals = arg.IndexOf('=');
                string name = equals < 0 ? arg : arg[..equals];
                string value = equals < 0 ? string.Empty : arg[(equals + 1)..].Trim().Trim('"');

                if (name is not ("--source" or "--target" or "--public-url"))
                {
                    why = $"Unknown option [{arg}].";
                    return null;
                }

                if (values.ContainsKey(name))
                {
                    why = $"[{name}] is given twice.";
                    return null;
                }

                if (value.Length == 0)
                {
                    why = $"[{name}] needs a value: {name}=...";
                    return null;
                }

                values[name] = value;
            }

            if (!values.TryGetValue("--source", out string? source))
            {
                why = "--source is required: the Data folder of the tree to convert.";
                return null;
            }

            values.TryGetValue("--target", out string? target);
            values.TryGetValue("--public-url", out string? publicUrl);

            if (target is null && !planOnly)
            {
                why = "--target is required: the folder to write the converted tree into (or --plan-only).";
                return null;
            }

            if (publicUrl is not null &&
                (!Uri.TryCreate(publicUrl, UriKind.Absolute, out Uri? url) || url.Scheme is not ("https" or "http")))
            {
                why = $"--public-url [{publicUrl}] is not a web address.";
                return null;
            }

            why = string.Empty;
            return new Options(source, target ?? string.Empty, publicUrl, planOnly);
        }

        public static int Run(IReadOnlyList<string> args, TextWriter output)
        {
            Options? options = DataTreeConvertCommand.Parse(args.Skip(1).ToList(), out string why);

            if (options is null)
            {
                output.WriteLine(why);
                output.WriteLine(DataTreeConvertCommand.UsageText);
                return DataTreeConvertCommand.WrongCommandLine;
            }

            output.WriteLine($"Reading [{options.Source}] - every file it carries is hashed, which takes a while on a full tree.");

            DataTreeConversionPlan plan = DataTreeConversion.Plan(options.Source);

            foreach (string line in DataTreeConvertCommand.Report(plan))
                output.WriteLine(line);

            if (!plan.CanWrite)
            {
                output.WriteLine($"Nothing was written: {plan.Problems.Count} problem(s) above.");
                return DataTreeConvertCommand.Refused;
            }

            if (options.PlanOnly)
            {
                output.WriteLine("Planned only - nothing was written.");
                return DataTreeConvertCommand.Done;
            }

            try
            {
                DataTreeWriteResult result = DataTreeConversion.Write(plan, options.Target, options.PublicUrl);

                output.WriteLine(
                    $"Written into [{options.Target}]: {result.Written} file(s) written, {result.Kept} already right, {result.Removed} removed" +
                    (result.ManifestEntries >= 0 ? $", manifest of {result.ManifestEntries} file(s)." : ", no manifest (no --public-url)."));

                return DataTreeConvertCommand.Done;
            }
            catch (Exception ex) when (ex is InvalidOperationException or IOException or UnauthorizedAccessException)
            {
                output.WriteLine($"Nothing more was written: {ex.Message}");
                return DataTreeConvertCommand.Refused;
            }
        }

        // ###########################################################################################
        // What the plan holds, in lines a person reads: the generation and boards converted, the
        // files and their sizes, what is left out (named, up to LeftOutNamed), the notes and every
        // problem.
        // ###########################################################################################
        public static IReadOnlyList<string> Report(DataTreeConversionPlan plan)
        {
            var lines = new List<string>
            {
                $"Generation converted: {(plan.Generation is null ? "none found" : plan.Generation.ToString())}",
                $"Boards: {plan.Boards.Count}",
                $"Oscilloscopes: {plan.Main.Oscilloscopes.Count}",
                $"Files the boards name: {plan.Boards.Sum(board => board.Document.Files.Count)} ({DataTreeConvertCommand.Megabytes(plan.BoardFileBytes)})",
                $"Stored once: {plan.StoredFiles.Count} file(s) ({DataTreeConvertCommand.Megabytes(plan.StoredBytes)})",
                $"MiniPro test files: {plan.CopiedFiles.Count}",
            };

            if (plan.LeftOut.Count > 0)
            {
                lines.Add($"Left out - not used by the generation converted (older workbooks, what only they cite, !-notes): {plan.LeftOut.Count} file(s)");
                lines.AddRange(plan.LeftOut.Take(DataTreeConvertCommand.LeftOutNamed).Select(path => "  " + path));

                if (plan.LeftOut.Count > DataTreeConvertCommand.LeftOutNamed)
                    lines.Add($"  ... and {plan.LeftOut.Count - DataTreeConvertCommand.LeftOutNamed} more");
            }

            lines.AddRange(plan.Notes.Select(note => "Note: " + note));
            lines.AddRange(plan.Problems.Select(problem => "PROBLEM: " + problem));

            return lines;
        }

        private static string Megabytes(long bytes) =>
            (bytes / 1_000_000.0).ToString("0.0", CultureInfo.InvariantCulture) + " MB";
    }
}
