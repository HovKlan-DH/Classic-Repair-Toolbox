// Parses minipro's logic-test output: the per-pin grid on stdout (ported from the
// PLA Doctor parser) and the "Logic test successful/failed" summary + connection
// errors on stderr. A clean pass on a large run prints NO per-vector lines (only
// failures are listed), so a missing grid + a "successful" summary IS a pass.

using System;
using System.Collections.Generic;
using System.Text;
using System.Text.RegularExpressions;

namespace Handlers.IcTesting;

public static class MiniproOutputParser
{
    // Strip ANSI colour: ESC '[' ... 'm'. The ESC byte is required so we never
    // remove legitimate '[...]m' content or leave a stray ESC that shifts columns.
    private static readonly Regex Ansi = new(((char)0x1b) + "\\[[0-9;]*m", RegexOptions.Compiled);
    private static readonly Regex VectorLine = new(@"^\s*(\d+):\s?(.*)$", RegexOptions.Compiled);

    /// <summary>Strip ANSI colour codes so raw minipro output reads cleanly in a TextBox.</summary>
    public static string StripAnsi(string? s) => Ansi.Replace(s ?? "", "");

    /// <summary>Re-space the pin-number header line above the per-vector grid so it lines
    /// up with the data: minipro prints each pin's data in a fixed 3-char field (see
    /// VectorLine/failure-scan above) but writes the header numbers at their natural
    /// width, so pins 10+ drift out of alignment with their column. We already know the
    /// grid's true column layout from the first data row, so rebuild the header to match
    /// it instead of trying to fix minipro's own spacing.</summary>
    public static string AlignVectorTableHeader(string? s)
    {
        var text = s ?? "";
        var lines = text.Replace("\r\n", "\n").Split('\n');

        int dataIndex = -1;
        Match? dataMatch = null;
        for (int i = 0; i < lines.Length; i++)
        {
            var m = VectorLine.Match(lines[i]);
            if (m.Success) { dataIndex = i; dataMatch = m; break; }
        }
        if (dataMatch is null) return text;   // no per-vector grid printed (e.g. clean pass)

        int prefixLen = dataMatch.Groups[2].Index;

        int headerIndex = -1;
        List<int>? columns = null;
        for (int i = 0; i < dataIndex; i++)
        {
            var tokens = lines[i].Split((char[]?)null, StringSplitOptions.RemoveEmptyEntries);
            if (tokens.Length < 2) continue;
            var nums = new List<int>(tokens.Length);
            bool ok = true;
            foreach (var t in tokens)
            {
                if (!int.TryParse(t, out var n)) { ok = false; break; }
                nums.Add(n);
            }
            if (!ok) continue;
            bool increasing = true;
            for (int k = 1; k < nums.Count; k++)
                if (nums[k] <= nums[k - 1]) { increasing = false; break; }
            if (increasing && nums[0] >= 1 && nums[0] <= 3) { headerIndex = i; columns = nums; break; }
        }
        if (headerIndex < 0 || columns is null) return text;   // no recognisable header

        var sb = new StringBuilder();
        sb.Append(' ', prefixLen);
        foreach (var pin in columns)
            sb.Append(pin.ToString().PadRight(2)).Append(' ');
        lines[headerIndex] = sb.ToString();

        return string.Join('\n', lines);
    }

    public sealed record Parsed(
        bool? Passed,
        int VectorsSeen,
        IReadOnlyList<(int Vector, int Pin)> Failures,
        IReadOnlyList<int> FailingPins,
        MiniproConnectionState State,
        int? ErrorCount);

    public static Parsed Parse(string? stdout, string? stderr)
    {
        var failures = new List<(int, int)>();
        var failingPins = new SortedSet<int>();
        int vectorsSeen = 0;

        foreach (var raw in (stdout ?? "").Split('\n'))
        {
            var line = Ansi.Replace(raw, "").TrimEnd('\r');
            var m = VectorLine.Match(line);
            if (!m.Success) continue;                  // header / blank / non-vector line
            if (!int.TryParse(m.Groups[1].Value, out var vidx)) continue;
            vectorsSeen++;
            var body = m.Groups[2].Value;
            // Each pin field is 3 chars: symbol, marker, separator. marker '-' = error.
            for (int k = 0; k * 3 + 1 < body.Length; k++)
            {
                if (body[k * 3 + 1] == '-')
                {
                    int pin = k + 1;
                    failures.Add((vidx, pin));
                    failingPins.Add(pin);
                }
            }
        }

        var err = stderr ?? "";
        bool? passed = null;
        int? errCount = null;
        if (Regex.IsMatch(err, "logic test successful", RegexOptions.IgnoreCase))
            passed = true;
        var fm = Regex.Match(err, @"logic test failed:\s*(\d+)", RegexOptions.IgnoreCase);
        if (fm.Success) { passed = false; errCount = int.Parse(fm.Groups[1].Value); }
        else if (Regex.IsMatch(err, "logic test failed", RegexOptions.IgnoreCase)) { passed = false; }

        if (failures.Count > 0 && passed != false)
            passed = false;   // grid showed failures even if the summary was absent

        return new Parsed(passed, vectorsSeen, failures, new List<int>(failingPins),
            ClassifyState(err, passed), errCount);
    }

    public static MiniproConnectionState ClassifyState(string stderr, bool? passed)
    {
        var e = (stderr ?? "").ToLowerInvariant();
        if (e.Contains("no programmer") || e.Contains("programmer not found"))
            return MiniproConnectionState.NoProgrammer;
        if (e.Contains("overcurrent") || e.Contains("over-current") || e.Contains("over current"))
            return MiniproConnectionState.Overcurrent;
        if (e.Contains("device not found") || e.Contains("0x000000") || e.Contains("no device"))
            return MiniproConnectionState.NoChip;
        return passed.HasValue ? MiniproConnectionState.Ok : MiniproConnectionState.Unknown;
    }
}
