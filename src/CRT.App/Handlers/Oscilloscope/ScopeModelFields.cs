using Handlers.DataHandling;
using System;
using System.Collections.Generic;
using System.Globalization;
using System.Linq;

namespace Handlers.Oscilloscope
{
    public enum ScopeModelFieldKind
    {
        Setting,
        Command,
        StepList,
    }

    // One value of an oscilloscope model the editor shows: its key (the main file's JSON name for
    // it), its words, what it is, and where it lives in the main file's row.
    public sealed record ScopeModelField(
        string Key,
        string Label,
        string Help,
        ScopeModelFieldKind Kind,
        ScopeCommand? Command,
        Func<MainDocument.OscilloscopeRow, string> Get,
        Action<MainDocument.OscilloscopeRow, string> Set);

    // ###########################################################################################
    // EVERY VALUE OF AN OSCILLOSCOPE MODEL THE OSCILLOSCOPE TAB'S EDITOR SHOWS, in the order shown
    // (CRT 4.0 step 5): the two settings, the sixteen commands - one per ScopeCommand, each with
    // "Try it" - and the two lists of steps the numpad keys move through. The words are the ones a
    // hobbyist with the scope's programming manual open needs; the Wiki's Main Excel page said the
    // same of each column. Every value of the main file's oscilloscope entry is here but the vendor
    // and series or model, which the editor's name boxes hold (ScopeModelFieldsTests counts them).
    // ###########################################################################################
    public static class ScopeModelFields
    {
        public const string PortKey = "port";

        public const string WaitKey = "debounceTime";

        public const string TimeStepsKey = "timeDivList";

        public const string VoltsStepsKey = "voltsDivList";

        public const string PortProblem = "A network port is a number from 1 to 65535.";

        public const string WaitProblem = "Write a time with its unit, such as 100ms - or leave it empty for 250ms.";

        public const string NoSteps = "Empty, so the numpad keys cannot step - fill in every step your scope offers.";

        public static readonly IReadOnlyList<ScopeModelField> All =
        [
            ScopeModelFields.Field(PortKey, ScopeModelFieldKind.Setting, "Network port",
                "The port your scope takes SCPI commands on - 5025 for most scopes, 5555 for Rigol."),
            ScopeModelFields.Field(WaitKey, ScopeModelFieldKind.Setting, "Wait between commands",
                "How long CRT waits between commands, with its unit, such as 100ms. Empty is 250ms."),

            ScopeModelFields.CommandField("identify", ScopeCommand.Identify, "Identify",
                "Asks the scope who it is. It answers its vendor, model, serial number and firmware. Nearly always *IDN?"),
            ScopeModelFields.CommandField("drainErrorQueue", ScopeCommand.DrainErrorQueue, "Last error",
                "Asks for the next error the scope has kept, such as :SYSTem:ERRor? - an answer starting with 0 means no error."),
            ScopeModelFields.CommandField("operationComplete", ScopeCommand.OperationComplete, "Finished",
                "Asks whether the scope has finished the command before. Nearly always *OPC?"),
            ScopeModelFields.CommandField("clearStatistics", ScopeCommand.ClearStatistics, "Clear statistics",
                "Clears the scope's measurement statistics."),
            ScopeModelFields.CommandField("queryActiveTrigger", ScopeCommand.QueryActiveTrigger, "Trigger state",
                "Asks whether the scope is running, waiting or has triggered."),
            ScopeModelFields.CommandField("stop", ScopeCommand.Stop, "Stop",
                "Freezes what the screen shows, now."),
            ScopeModelFields.CommandField("single", ScopeCommand.Single, "Single",
                "Freezes the screen at the next trigger."),
            ScopeModelFields.CommandField("run", ScopeCommand.Run, "Run",
                "Starts capturing again."),
            ScopeModelFields.CommandField("queryTriggerMode", ScopeCommand.QueryTriggerMode, "Trigger mode",
                "Asks the trigger mode, such as Edge."),
            ScopeModelFields.CommandField("queryTriggerLevel", ScopeCommand.QueryTriggerLevel, "Read trigger level",
                "Asks the trigger level. The scope answers a number of volts."),
            ScopeModelFields.CommandField("setTriggerLevel", ScopeCommand.SetTriggerLevel, "Set trigger level",
                "Sets the trigger level. Write {0} where the volts go - without it CRT adds the value at the end."),
            ScopeModelFields.CommandField("queryTimeDiv", ScopeCommand.QueryTimeDiv, "Read time per division",
                "Asks the time base. The scope answers a number of seconds per division."),
            ScopeModelFields.CommandField("setTimeDiv", ScopeCommand.SetTimeDiv, "Set time per division",
                "Sets the time base. Write {0} where the seconds go - without it CRT adds the value at the end."),
            ScopeModelFields.CommandField("queryVoltsDiv", ScopeCommand.QueryVoltsDiv, "Read volts per division",
                "Asks the volts per division of channel 1. The scope answers a number of volts."),
            ScopeModelFields.CommandField("setVoltsDiv", ScopeCommand.SetVoltsDiv, "Set volts per division",
                "Sets the volts per division. Write {0} where the volts go - without it CRT adds the value at the end."),
            ScopeModelFields.CommandField("dumpImage", ScopeCommand.DumpImage, "Screen picture",
                "Fetches what the screen shows, as a picture. This differs the most between vendors."),

            ScopeModelFields.Field(TimeStepsKey, ScopeModelFieldKind.StepList, "Time per division steps",
                "Every time per division your scope offers, separated by commas, such as 1us,2us,5us. The numpad keys step through them."),
            ScopeModelFields.Field(VoltsStepsKey, ScopeModelFieldKind.StepList, "Volts per division steps",
                "Every volts per division your scope offers, separated by commas, such as 500mV,1V,2V. The numpad keys step through them."),
        ];

        private static readonly IReadOnlyDictionary<ScopeCommand, ScopeModelField> ByCommand =
            ScopeModelFields.All.Where(field => field.Command is not null).ToDictionary(field => field.Command!.Value);

        // The command field of `command`.
        public static ScopeModelField Of(ScopeCommand command) =>
            ScopeModelFields.ByCommand.TryGetValue(command, out ScopeModelField? field)
                ? field
                : throw new ArgumentOutOfRangeException(nameof(command), command, "Unknown scope command");

        // ###########################################################################################
        // What is wrong with `value` for `field` - null when nothing. A setting that is wrong is not
        // saved (the tab reads it as a number or a time); a step list is saved anyway and its
        // problem said: the steps CRT can read still work, and the user may be half way through.
        // A command is the scope's business - Try it says whether it works.
        // ###########################################################################################
        public static string? ProblemOf(ScopeModelField field, string? value)
        {
            ArgumentNullException.ThrowIfNull(field);

            string text = value?.Trim() ?? string.Empty;

            switch (field.Key)
            {
                case PortKey:
                    return int.TryParse(text, NumberStyles.None, CultureInfo.InvariantCulture, out int port) && port is >= 1 and <= 65535
                        ? null
                        : ScopeModelFields.PortProblem;

                case WaitKey:
                    return text.Length == 0 || ScopeValueMapper.TryParseTimeValue(text, out _) ? null : ScopeModelFields.WaitProblem;

                case TimeStepsKey:
                case VoltsStepsKey:
                    if (text.Length == 0)
                    {
                        return ScopeModelFields.NoSteps;
                    }

                    bool isTime = field.Key == TimeStepsKey;

                    string[] unreadable = text
                        .Split(',', StringSplitOptions.RemoveEmptyEntries | StringSplitOptions.TrimEntries)
                        .Where(step => isTime ? !ScopeValueMapper.TryParseTimeValue(step, out _) : !ScopeValueMapper.TryParseVoltageValue(step, out _))
                        .ToArray();

                    return unreadable.Length == 0
                        ? null
                        : $"CRT cannot read {string.Join(", ", unreadable.Select(step => $"\"{step}\""))} - write each step with its unit, such as {(isTime ? "5us" : "500mV")}.";

                default:
                    return null;
            }
        }

        // Whether a value wrong for `field` is kept back rather than saved (see ProblemOf).
        public static bool IsHeldBackWhenWrong(ScopeModelField field) => field.Kind == ScopeModelFieldKind.Setting;

        private static ScopeModelField CommandField(string key, ScopeCommand command, string label, string help) =>
            ScopeModelFields.Field(key, ScopeModelFieldKind.Command, label, help, command);

        // ###########################################################################################
        // A field read and written by its KEY - the main file's name for the value - through the one
        // map of a model's values (OscilloscopeModelValues; code review, 2026-10-29: a switch of its
        // own here was the third copy). The words are this table's; where the value lives is not.
        // ###########################################################################################
        private static ScopeModelField Field(string key, ScopeModelFieldKind kind, string label, string help, ScopeCommand? command = null) => new(
            key,
            label,
            help,
            kind,
            command,
            row => OscilloscopeModelValues.Get(row, key),
            (row, value) => OscilloscopeModelValues.Set(row, key, value));
    }
}
