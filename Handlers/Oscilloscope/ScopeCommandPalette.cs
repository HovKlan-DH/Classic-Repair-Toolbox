using System;
using System.Collections.Generic;

namespace Handlers.Oscilloscope
{
    public enum ScopeCommand
    {
        Identify,
        DrainErrorQueue,
        OperationComplete,
        ClearStatistics,
        QueryActiveTrigger,
        Stop,
        Single,
        Run,
        QueryTriggerMode,
        QueryTriggerLevel,
        SetTriggerLevel,
        QueryTimeDiv,
        SetTimeDiv,
        QueryVoltsDiv,
        SetVoltsDiv,
        DumpImage
    }

    public enum ScopeCommandPalette
    {
        Identify,
        DrainErrorQueue,
        OperationComplete,
        ClearStatistics,
        QueryActiveTrigger,
        Stop,
        Single,
        Run,
        QueryTriggerMode,
        QueryTriggerLevel,
        SetTriggerLevel,
        QueryTimeDiv,
        SetTimeDiv,
        QueryVoltsDiv,
        SetVoltsDiv,
        DumpImage
    }

    public static class ScopeCommandPaletteDefinitions
    {
        private static readonly IReadOnlyDictionary<ScopeCommandPalette, ScopeCommand[]> CommandPalettes =
            new Dictionary<ScopeCommandPalette, ScopeCommand[]>
            {
                [ScopeCommandPalette.Identify] = new[]
                {
                    ScopeCommand.Identify
                },
                [ScopeCommandPalette.DrainErrorQueue] = new[]
                {
                    ScopeCommand.DrainErrorQueue
                },
                [ScopeCommandPalette.OperationComplete] = new[]
                {
                    ScopeCommand.OperationComplete
                },
                [ScopeCommandPalette.ClearStatistics] = new[]
                {
                    ScopeCommand.ClearStatistics,
                    ScopeCommand.OperationComplete,
                    ScopeCommand.DrainErrorQueue
                },
                [ScopeCommandPalette.QueryActiveTrigger] = new[]
                {
                    ScopeCommand.QueryActiveTrigger
                },
                [ScopeCommandPalette.Stop] = new[]
                {
                    ScopeCommand.Stop,
                    ScopeCommand.OperationComplete,
                    ScopeCommand.DrainErrorQueue
                },
                [ScopeCommandPalette.Single] = new[]
                {
                    ScopeCommand.Single,
                    ScopeCommand.OperationComplete,
                    ScopeCommand.DrainErrorQueue
                },
                [ScopeCommandPalette.Run] = new[]
                {
                    ScopeCommand.Run,
                    ScopeCommand.OperationComplete,
                    ScopeCommand.DrainErrorQueue
                },
                [ScopeCommandPalette.QueryTriggerMode] = new[]
                {
                    ScopeCommand.QueryTriggerMode
                },
                [ScopeCommandPalette.QueryTriggerLevel] = new[]
                {
                    ScopeCommand.QueryTriggerLevel
                },
                [ScopeCommandPalette.SetTriggerLevel] = new[]
                {
                    ScopeCommand.SetTriggerLevel,
                    ScopeCommand.OperationComplete,
                    ScopeCommand.DrainErrorQueue
                },
                [ScopeCommandPalette.QueryTimeDiv] = new[]
                {
                    ScopeCommand.QueryTimeDiv
                },
                [ScopeCommandPalette.SetTimeDiv] = new[]
                {
                    ScopeCommand.SetTimeDiv,
                    ScopeCommand.OperationComplete,
                    ScopeCommand.DrainErrorQueue
                },
                [ScopeCommandPalette.QueryVoltsDiv] = new[]
                {
                    ScopeCommand.QueryVoltsDiv
                },
                [ScopeCommandPalette.SetVoltsDiv] = new[]
                {
                    ScopeCommand.SetVoltsDiv,
                    ScopeCommand.OperationComplete,
                    ScopeCommand.DrainErrorQueue
                },
                [ScopeCommandPalette.DumpImage] = new[]
                {
                    ScopeCommand.DumpImage
                }
            };

        private static readonly IReadOnlyList<ScopeCommandPalette> FullCommandPaletteExecutionOrder =
            new[]
            {
                ScopeCommandPalette.Identify,
                ScopeCommandPalette.DrainErrorQueue,
                ScopeCommandPalette.OperationComplete,
                ScopeCommandPalette.ClearStatistics,
                ScopeCommandPalette.QueryActiveTrigger,
                ScopeCommandPalette.Stop,
                ScopeCommandPalette.Single,
                ScopeCommandPalette.Run,
                ScopeCommandPalette.QueryTriggerMode,
                ScopeCommandPalette.QueryTriggerLevel,
                ScopeCommandPalette.SetTriggerLevel,
                ScopeCommandPalette.QueryTimeDiv,
                ScopeCommandPalette.SetTimeDiv,
                ScopeCommandPalette.QueryVoltsDiv,
                ScopeCommandPalette.SetVoltsDiv,
                ScopeCommandPalette.DumpImage
            };

        // ###########################################################################################
        // Returns the ordered command list for a named scope command palette.
        // ###########################################################################################
        public static IReadOnlyList<ScopeCommand> GetCommands(ScopeCommandPalette palette)
        {
            if (CommandPalettes.TryGetValue(palette, out var commands))
            {
                return commands;
            }

            throw new ArgumentOutOfRangeException(nameof(palette), palette, "Unknown scope command palette");
        }

        // ###########################################################################################
        // Returns the standard execution order for the full oscilloscope command run.
        // ###########################################################################################
        public static IReadOnlyList<ScopeCommandPalette> GetFullCommandPaletteExecutionOrder()
        {
            return FullCommandPaletteExecutionOrder;
        }
    }
}