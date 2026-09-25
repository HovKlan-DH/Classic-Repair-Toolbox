using Handlers.DataHandling;
using System;

namespace Handlers.Oscilloscope
{
    public static class ScopeCommandResolver
    {
        // ###########################################################################################
        // Resolves a scope command enum to the SCPI command text defined by the selected scope entry.
        // ###########################################################################################
        public static string GetCommandText(OscilloscopeEntry oscilloscopeEntry, ScopeCommand command)
        {
            return command switch
            {
                ScopeCommand.Identify => oscilloscopeEntry.Identify,
                ScopeCommand.DrainErrorQueue => oscilloscopeEntry.DrainErrorQueue,
                ScopeCommand.OperationComplete => oscilloscopeEntry.OperationComplete,
                ScopeCommand.ClearStatistics => oscilloscopeEntry.ClearStatistics,
                ScopeCommand.QueryActiveTrigger => oscilloscopeEntry.QueryActiveTrigger,
                ScopeCommand.Stop => oscilloscopeEntry.Stop,
                ScopeCommand.Single => oscilloscopeEntry.Single,
                ScopeCommand.Run => oscilloscopeEntry.Run,
                ScopeCommand.QueryTriggerMode => oscilloscopeEntry.QueryTriggerMode,
                ScopeCommand.QueryTriggerLevel => oscilloscopeEntry.QueryTriggerLevel,
                ScopeCommand.SetTriggerLevel => oscilloscopeEntry.SetTriggerLevel,
                ScopeCommand.QueryTimeDiv => oscilloscopeEntry.QueryTimeDiv,
                ScopeCommand.SetTimeDiv => oscilloscopeEntry.SetTimeDiv,
                ScopeCommand.QueryVoltsDiv => oscilloscopeEntry.QueryVoltsDiv,
                ScopeCommand.SetVoltsDiv => oscilloscopeEntry.SetVoltsDiv,
                ScopeCommand.DumpImage => oscilloscopeEntry.DumpImage,
                _ => throw new ArgumentOutOfRangeException(nameof(command), command, "Unknown scope command")
            };
        }

        // ###########################################################################################
        // Returns true when the scope command is expected to return a text response.
        // ###########################################################################################
        public static bool ExpectsTextResponse(ScopeCommand command)
        {
            return command == ScopeCommand.Identify ||
                   command == ScopeCommand.DrainErrorQueue ||
                   command == ScopeCommand.OperationComplete ||
                   command == ScopeCommand.QueryActiveTrigger ||
                   command == ScopeCommand.QueryTriggerMode ||
                   command == ScopeCommand.QueryTriggerLevel ||
                   command == ScopeCommand.QueryTimeDiv ||
                   command == ScopeCommand.QueryVoltsDiv;
        }
    }
}