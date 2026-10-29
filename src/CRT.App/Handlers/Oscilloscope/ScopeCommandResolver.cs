using Handlers.DataHandling;
using System;

namespace Handlers.Oscilloscope
{
    public static class ScopeCommandResolver
    {
        // ###########################################################################################
        // Resolves a scope command enum to the SCPI command text defined by the selected scope entry:
        // the value its editor field names (ScopeModelFields - the command's key, the main file's name
        // for it), read off the entry by that name. No switch of its own (code review, 2026-10-29).
        // ###########################################################################################
        public static string GetCommandText(OscilloscopeEntry oscilloscopeEntry, ScopeCommand command)
        {
            ArgumentNullException.ThrowIfNull(oscilloscopeEntry);

            return oscilloscopeEntry.ValueOf(ScopeModelFields.Of(command).Key);
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