using Handlers.DataHandling;
using System;
using System.Globalization;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.Oscilloscope
{
    // What one Try it came to: whether the scope took the command, and what to say beside it.
    public sealed record ScopeTrialOutcome(bool Worked, string Reply);

    // ###########################################################################################
    // "TRY IT" - ONE COMMAND OF A MODEL BEING EDITED, SENT TO THE SCOPE CONNECTED (CRT 4.0 step 5 -
    // the plan: "edit each command with Try it ... that shows the scope's reply"). It takes the
    // model as the editor holds it, so the command tried is the one typed, and says in plain words
    // what came back:
    //
    //   - a QUERY: the scope's answer (the serial number in an Identify answer masked, as the
    //     output pane masks it); a query for a number says the number as CRT read it, or that it
    //     is none CRT can read - the keyboard steps would do nothing with it;
    //   - a command that ANSWERS NOTHING (Stop, Single, Run, Clear statistics): sent, then the
    //     model's own "Finished" and "Last error" asked, as the palettes do - an error the scope
    //     kept (an answer starting with a number other than 0, SCPI's way) means it did not take it;
    //   - a SET command: the value is the one the scope HAS, read with the model's matching Read
    //     command first and set back - so trying it changes nothing on the scope, and works only
    //     once the Read command does;
    //   - the SCREEN PICTURE: how much came back, and whether it is a picture CRT can read.
    //
    // A scope that does not answer in time throws, as every scope call does: the session's own
    // handling drops the connection, since nobody knows what the scope will send next. The tab
    // says so beside the command.
    // ###########################################################################################
    public static class ScopeCommandTrial
    {
        public const string NoCommand = "Type the command first.";

        public static async Task<ScopeTrialOutcome> TryAsync(
            IScopeClient client,
            MainDocument.OscilloscopeRow model,
            ScopeCommand command,
            Action<string> log,
            CancellationToken cancellationToken)
        {
            ArgumentNullException.ThrowIfNull(client);
            ArgumentNullException.ThrowIfNull(model);
            ArgumentNullException.ThrowIfNull(log);

            OscilloscopeEntry entry = OscilloscopeEntry.From(model);
            string text = ScopeCommandResolver.GetCommandText(entry, command).Trim();

            if (text.Length == 0)
            {
                return new ScopeTrialOutcome(false, ScopeCommandTrial.NoCommand);
            }

            switch (command)
            {
                case ScopeCommand.SetTriggerLevel:
                    return await ScopeCommandTrial.TrySetAsync(client, entry, text, ScopeCommand.QueryTriggerLevel, ScopeFormatting.FormatVoltage, log, cancellationToken).ConfigureAwait(false);

                case ScopeCommand.SetTimeDiv:
                    return await ScopeCommandTrial.TrySetAsync(client, entry, text, ScopeCommand.QueryTimeDiv, ScopeFormatting.FormatTime, log, cancellationToken).ConfigureAwait(false);

                case ScopeCommand.SetVoltsDiv:
                    return await ScopeCommandTrial.TrySetAsync(client, entry, text, ScopeCommand.QueryVoltsDiv, ScopeFormatting.FormatVoltage, log, cancellationToken).ConfigureAwait(false);

                case ScopeCommand.DumpImage:
                    log($"SCPI >> {text}");
                    byte[] raw = await client.QueryBinaryBlockAsync(text, cancellationToken).ConfigureAwait(false);
                    log($"SCPI << <{raw.Length} bytes binary>");
                    return ScopeCommandTrial.DescribePicture(raw);
            }

            if (!ScopeCommandResolver.ExpectsTextResponse(command))
            {
                log($"SCPI >> {text}");
                await client.SendAsync(text, cancellationToken).ConfigureAwait(false);

                return await ScopeCommandTrial.AfterSendingAsync(client, entry, "Sent.", log, cancellationToken).ConfigureAwait(false);
            }

            string reply = await ScopeCommandTrial.QueryAsync(client, text, command, log, cancellationToken).ConfigureAwait(false);

            if (reply.Length == 0)
            {
                return new ScopeTrialOutcome(false, "The scope answered with an empty line.");
            }

            switch (command)
            {
                case ScopeCommand.QueryTriggerLevel:
                case ScopeCommand.QueryVoltsDiv:
                    return ScopeCommandTrial.ReadNumber(reply, ScopeFormatting.FormatVoltage);

                case ScopeCommand.QueryTimeDiv:
                    return ScopeCommandTrial.ReadNumber(reply, ScopeFormatting.FormatTime);

                case ScopeCommand.DrainErrorQueue:
                    return new ScopeTrialOutcome(true, ScopeCommandTrial.IsAnError(reply)
                        ? $"{reply} - an error the scope kept from before"
                        : reply);

                case ScopeCommand.Identify:
                    return new ScopeTrialOutcome(true, ScopeFormatting.MaskIdentifyResponseSerial(reply));

                default:
                    return new ScopeTrialOutcome(true, reply);
            }
        }

        // ###########################################################################################
        // Whether an answer of "Last error" names an error: SCPI's error queue answers a number, then
        // a comma and the words - 0 (or +0) when there is none.
        // ###########################################################################################
        public static bool IsAnError(string reply)
        {
            string head = (reply ?? string.Empty).Split(',')[0].Trim();

            return int.TryParse(head, NumberStyles.AllowLeadingSign, CultureInfo.InvariantCulture, out int code) && code != 0;
        }

        private static async Task<ScopeTrialOutcome> TrySetAsync(
            IScopeClient client,
            OscilloscopeEntry entry,
            string text,
            ScopeCommand readCommand,
            Func<double, string> format,
            Action<string> log,
            CancellationToken cancellationToken)
        {
            string readText = ScopeCommandResolver.GetCommandText(entry, readCommand).Trim();
            string readLabel = ScopeModelFields.Of(readCommand).Label;

            if (readText.Length == 0)
            {
                return new ScopeTrialOutcome(false, $"Fill in \"{readLabel}\" first - CRT sets the value the scope already has, so trying this changes nothing.");
            }

            string reply = await ScopeCommandTrial.QueryAsync(client, readText, readCommand, log, cancellationToken).ConfigureAwait(false);

            if (!double.TryParse(reply, NumberStyles.Float, CultureInfo.InvariantCulture, out double value))
            {
                return new ScopeTrialOutcome(false, $"\"{readLabel}\" answered \"{reply}\", which is no number CRT can read - make that one work first.");
            }

            string set = ScopeFormatting.FillValue(text, value);

            log($"SCPI >> {set}");
            await client.SendAsync(set, cancellationToken).ConfigureAwait(false);

            return await ScopeCommandTrial.AfterSendingAsync(
                client,
                entry,
                $"Sent {set} - {format(value)}, the value the scope had.",
                log,
                cancellationToken).ConfigureAwait(false);
        }

        // The model's "Finished" and "Last error" after a command that answers nothing - each only
        // when the model has it.
        private static async Task<ScopeTrialOutcome> AfterSendingAsync(
            IScopeClient client,
            OscilloscopeEntry entry,
            string sent,
            Action<string> log,
            CancellationToken cancellationToken)
        {
            string finishedText = entry.OperationComplete.Trim();
            string errorText = entry.DrainErrorQueue.Trim();
            string said = sent;

            if (finishedText.Length > 0)
            {
                string finished = await ScopeCommandTrial.QueryAsync(client, finishedText, ScopeCommand.OperationComplete, log, cancellationToken).ConfigureAwait(false);
                said += $" Finished: {finished}.";
            }

            if (errorText.Length == 0)
            {
                return new ScopeTrialOutcome(true, said);
            }

            string error = await ScopeCommandTrial.QueryAsync(client, errorText, ScopeCommand.DrainErrorQueue, log, cancellationToken).ConfigureAwait(false);

            return ScopeCommandTrial.IsAnError(error)
                ? new ScopeTrialOutcome(false, $"{said} But the scope says: {error}")
                : new ScopeTrialOutcome(true, $"{said} Last error: {error}");
        }

        private static async Task<string> QueryAsync(IScopeClient client, string text, ScopeCommand command, Action<string> log, CancellationToken cancellationToken)
        {
            log($"SCPI >> {text}");

            string reply = (await client.QueryLineAsync(text, cancellationToken).ConfigureAwait(false) ?? string.Empty).Trim();

            log($"SCPI << {(command == ScopeCommand.Identify ? ScopeFormatting.MaskIdentifyResponseSerial(reply) : reply)}");

            return reply;
        }

        private static ScopeTrialOutcome ReadNumber(string reply, Func<double, string> format) =>
            double.TryParse(reply, NumberStyles.Float, CultureInfo.InvariantCulture, out double value)
                ? new ScopeTrialOutcome(true, $"{reply} - read as {format(value)}")
                : new ScopeTrialOutcome(false, $"The scope answered \"{reply}\", which is no number CRT can read.");

        // ###########################################################################################
        // What the screen picture's answer is: its payload (the binary block's header taken off) a
        // BMP, PNG or JPEG - the pictures the capture decodes - with its size where the header says it.
        // The size is read by CRT's one reader of picture headers (ImagePixelSize.FromBytes - code
        // review, 2026-10-29: read here by hand, a JPEG's size was never said).
        // ###########################################################################################
        internal static ScopeTrialOutcome DescribePicture(byte[] raw)
        {
            byte[] payload = ScopePayloadParser.TryExtractBinaryPayload(raw, out byte[] extracted) && extracted.Length > 0 ? extracted : raw;
            string kilobytes = (payload.Length / 1024d).ToString("0.0", CultureInfo.InvariantCulture);

            string? kind = ScopePayloadParser.TryReadBmpMetadata(payload, out _, out _, out _) ? "BMP"
                : WorkbookPdfExporter.IsPngSignature(payload) ? "PNG"
                : payload.Length >= 3 && payload[0] == 0xFF && payload[1] == 0xD8 && payload[2] == 0xFF ? "JPEG"
                : null;

            if (kind is null)
            {
                return new ScopeTrialOutcome(false, $"{kilobytes} KB came back, but not as a picture CRT can read (BMP, PNG or JPEG).");
            }

            return new ScopeTrialOutcome(true, ImagePixelSize.FromBytes(payload) is { } size
                ? $"A {kind} picture of {size.Width}x{size.Height} pixels came back ({kilobytes} KB)."
                : $"A {kind} picture came back ({kilobytes} KB).");
        }
    }
}
