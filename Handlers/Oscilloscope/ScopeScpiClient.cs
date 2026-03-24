using System;
using System.Collections.Generic;
using System.Net.Sockets;
using System.Text;
using System.Threading;
using System.Threading.Tasks;

namespace Handlers.Oscilloscope
{
    public sealed class ScopeScpiClient : IAsyncDisposable
    {
        private readonly TcpClient thisClient = new();
        private NetworkStream? thisStream;

        // ###########################################################################################
        // Connects to the target scope over TCP and prepares the SCPI network stream.
        // ###########################################################################################
        public async Task ConnectAsync(string host, int port, CancellationToken cancellationToken)
        {
            await this.thisClient.ConnectAsync(host, port, cancellationToken);
            this.thisStream = this.thisClient.GetStream();
        }

        // ###########################################################################################
        // Sends a single SCPI command line to the connected scope.
        // ###########################################################################################
        public async Task SendAsync(string commandText, CancellationToken cancellationToken)
        {
            if (this.thisStream == null)
            {
                throw new InvalidOperationException("No active oscilloscope session exists.");
            }

            string normalizedCommand = commandText.EndsWith("\n", StringComparison.Ordinal)
                ? commandText
                : commandText + "\n";

            byte[] payload = Encoding.ASCII.GetBytes(normalizedCommand);
            await this.thisStream.WriteAsync(payload.AsMemory(0, payload.Length), cancellationToken);
            await this.thisStream.FlushAsync(cancellationToken);
        }

        // ###########################################################################################
        // Sends a SCPI query command and reads a single line response from the scope.
        // ###########################################################################################
        public async Task<string> QueryLineAsync(string commandText, CancellationToken cancellationToken)
        {
            await this.SendAsync(commandText, cancellationToken);
            return await this.ReadLineAsync(cancellationToken);
        }

        // ###########################################################################################
        // Sends a SCPI query command and reads a definite-length binary block response.
        // Returns the raw SCPI block including its header.
        // ###########################################################################################
        public async Task<byte[]> QueryBinaryBlockAsync(string commandText, CancellationToken cancellationToken)
        {
            if (this.thisStream == null)
            {
                throw new InvalidOperationException("No active oscilloscope session exists.");
            }

            await this.SendAsync(commandText, cancellationToken);

            byte[] hashByte = await this.ReadExactlyAsync(1, cancellationToken);
            if (hashByte[0] != (byte)'#')
            {
                throw new InvalidOperationException("Expected SCPI binary block response.");
            }

            byte[] lengthDigitsByte = await this.ReadExactlyAsync(1, cancellationToken);
            int lengthDigits = lengthDigitsByte[0] - (byte)'0';

            if (lengthDigits < 1 || lengthDigits > 9)
            {
                throw new InvalidOperationException("Invalid SCPI binary block length header.");
            }

            byte[] payloadLengthBytes = await this.ReadExactlyAsync(lengthDigits, cancellationToken);
            string payloadLengthText = Encoding.ASCII.GetString(payloadLengthBytes);

            if (!int.TryParse(payloadLengthText, out int payloadLength) || payloadLength < 0)
            {
                throw new InvalidOperationException("Invalid SCPI binary block payload length.");
            }

            byte[] payloadBytes = await this.ReadExactlyAsync(payloadLength, cancellationToken);

            var rawBytes = new byte[2 + lengthDigits + payloadLength];
            rawBytes[0] = hashByte[0];
            rawBytes[1] = lengthDigitsByte[0];
            Buffer.BlockCopy(payloadLengthBytes, 0, rawBytes, 2, lengthDigits);
            Buffer.BlockCopy(payloadBytes, 0, rawBytes, 2 + lengthDigits, payloadLength);

            await this.DiscardBinaryBlockTerminatorAsync(cancellationToken);

            return rawBytes;
        }

        // ###########################################################################################
        // Consumes the trailing CR/LF terminator that may follow a SCPI binary block response.
        // ###########################################################################################
        private async Task DiscardBinaryBlockTerminatorAsync(CancellationToken cancellationToken)
        {
            if (this.thisStream == null)
            {
                throw new InvalidOperationException("No active oscilloscope session exists.");
            }

            if (!this.thisStream.DataAvailable)
            {
                return;
            }

            var terminatorBuffer = new byte[1];

            while (this.thisStream.DataAvailable)
            {
                int read = await this.thisStream.ReadAsync(terminatorBuffer.AsMemory(0, 1), cancellationToken);
                if (read == 0)
                {
                    return;
                }

                if (terminatorBuffer[0] != (byte)'\r' &&
                    terminatorBuffer[0] != (byte)'\n')
                {
                    throw new InvalidOperationException("Unexpected extra data after SCPI binary block.");
                }
            }
        }

        // ###########################################################################################
        // Reads a single line response from the network stream, trimming CR/LF terminators.
        // ###########################################################################################
        private async Task<string> ReadLineAsync(CancellationToken cancellationToken)
        {
            if (this.thisStream == null)
            {
                throw new InvalidOperationException("No active oscilloscope session exists.");
            }

            var bytes = new List<byte>();
            var singleByte = new byte[1];

            while (true)
            {
                int read = await this.thisStream.ReadAsync(singleByte.AsMemory(0, 1), cancellationToken);

                if (read == 0)
                {
                    break;
                }

                if (singleByte[0] == (byte)'\n')
                {
                    break;
                }

                if (singleByte[0] != (byte)'\r')
                {
                    bytes.Add(singleByte[0]);
                }
            }

            return Encoding.ASCII.GetString(bytes.ToArray());
        }

        // ###########################################################################################
        // Reads the exact number of bytes requested from the network stream.
        // ###########################################################################################
        private async Task<byte[]> ReadExactlyAsync(int byteCount, CancellationToken cancellationToken)
        {
            if (this.thisStream == null)
            {
                throw new InvalidOperationException("No active oscilloscope session exists.");
            }

            var buffer = new byte[byteCount];
            int offset = 0;

            while (offset < byteCount)
            {
                int read = await this.thisStream.ReadAsync(
                    buffer.AsMemory(offset, byteCount - offset),
                    cancellationToken);

                if (read == 0)
                {
                    throw new InvalidOperationException("Unexpected end of SCPI network stream.");
                }

                offset += read;
            }

            return buffer;
        }

        // ###########################################################################################
        // Closes the current SCPI network session and releases socket resources.
        // ###########################################################################################
        public ValueTask DisposeAsync()
        {
            this.thisStream?.Dispose();
            this.thisClient.Dispose();
            return ValueTask.CompletedTask;
        }
    }
}