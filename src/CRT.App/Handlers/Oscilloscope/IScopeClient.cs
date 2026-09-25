using System.Threading;
using System.Threading.Tasks;

namespace Handlers.Oscilloscope
{
    // ###########################################################################################
    // The thin "talk to the scope" seam, exactly the same idea as IMiniproRunner: everything above
    // this line is command sequencing and response handling, which is hardware-agnostic and worth
    // testing; everything below it is a TCP socket, which is not.
    //
    // ScopeScpiClient is the real implementation and is deliberately left uncovered (see
    // .claude/CLAUDE.md's "deliberately not covered" list - real TCP is an I/O boundary). The
    // sequencing that used to be welded to it can now be driven against a fake that returns canned
    // responses, with no scope on the network.
    //
    // Only the three methods the Oscilloscope tab actually calls are here. ConnectAsync and
    // DisposeAsync stay off the interface on purpose: they are connection lifecycle, owned by the
    // tab's connect path, and nothing in the sequencing logic calls them.
    // ###########################################################################################
    public interface IScopeClient
    {
        /// <summary>Sends a SCPI command that expects no response.</summary>
        Task SendAsync(string commandText, CancellationToken cancellationToken);

        /// <summary>Sends a SCPI query and reads a single line back.</summary>
        Task<string> QueryLineAsync(string commandText, CancellationToken cancellationToken);

        /// <summary>
        /// Sends a SCPI query and reads a definite-length binary block, returned raw INCLUDING its
        /// "#&lt;digits&gt;&lt;length&gt;" header - the screenshot dump path relies on the header
        /// still being present.
        /// </summary>
        Task<byte[]> QueryBinaryBlockAsync(string commandText, CancellationToken cancellationToken);
    }
}
