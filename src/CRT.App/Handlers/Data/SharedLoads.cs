using System;
using System.Collections.Concurrent;
using System.Collections.Generic;
using System.Threading.Tasks;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Loads that share one run while they overlap: a second request for a key whose load is still
    // running gets that same load's result instead of starting its own (code review, 2026-10-13).
    //
    // Made for the board workbooks (DataManager.LoadBoardDataAsync): BoardDataReader's cache only
    // answers once a parse has FINISHED, so the Configuration tab's catalogue, the launch validation
    // and the board on screen asking for one workbook at the same moment parsed it two or three
    // times over. Nothing is remembered once a load is done - that is BoardDataReader's cache.
    // A load that fails fails for everybody who joined it.
    // ###########################################################################################
    internal sealed class SharedLoads<TValue>
    {
        private readonly ConcurrentDictionary<string, Task<TValue>> thisRunning;

        internal SharedLoads(StringComparer keyComparer)
        {
            this.thisRunning = new ConcurrentDictionary<string, Task<TValue>>(keyComparer);
        }

        // How many loads are running now.
        internal int RunningCount => this.thisRunning.Count;

        internal Task<TValue> RunAsync(string key, Func<Task<TValue>> load)
        {
            while (true)
            {
                if (this.thisRunning.TryGetValue(key, out var running))
                {
                    return running;
                }

                var completion = new TaskCompletionSource<TValue>(TaskCreationOptions.RunContinuationsAsynchronously);

                if (this.thisRunning.TryAdd(key, completion.Task))
                {
                    _ = this.RunAndCompleteAsync(key, completion, load);
                    return completion.Task;
                }
            }
        }

        private async Task RunAndCompleteAsync(string key, TaskCompletionSource<TValue> completion, Func<Task<TValue>> load)
        {
            try
            {
                completion.SetResult(await load().ConfigureAwait(false));
            }
            catch (Exception ex)
            {
                completion.SetException(ex);
            }
            finally
            {
                // Only this load's entry - a later one may already stand under the key.
                this.thisRunning.TryRemove(new KeyValuePair<string, Task<TValue>>(key, completion.Task));
            }
        }
    }
}
