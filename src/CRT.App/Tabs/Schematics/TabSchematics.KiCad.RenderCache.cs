using Avalonia;
using Avalonia.Controls;
using Avalonia.Controls.Shapes;
using Avalonia.Input;
using Avalonia.Input.GestureRecognizers;
using Avalonia.Input.Platform;
using Avalonia.Interactivity;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using Avalonia.Threading;
using Handlers.DataHandling;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using Tabs.TabSchematics;
using Handlers.Geometry;

namespace CRT;

// ###########################################################################################
// Builds and caches the per-net PCB render nodes that KiCad.Render draws, including the
// connected-segment chaining used to turn loose segments into polylines.
//
// Part of the TabSchematics partial class - see TabSchematics.axaml.cs for the tab overview.
// ###########################################################################################
public partial class TabSchematics
{
    private readonly object thisKiCadPcbNetRenderCacheSync = new();

    private readonly Dictionary<string, Task> thisKiCadPcbNetRenderBuildTaskByKey = new(StringComparer.OrdinalIgnoreCase);

    private readonly Dictionary<string, KiCadPcbNetRenderCache> thisKiCadPcbNetRenderCacheByKey = new(StringComparer.OrdinalIgnoreCase);

    // ###########################################################################################
    // Returns the cached PCB net graph for the requested net/layer.
    // The cache is stored both in the current working dictionaries and in the active persistent
    // per-board runtime cache scope so revisiting the same board can reuse the heavy build result.
    // ###########################################################################################
    private KiCadPcbNetRenderCache? GetOrCreateKiCadPcbNetRenderCache(
        KiCadPcb pcb,
        int pcbIndex,
        string netId,
        KiCadPcbHighlightBucket bucket,
        string requiredLayer)
    {
        string cacheKey = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCacheKey(pcbIndex, netId, requiredLayer);
        KiCadProjectBundle? expectedProject = this.thisKiCadProject;
        string expectedScopeKey = this.thisCurrentKiCadRuntimeCacheScopeKey;
        var activeScope = this.GetOrCreateCurrentKiCadRuntimeCacheScope();

        lock (this.thisKiCadPcbNetRenderCacheSync)
        {
            if (this.thisKiCadPcbNetRenderCacheByKey.TryGetValue(cacheKey, out var cache))
            {
                return cache;
            }

            if (activeScope != null &&
                activeScope.NetRenderCacheByKey.TryGetValue(cacheKey, out var scopedCache))
            {
                this.thisKiCadPcbNetRenderCacheByKey[cacheKey] = scopedCache;
                return scopedCache;
            }

            if (this.thisKiCadPcbNetRenderBuildTaskByKey.ContainsKey(cacheKey) ||
                (activeScope != null && activeScope.NetRenderBuildTaskByKey.ContainsKey(cacheKey)))
            {
                return null;
            }

            Task buildTask = Task.Run(() =>
            {
                try
                {
                    var builtCache = KiCadNetGraphBuilder.BuildKiCadPcbNetRenderCache(pcb, bucket, requiredLayer);

                    lock (this.thisKiCadPcbNetRenderCacheSync)
                    {
                        if (!ReferenceEquals(expectedProject, this.thisKiCadProject) ||
                            !string.Equals(expectedScopeKey, this.thisCurrentKiCadRuntimeCacheScopeKey, StringComparison.OrdinalIgnoreCase))
                        {
                            return;
                        }

                        this.thisKiCadPcbNetRenderCacheByKey[cacheKey] = builtCache;

                        if (activeScope != null)
                        {
                            activeScope.NetRenderCacheByKey[cacheKey] = builtCache;
                        }
                    }
                }
                catch (Exception ex)
                {
                    Logger.Warning($"Failed to build KiCad PCB net render cache [{cacheKey}] - [{ex.Message}]");
                }
                finally
                {
                    lock (this.thisKiCadPcbNetRenderCacheSync)
                    {
                        this.thisKiCadPcbNetRenderBuildTaskByKey.Remove(cacheKey);

                        if (activeScope != null)
                        {
                            activeScope.NetRenderBuildTaskByKey.Remove(cacheKey);
                        }
                    }

                    if (ReferenceEquals(expectedProject, this.thisKiCadProject) &&
                        string.Equals(expectedScopeKey, this.thisCurrentKiCadRuntimeCacheScopeKey, StringComparison.OrdinalIgnoreCase))
                    {
                        Dispatcher.UIThread.Post(
                            () => this.RefreshKiCadOverlay(),
                            DispatcherPriority.Background);
                    }
                }
            });

            this.thisKiCadPcbNetRenderBuildTaskByKey[cacheKey] = buildTask;

            if (activeScope != null)
            {
                activeScope.NetRenderBuildTaskByKey[cacheKey] = buildTask;
            }

            return null;
        }
    }
}