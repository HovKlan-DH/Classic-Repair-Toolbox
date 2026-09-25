using System;
using System.IO;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace Handlers.DataHandling
{
    // ###########################################################################################
    // Reading and writing system.json (NewContributeStrategy.md Phase 4, task 7) - the file that
    // makes a published system folder self-describing.
    //
    // *** A system.json IS UNTRUSTED INPUT, EVEN THOUGH IT IS "OURS". *** It sits in the synced
    // Data tree on the user's own disk, where anyone can edit it, and it arrives over the network
    // during a sync. So nothing read out of it may be allowed to GRANT anything: the Maintainers
    // list is display names for showing in the app, and authority is decided by the server's
    // maintainers table alone. That separation is already stated on SystemDescriptor itself and is
    // repeated here because this is the class that would be tempting to trust.
    //
    // FAILURES ARE SOFT. A system with no descriptor, or with one that will not parse, is an
    // ordinary system that simply has no extra metadata - every board that shipped before this
    // file existed is exactly that. It must never stop a board loading, which is why Read answers
    // null rather than throwing.
    // ###########################################################################################
    public static class SystemDescriptorStore
    {
        // Named for what it describes, and sitting inside the system's own folder so a system is
        // one directory containing everything about itself - the same "one folder is the whole
        // record" convention WorklogManager uses for a workbook.
        public const string FileName = "system.json";

        private static readonly JsonSerializerOptions JsonOptions = new()
        {
            WriteIndented = true,

            // Missing members are the ordinary case, not an error: a descriptor written by an
            // older build has fewer fields, and one written by a newer build has more. Neither
            // should fail to load - a system.json this build only half understands is still worth
            // more than none.
            UnmappedMemberHandling = JsonUnmappedMemberHandling.Skip
        };

        // ###########################################################################################
        // The descriptor for one system folder, or null when there is none or it cannot be read.
        //
        // A DESCRIPTOR WITH A MALFORMED SystemId IS REJECTED rather than returned, because that
        // value reaches a lookup key. Everything else is taken as-is: a blank Origin or an empty
        // Maintainers list is simply less information, not a corrupt file.
        // ###########################################################################################
        public static SystemDescriptor? Read(string systemFolder)
        {
            if (string.IsNullOrWhiteSpace(systemFolder))
                return null;

            string path = Path.Combine(systemFolder, SystemDescriptorStore.FileName);

            if (!File.Exists(path))
                return null;

            try
            {
                string json = File.ReadAllText(path);

                SystemDescriptor? descriptor =
                    JsonSerializer.Deserialize<SystemDescriptor>(json, SystemDescriptorStore.JsonOptions);

                if (descriptor is null)
                    return null;

                if (!SystemDescriptorRules.IsValidSystemId(descriptor.SystemId))
                {
                    CrtLog.Warning(
                        $"Ignoring [{path}]: its SystemId is not a valid system id.");

                    return null;
                }

                return descriptor;
            }
            catch (Exception ex) when (ex is IOException or JsonException or UnauthorizedAccessException)
            {
                // Soft by design - see the class header. A board with an unreadable descriptor is
                // still a board.
                CrtLog.Warning($"Could not read [{path}]: [{ex.Message}]");
                return null;
            }
        }

        // ###########################################################################################
        // Writes a descriptor into a system folder.
        //
        // *** NOTHING IN THE RUNNING SERVER MAY CALL THIS AGAINST THE PRODUCTION TREE. *** The
        // service has no write permission there and the kernel refuses it (Phase 3 step 0's
        // filesystem interlock), which is deliberate and must not be worked around. This exists
        // for the publishing tool the maintainer runs themselves, and for tests.
        //
        // Throws rather than reporting a bool, unlike Read: a failed WRITE means a system was
        // published without its descriptor, which is a real failure the caller has to know about,
        // whereas a failed read is an ordinary absence.
        // ###########################################################################################
        public static void Write(string systemFolder, SystemDescriptor descriptor)
        {
            ArgumentException.ThrowIfNullOrWhiteSpace(systemFolder);
            ArgumentNullException.ThrowIfNull(descriptor);

            if (!SystemDescriptorRules.IsValidSystemId(descriptor.SystemId))
            {
                throw new ArgumentException(
                    $"Refusing to write a descriptor with an invalid SystemId [{descriptor.SystemId}].",
                    nameof(descriptor));
            }

            Directory.CreateDirectory(systemFolder);

            string path = Path.Combine(systemFolder, SystemDescriptorStore.FileName);
            string json = JsonSerializer.Serialize(descriptor, SystemDescriptorStore.JsonOptions);

            File.WriteAllText(path, json);
        }
    }
}
