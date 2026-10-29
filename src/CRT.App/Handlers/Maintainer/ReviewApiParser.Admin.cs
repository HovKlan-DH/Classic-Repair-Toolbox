using System.Linq;
using System.Text.Json;
using Handlers.DataHandling;

namespace Handlers.MaintainerHandling
{
    // ###########################################################################################
    // ReviewApiParser, two of the administrator's tools (owner requests, 2026-10-04): "Reset
    // contribution data" - the counts shown and what the reset deleted - and "API usage". All read
    // into CRT.Data's own records, which the server writes; ReviewWireContractTests holds the two ends
    // together. The rules in ReviewApiParser.cs's header hold here.
    // ###########################################################################################
    public static partial class ReviewApiParser
    {
        // ###########################################################################################
        // GET /api/admin/reset. Refused without a fingerprint - counts that cannot be sent back are
        // not counts the reset would accept, so no button may be offered on them.
        // ###########################################################################################
        public static DataResetPlanAnswer? ParseDataResetPlan(string? json)
        {
            DataResetPlanAnswer? plan = ReviewApiParser.Read<DataResetPlanAnswer>(json);

            return string.IsNullOrWhiteSpace(plan?.Fingerprint) ? null : plan;
        }

        // POST /api/admin/reset - what went. Every count is a number; a missing one reads 0.
        public static DataResetAnswer? ParseDataReset(string? json) =>
            ReviewApiParser.Read<DataResetAnswer>(json);

        // ###########################################################################################
        // GET /api/admin/oscilloscope-models (CRT 4.0 step 5). Refused without a list. An entry
        // without its name is left out rather than shown blank; missing values and words read empty.
        // ###########################################################################################
        public static OscilloscopeModelQueueAnswer? ParseOscilloscopeModelQueue(string? json)
        {
            OscilloscopeModelQueueAnswer? answer = ReviewApiParser.Read<OscilloscopeModelQueueAnswer>(json);

            if (answer?.Entries is null)
                return null;

            return answer with
            {
                Entries = answer.Entries
                    .Where(entry => entry is not null && !string.IsNullOrWhiteSpace(entry.Vendor) && !string.IsNullOrWhiteSpace(entry.SeriesOrModel))
                    .Select(entry => entry with
                    {
                        Values = entry.Values ?? new System.Collections.Generic.Dictionary<string, string>(),
                        ContactName = entry.ContactName ?? string.Empty,
                        ContactEmail = entry.ContactEmail ?? string.Empty,
                        Message = entry.Message ?? string.Empty,
                        ClientVersion = entry.ClientVersion ?? string.Empty,
                    })
                    .ToList()
            };
        }

        // POST /api/admin/oscilloscope-models/publish and /reject - refused without its sentence.
        public static OscilloscopeModelDecisionAnswer? ParseOscilloscopeModelDecision(string? json)
        {
            OscilloscopeModelDecisionAnswer? answer = ReviewApiParser.Read<OscilloscopeModelDecisionAnswer>(json);

            return string.IsNullOrWhiteSpace(answer?.Message) ? null : answer;
        }

        // ###########################################################################################
        // GET /api/admin/api-usage. Refused without a route list - every server that has the route
        // maps routes, so an answer with none is not one. A route or version without its name is left
        // out rather than shown blank; missing lists read as empty.
        // ###########################################################################################
        public static ApiUsageAnswer? ParseApiUsage(string? json)
        {
            ApiUsageAnswer? answer = ReviewApiParser.Read<ApiUsageAnswer>(json);

            if (answer?.Routes is null)
                return null;

            return answer with
            {
                Routes = answer.Routes
                    .Where(route => route is not null && !string.IsNullOrWhiteSpace(route.Route))
                    .Select(route => route with
                    {
                        Method = route.Method ?? string.Empty,
                        Area = route.Area ?? string.Empty,
                        Versions = (route.Versions ?? [])
                            .Where(version => version is not null && !string.IsNullOrWhiteSpace(version.Version))
                            .ToList()
                    })
                    .ToList(),
                Installations = (answer.Installations ?? [])
                    .Where(installation => installation is not null && !string.IsNullOrWhiteSpace(installation.Version))
                    .ToList()
            };
        }

        private static T? Read<T>(string? json)
            where T : class
        {
            JsonElement root = ReviewApiParser.Root(json);

            if (root.ValueKind != JsonValueKind.Object)
                return null;

            try
            {
                return root.Deserialize<T>(ReviewApiParser.FactOptions);
            }
            catch (JsonException)
            {
                return null;
            }
        }
    }
}
