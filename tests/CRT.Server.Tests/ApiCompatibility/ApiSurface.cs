using System.Reflection;
using System.Text.Json;
using System.Text.Json.Nodes;
using System.Text.Json.Serialization.Metadata;
using CRT.Server.Handlers.Health;
using Handlers.DataHandling;
using Microsoft.AspNetCore.Http.Metadata;
using Microsoft.AspNetCore.Routing;

namespace CRT.Server.Tests.ApiCompatibility
{
    // ###########################################################################################
    // EVERYTHING AN INSTALLED CRT CAN DEPEND ON, AS LINES OF TEXT (2026-10-04, "Installed CRTs keep
    // working"). One line per fact, sorted, so a released CRT's surface can be frozen into a file
    // and every later surface compared with it line by line:
    //
    //   route POST /api/submissions                         a route, by method and pattern
    //   body POST /api/accounts/login | .email string       a JSON request body's field, per route
    //   answer SessionAnswer | .account.id number           an answer record's field, by record
    //   enum ApproverRole | Maintainer -> "Maintainer"      how an enum travels, member by member
    //   const CheckInContract.ControlField = "control"      a wire name of a form or header
    //
    // Fields are read from System.Text.Json's own metadata under the server's wire settings, so a
    // [JsonPropertyName], a naming-policy change or a computed property on the wire is seen exactly
    // as a client sees it. Request bodies come from the route table (each route's accepted type),
    // so a server-side body record is covered with no list to keep. Answers are the CRT.Data
    // records named *Answer or *Response, plus the few below - the server writes no anonymous
    // answer with data in it (ApiCompatibilityTests holds it to that).
    //
    // Only STRING constants are listed: a numeric limit raised is not a break, and listing it would
    // make every raise look like one.
    // ###########################################################################################
    internal static class ApiSurface
    {
        // Answers whose names do not end in Answer/Response - and the health check, a server type.
        private static readonly Type[] OtherAnswers =
        [
            typeof(SubmissionResult),
            typeof(SubmissionStatus),
            typeof(ValidationFinding),
            typeof(ReviewTableData),
            typeof(HealthStatus)
        ];

        // The classes whose string constants are wire names (form fields, paths, headers, codes).
        private static readonly Type[] WireNameHolders =
        [
            typeof(CheckInContract),
            typeof(FeedbackContract),
            typeof(BoardViewContract),
            typeof(DraftDiscardContract),
            typeof(ClientVersionContract),
            typeof(SubmissionFormat)
        ];

        private const int MaximumDepth = 12;

        public static SortedSet<string> Build()
        {
            var lines = new SortedSet<string>(StringComparer.Ordinal);
            var enums = new HashSet<Type>();
            JsonSerializerOptions options = ReviewApiContract.WireSettings;

            foreach (RouteEndpoint route in ServerRouteTable.Routes)
            {
                string name = ServerRouteTable.NameOf(route);
                lines.Add("route " + name);

                Type? body = route.Metadata.GetMetadata<IAcceptsMetadata>()?.RequestType;

                if (body is not null)
                    ApiSurface.Walk(lines, enums, options, "body " + name, string.Empty, body, 0);
            }

            foreach (Type answer in ApiSurface.AnswerRoots())
                ApiSurface.Walk(lines, enums, options, "answer " + answer.Name, string.Empty, answer, 0);

            foreach (Type type in enums)
            {
                foreach (object member in Enum.GetValues(type))
                    lines.Add($"enum {type.Name} | {member} -> {JsonSerializer.Serialize(member, type, options)}");
            }

            foreach (Type holder in ApiSurface.WireNameHolders)
            {
                foreach (FieldInfo field in holder.GetFields(BindingFlags.Public | BindingFlags.Static))
                {
                    if (field.IsLiteral && field.FieldType == typeof(string))
                        lines.Add($"const {holder.Name}.{field.Name} = \"{field.GetRawConstantValue()}\"");
                }
            }

            return lines;
        }

        // ###########################################################################################
        // Every line of `frozen` that `current` no longer has - each one something a released CRT
        // relied on - plus every enum member `current` ADDED to an enum `frozen` lists: an installed
        // CRT reading a name it has never seen into its enum fails on it. Lines starting with "#"
        // are comments; a break listed in `allowed` (its line, before any " || reason") is a
        // decision already taken and is not reported again.
        // ###########################################################################################
        public static IReadOnlyList<string> Breaks(
            IEnumerable<string> frozen,
            IReadOnlySet<string> current,
            IEnumerable<string> allowed)
        {
            var allowedLines = new HashSet<string>(
                allowed
                    .Select(line => line.Trim())
                    .Where(line => line.Length > 0 && !line.StartsWith('#'))
                    .Select(line => line.Split(" || ", 2)[0].Trim()),
                StringComparer.Ordinal);

            var frozenLines = frozen
                .Select(line => line.Trim())
                .Where(line => line.Length > 0 && !line.StartsWith('#'))
                .ToHashSet(StringComparer.Ordinal);

            var breaks = new List<string>();

            foreach (string line in frozenLines.OrderBy(line => line, StringComparer.Ordinal))
            {
                if (!current.Contains(line) && !allowedLines.Contains(line))
                    breaks.Add("gone:  " + line);
            }

            var frozenEnums = frozenLines
                .Where(line => line.StartsWith("enum ", StringComparison.Ordinal))
                .Select(ApiSurface.EnumNameOf)
                .ToHashSet(StringComparer.Ordinal);

            foreach (string line in current)
            {
                if (line.StartsWith("enum ", StringComparison.Ordinal) &&
                    frozenEnums.Contains(ApiSurface.EnumNameOf(line)) &&
                    !frozenLines.Contains(line) &&
                    !allowedLines.Contains(line))
                {
                    breaks.Add("added: " + line);
                }
            }

            return breaks;
        }

        private static string EnumNameOf(string enumLine) =>
            enumLine.Substring("enum ".Length).Split(" | ", 2)[0];

        private static IEnumerable<Type> AnswerRoots() =>
            typeof(ReviewApiContract).Assembly.GetExportedTypes()
                .Where(type =>
                    type is { IsClass: true, IsAbstract: false, IsGenericTypeDefinition: false } &&
                    (type.Name.EndsWith("Answer", StringComparison.Ordinal) ||
                     type.Name.EndsWith("Response", StringComparison.Ordinal)))
                .Concat(ApiSurface.OtherAnswers)
                .Distinct();

        private static void Walk(
            SortedSet<string> lines,
            HashSet<Type> enums,
            JsonSerializerOptions options,
            string prefix,
            string path,
            Type declared,
            int depth)
        {
            Type type = Nullable.GetUnderlyingType(declared) ?? declared;

            if (path.Length == 0)
                lines.Add($"{prefix} | (root) {ApiSurface.KindOf(type, options)}");

            if (depth > ApiSurface.MaximumDepth || ApiSurface.IsScalar(type))
                return;

            if (type.IsEnum)
            {
                enums.Add(type);
                return;
            }

            JsonTypeInfo info = options.GetTypeInfo(type);

            switch (info.Kind)
            {
                case JsonTypeInfoKind.Enumerable:
                    ApiSurface.Walk(lines, enums, options, prefix, path + "[]", info.ElementType!, depth + 1);
                    break;

                case JsonTypeInfoKind.Dictionary:
                    ApiSurface.Walk(lines, enums, options, prefix, path + "{}", info.ElementType!, depth + 1);
                    break;

                case JsonTypeInfoKind.Object:
                    foreach (JsonPropertyInfo property in info.Properties)
                    {
                        string fieldPath = $"{path}.{property.Name}";
                        lines.Add($"{prefix} | {fieldPath} {ApiSurface.KindOf(property.PropertyType, options)}");
                        ApiSurface.Walk(lines, enums, options, prefix, fieldPath, property.PropertyType, depth + 1);
                    }

                    break;
            }
        }

        // How a value of this type looks in JSON - what a reader has to cope with.
        private static string KindOf(Type declared, JsonSerializerOptions options)
        {
            Type type = Nullable.GetUnderlyingType(declared) ?? declared;

            if (type == typeof(bool))
                return "boolean";

            if (type == typeof(string) || type == typeof(char) || type == typeof(Guid) ||
                type == typeof(DateTime) || type == typeof(DateTimeOffset) || type == typeof(TimeSpan) ||
                type == typeof(DateOnly) || type == typeof(TimeOnly))
            {
                return "string";
            }

            if (type.IsEnum)
                return "enum " + type.Name;

            if (type.IsPrimitive || type == typeof(decimal))
                return "number";

            if (type == typeof(object) || type == typeof(JsonElement) || typeof(JsonNode).IsAssignableFrom(type))
                return "any";

            JsonTypeInfo info = options.GetTypeInfo(type);

            return info.Kind switch
            {
                JsonTypeInfoKind.Enumerable => "array of " + ApiSurface.KindOf(info.ElementType!, options),
                JsonTypeInfoKind.Dictionary => "map of " + ApiSurface.KindOf(info.ElementType!, options),
                JsonTypeInfoKind.Object => "object",
                _ => "value " + type.Name
            };
        }

        private static bool IsScalar(Type type) =>
            type.IsPrimitive || type == typeof(string) || type == typeof(decimal) || type == typeof(Guid) ||
            type == typeof(DateTime) || type == typeof(DateTimeOffset) || type == typeof(TimeSpan) ||
            type == typeof(DateOnly) || type == typeof(TimeOnly) || type == typeof(object) ||
            type == typeof(JsonElement) || typeof(JsonNode).IsAssignableFrom(type);
    }
}
