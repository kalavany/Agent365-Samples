// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.IdentityModel.Tokens.Jwt;
using System.Text.Json;

namespace W365ComputerUseSample.ComputerUse;

/// <summary>
/// StartSession arguments for one turn. <c>sessionKind</c> and <c>sessionProperties</c> are optional and
/// come only from the user's Activity.Value; nothing is filled in from configuration, so a turn without them
/// keeps the default W365 checkout. <c>userObjectId</c> comes only from the MCP access token.
/// </summary>
public sealed class W365StartSessionRequest
{
    private static readonly string[] IntegerProperties = ["displayWidth", "displayHeight"];
    private static readonly string[] StringProperties = ["browserMode"];

    private readonly string? sessionPropertiesJson;

    private W365StartSessionRequest(string? sessionKind, string userObjectId, JsonElement? sessionProperties)
    {
        SessionKind = sessionKind;
        UserObjectId = userObjectId;
        sessionPropertiesJson = sessionProperties?.GetRawText();
        if (sessionProperties is { } properties)
        {
            DisplayWidth = properties.TryGetProperty("displayWidth", out var width) ? width.GetInt32() : null;
            DisplayHeight = properties.TryGetProperty("displayHeight", out var height) ? height.GetInt32() : null;
        }
    }

    /// <summary>User-supplied session kind, or null when not supplied.</summary>
    public string? SessionKind { get; }

    public string UserObjectId { get; }

    /// <summary>User-supplied display width, or null when not supplied.</summary>
    public int? DisplayWidth { get; }

    /// <summary>User-supplied display height, or null when not supplied.</summary>
    public int? DisplayHeight { get; }

    /// <summary>
    /// Reads the optional <c>sessionKind</c> and <c>sessionProperties</c> from Activity.Value, and the
    /// required caller <c>oid</c> from the outbound MCP token. Reading the claim does not validate the
    /// token; the MCP server authenticates it. The MCP server enforces the <c>sessionKind</c> values and
    /// that <c>sessionProperties</c> requires <c>sessionKind</c>.
    /// </summary>
    public static W365StartSessionRequest FromActivityValue(object? value, string accessToken)
    {
        var userObjectId = GetCallerObjectId(accessToken);
        if (value == null)
        {
            return new W365StartSessionRequest(null, userObjectId, null);
        }

        JsonElement options;
        try
        {
            options = value is string json
                ? JsonSerializer.Deserialize<JsonElement>(json)
                : JsonSerializer.SerializeToElement(value);
        }
        catch (JsonException ex)
        {
            throw new ArgumentException("Activity.Value must be a JSON object.", nameof(value), ex);
        }

        if (options.ValueKind != JsonValueKind.Object)
        {
            throw new ArgumentException("Activity.Value must be a JSON object.", nameof(value));
        }

        string? sessionKind = null;
        if (options.TryGetProperty("sessionKind", out var suppliedKind))
        {
            if (suppliedKind.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(suppliedKind.GetString()))
            {
                throw new ArgumentException("sessionKind must be a nonempty string.", nameof(value));
            }

            sessionKind = suppliedKind.GetString();
        }

        JsonElement? sessionProperties = null;
        if (options.TryGetProperty("sessionProperties", out var suppliedProperties))
        {
            ValidateSessionProperties(suppliedProperties);
            sessionProperties = suppliedProperties.Clone();
        }

        return new W365StartSessionRequest(sessionKind, userObjectId, sessionProperties);
    }

    /// <summary>
    /// Builds fresh MCP arguments so tool invocation cannot mutate this request. <c>sessionKind</c> and
    /// <c>sessionProperties</c> are sent exactly as supplied and omitted when the user did not supply them.
    /// </summary>
    public Dictionary<string, object?> ToArguments()
    {
        var arguments = new Dictionary<string, object?>
        {
            ["userObjectId"] = UserObjectId,
        };
        if (SessionKind != null)
        {
            arguments["sessionKind"] = SessionKind;
        }

        if (sessionPropertiesJson != null)
        {
            arguments["sessionProperties"] = JsonSerializer.Deserialize<Dictionary<string, JsonElement>>(sessionPropertiesJson)!
                .ToDictionary(property => property.Key, property => (object?)property.Value);
        }

        return arguments;
    }

    internal static string GetCallerObjectId(string accessToken)
    {
        var handler = new JwtSecurityTokenHandler();
        if (string.IsNullOrWhiteSpace(accessToken) || !handler.CanReadToken(accessToken))
        {
            throw new InvalidOperationException("The W365 MCP access token must be a JWT with an oid claim.");
        }

        var claims = handler.ReadJwtToken(accessToken).Claims.Where(claim => claim.Type == "oid").ToArray();
        if (claims.Length != 1 || !Guid.TryParse(claims[0].Value, out var oid) || oid == Guid.Empty)
        {
            throw new InvalidOperationException("The W365 MCP access token must contain exactly one nonempty GUID oid claim.");
        }
        return oid.ToString();
    }

    private static void ValidateSessionProperties(JsonElement properties)
    {
        if (properties.ValueKind != JsonValueKind.Object)
        {
            throw new ArgumentException("sessionProperties must be a JSON object.");
        }

        foreach (var name in IntegerProperties)
        {
            if (properties.TryGetProperty(name, out var supplied)
                && (supplied.ValueKind != JsonValueKind.Number || !supplied.TryGetInt32(out var number) || number <= 0))
            {
                throw new ArgumentException($"sessionProperties.{name} must be a positive integer.");
            }
        }

        foreach (var name in StringProperties)
        {
            if (properties.TryGetProperty(name, out var supplied)
                && (supplied.ValueKind != JsonValueKind.String || string.IsNullOrWhiteSpace(supplied.GetString())))
            {
                throw new ArgumentException($"sessionProperties.{name} must be a nonempty string.");
            }
        }
    }
}
