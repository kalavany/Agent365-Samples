// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.IdentityModel.Tokens.Jwt;
using System.Text.Json;

namespace W365ComputerUseSample.ComputerUse;

/// <summary>
/// StartSession arguments for one turn. <c>sessionKind</c> and <c>sessionProperties</c> come only
/// from the user's Activity.Value; <c>userObjectId</c> comes only from the MCP access token.
/// No value is filled in from configuration.
/// </summary>
public sealed class W365StartSessionRequest
{
    private static readonly string[] IntegerProperties = ["displayWidth", "displayHeight"];
    private static readonly string[] StringProperties = ["browserMode"];

    private readonly string? sessionPropertiesJson;

    private W365StartSessionRequest(string sessionKind, string userObjectId, JsonElement? sessionProperties)
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

    public string SessionKind { get; }
    public string UserObjectId { get; }

    /// <summary>User-supplied display width, or null when not supplied.</summary>
    public int? DisplayWidth { get; }

    /// <summary>User-supplied display height, or null when not supplied.</summary>
    public int? DisplayHeight { get; }

    /// <summary>
    /// Reads the required <c>sessionKind</c> and optional <c>sessionProperties</c> from Activity.Value,
    /// and the required caller <c>oid</c> from the outbound MCP token. Reading the claim does not
    /// validate the token; the MCP server authenticates it.
    /// </summary>
    public static W365StartSessionRequest FromActivityValue(object? value, string accessToken)
    {
        var userObjectId = GetCallerObjectId(accessToken);
        if (value == null)
        {
            throw new ArgumentException("Activity.Value must be a JSON object containing sessionKind.", nameof(value));
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
            throw new ArgumentException("Activity.Value must be a JSON object containing sessionKind.", nameof(value), ex);
        }

        if (options.ValueKind != JsonValueKind.Object)
        {
            throw new ArgumentException("Activity.Value must be a JSON object containing sessionKind.", nameof(value));
        }

        if (!options.TryGetProperty("sessionKind", out var sessionKind)
            || sessionKind.ValueKind != JsonValueKind.String
            || string.IsNullOrWhiteSpace(sessionKind.GetString()))
        {
            throw new ArgumentException("sessionKind is required and must be a nonempty string.", nameof(value));
        }

        JsonElement? sessionProperties = null;
        if (options.TryGetProperty("sessionProperties", out var suppliedProperties))
        {
            ValidateSessionProperties(suppliedProperties);
            sessionProperties = suppliedProperties.Clone();
        }

        return new W365StartSessionRequest(sessionKind.GetString()!.Trim(), userObjectId, sessionProperties);
    }

    /// <summary>
    /// Builds fresh MCP arguments so tool invocation cannot mutate this request. <c>sessionProperties</c>
    /// is sent exactly as supplied and omitted when the user did not supply it.
    /// </summary>
    public Dictionary<string, object?> ToArguments()
    {
        var arguments = new Dictionary<string, object?>
        {
            ["sessionKind"] = SessionKind,
            ["userObjectId"] = UserObjectId,
        };
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
