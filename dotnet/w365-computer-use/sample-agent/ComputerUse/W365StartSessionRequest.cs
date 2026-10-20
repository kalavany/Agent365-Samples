// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.IdentityModel.Tokens.Jwt;
using System.Text.Json;

namespace W365ComputerUseSample.ComputerUse;

/// <summary>
/// Resolved options for one turn's W365 session starts. The caller identity comes only
/// from the access token sent to MCP, never from activity or model arguments.
/// </summary>
public sealed class W365StartSessionRequest
{
    private W365StartSessionRequest(string sessionKind, string userObjectId, int displayWidth, int displayHeight, string browserMode)
    {
        SessionKind = sessionKind;
        UserObjectId = userObjectId;
        DisplayWidth = displayWidth;
        DisplayHeight = displayHeight;
        BrowserMode = browserMode;
    }

    public string SessionKind { get; }
    public string UserObjectId { get; }
    public int DisplayWidth { get; }
    public int DisplayHeight { get; }
    public string BrowserMode { get; }

    /// <summary>
    /// Resolves Activity.Value overrides and configuration defaults using the outbound
    /// MCP token's oid. Reading this claim does not validate the token; MCP authenticates it.
    /// </summary>
    public static W365StartSessionRequest FromActivityValue(object? value, IConfiguration configuration, string accessToken)
    {
        ArgumentNullException.ThrowIfNull(configuration);
        var userObjectId = GetCallerObjectId(accessToken);
        JsonElement? options = null;
        if (value != null)
        {
            try
            {
                options = value is string json
                    ? JsonSerializer.Deserialize<JsonElement>(json)
                    : JsonSerializer.SerializeToElement(value);
            }
            catch (JsonException ex)
            {
                throw new ArgumentException("Activity.Value must be a JSON object containing session options.", nameof(value), ex);
            }

            if (options.Value.ValueKind != JsonValueKind.Object)
            {
                throw new ArgumentException("Activity.Value must be a JSON object containing session options.", nameof(value));
            }
        }

        JsonElement? properties = null;
        if (options is { } root && root.TryGetProperty("sessionProperties", out var suppliedProperties))
        {
            if (suppliedProperties.ValueKind != JsonValueKind.Object)
            {
                throw new ArgumentException("sessionProperties must be a JSON object.", nameof(value));
            }
            properties = suppliedProperties;
        }

        return new W365StartSessionRequest(
            ReadString(options, "sessionKind", configuration["ComputerUse:SessionKind"] ?? "local"),
            userObjectId,
            ReadDimension(properties, "displayWidth", configuration.GetValue("ComputerUse:DisplayWidth", 1024)),
            ReadDimension(properties, "displayHeight", configuration.GetValue("ComputerUse:DisplayHeight", 768)),
            ReadString(properties, "browserMode", configuration["ComputerUse:BrowserMode"] ?? "chromium"));
    }

    /// <summary>Builds fresh MCP arguments so tool invocation cannot mutate this request.</summary>
    public Dictionary<string, object?> ToArguments() => new()
    {
        ["sessionKind"] = SessionKind,
        ["userObjectId"] = UserObjectId,
        ["sessionProperties"] = new Dictionary<string, object?>
        {
            ["displayWidth"] = DisplayWidth,
            ["displayHeight"] = DisplayHeight,
            ["browserMode"] = BrowserMode
        }
    };

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

    private static string ReadString(JsonElement? source, string name, string fallback)
    {
        var value = fallback;
        if (source is { } element && element.TryGetProperty(name, out var supplied))
        {
            if (supplied.ValueKind != JsonValueKind.String)
            {
                throw new ArgumentException($"{name} must be a nonempty string.");
            }
            value = supplied.GetString()!;
        }
        if (string.IsNullOrWhiteSpace(value))
        {
            throw new ArgumentException($"{name} must be a nonempty string.");
        }
        return value.Trim();
    }

    private static int ReadDimension(JsonElement? source, string name, int fallback)
    {
        var value = fallback;
        if (source is { } element && element.TryGetProperty(name, out var supplied)
            && (supplied.ValueKind != JsonValueKind.Number || !supplied.TryGetInt32(out value)))
        {
            throw new ArgumentException($"{name} must be a positive integer.");
        }
        if (value <= 0)
        {
            throw new ArgumentException($"{name} must be a positive integer.");
        }
        return value;
    }
}
