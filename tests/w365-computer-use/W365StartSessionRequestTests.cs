// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text.Json;
using System.Text.Json.Nodes;
using Microsoft.Extensions.Configuration;
using W365ComputerUseSample.ComputerUse;
using Xunit;

namespace W365ComputerUseSample.Tests;

public sealed class W365StartSessionRequestTests
{
    private const string CallerOid = "11111111-1111-4111-8111-111111111111";

    [Fact]
    public void MissingActivityValueUsesDefaultsAndTokenIdentity()
    {
        var request = W365StartSessionRequest.FromActivityValue(null, new ConfigurationBuilder().Build(), Token(CallerOid));

        AssertPayload(request, "local", CallerOid, 1024, 768, "chromium");
    }

    [Fact]
    public void MissingFieldsUseConfiguredDefaults()
    {
        var configuration = new ConfigurationBuilder().AddInMemoryCollection(new Dictionary<string, string?>
        {
            ["ComputerUse:SessionKind"] = "configured-kind",
            ["ComputerUse:DisplayWidth"] = "1600",
            ["ComputerUse:DisplayHeight"] = "900",
            ["ComputerUse:BrowserMode"] = "configured-browser"
        }).Build();

        var request = W365StartSessionRequest.FromActivityValue(new { sessionProperties = new { displayHeight = 800 } }, configuration, Token(CallerOid));

        AssertPayload(request, "configured-kind", CallerOid, 1600, 800, "configured-browser");
    }

    [Fact]
    public void SuppliedIdentityCannotOverrideMcpTokenIdentity()
    {
        var value = new
        {
            sessionKind = "request-kind",
            userObjectId = "22222222-2222-4222-8222-222222222222",
            sessionProperties = new
            {
                displayWidth = 1280,
                displayHeight = 720,
                browserMode = "request-browser"
            }
        };

        var request = W365StartSessionRequest.FromActivityValue(value, new ConfigurationBuilder().Build(), Token(CallerOid));

        AssertPayload(request, "request-kind", CallerOid, 1280, 720, "request-browser");
    }

    [Fact]
    public void JsonElementValuePreservesOverrides()
    {
        var value = JsonSerializer.Deserialize<JsonElement>("""{"sessionProperties":{"displayWidth":1440,"displayHeight":900}}""");

        var request = W365StartSessionRequest.FromActivityValue(value, new ConfigurationBuilder().Build(), Token(CallerOid));

        AssertPayload(request, "local", CallerOid, 1440, 900, "chromium");
    }

    [Fact]
    public void JsonNodeValuePreservesOverrides()
    {
        var value = JsonNode.Parse("""{"sessionKind":"node-kind"}""");

        var request = W365StartSessionRequest.FromActivityValue(value, new ConfigurationBuilder().Build(), Token(CallerOid));

        AssertPayload(request, "node-kind", CallerOid, 1024, 768, "chromium");
    }

    [Theory]
    [InlineData("[]")]
    [InlineData("null")]
    [InlineData("{")]
    [InlineData("""{"sessionKind":""}""")]
    [InlineData("""{"sessionKind":null}""")]
    [InlineData("""{"sessionProperties":null}""")]
    [InlineData("""{"sessionProperties":[]}""")]
    [InlineData("""{"sessionProperties":{"browserMode":" "}}""")]
    [InlineData("""{"sessionProperties":{"displayWidth":0}}""")]
    [InlineData("""{"sessionProperties":{"displayHeight":-1}}""")]
    [InlineData("""{"sessionProperties":{"displayWidth":"1280"}}""")]
    [InlineData("""{"sessionProperties":{"displayWidth":1.5}}""")]
    [InlineData("""{"sessionProperties":{"displayWidth":2147483648}}""")]
    public void InvalidSuppliedOptionsAreRejected(string value)
    {
        Assert.Throws<ArgumentException>(() =>
            W365StartSessionRequest.FromActivityValue(value, new ConfigurationBuilder().Build(), Token(CallerOid)));
    }

    [Theory]
    [InlineData("")]
    [InlineData("not-a-guid")]
    [InlineData("00000000-0000-0000-0000-000000000000")]
    public void InvalidTokenOidIsRejected(string oid)
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            W365StartSessionRequest.FromActivityValue(null, new ConfigurationBuilder().Build(), Token(oid)));

        Assert.Equal("The W365 MCP access token must contain exactly one nonempty GUID oid claim.", exception.Message);
    }

    [Fact]
    public void MissingTokenOidDoesNotFallBackToMessageIdentity()
    {
        var token = new JwtSecurityTokenHandler().WriteToken(new JwtSecurityToken(claims: [new Claim("sub", CallerOid)]));

        Assert.Throws<InvalidOperationException>(() =>
            W365StartSessionRequest.FromActivityValue(new { userObjectId = CallerOid }, new ConfigurationBuilder().Build(), token));
    }

    [Fact]
    public void NonJwtTokenIsRejected()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            W365StartSessionRequest.FromActivityValue(null, new ConfigurationBuilder().Build(), "opaque-token"));

        Assert.Equal("The W365 MCP access token must be a JWT with an oid claim.", exception.Message);
    }

    [Fact]
    public void ReturnedArgumentsCannotMutateRequestOrLaterCalls()
    {
        var request = W365StartSessionRequest.FromActivityValue(null, new ConfigurationBuilder().Build(), Token(CallerOid));
        var arguments = request.ToArguments();
        arguments["userObjectId"] = "spoofed";
        var properties = Assert.IsType<Dictionary<string, object?>>(arguments["sessionProperties"]);
        properties["displayWidth"] = 1;

        AssertPayload(request, "local", CallerOid, 1024, 768, "chromium");
    }

    private static string Token(string oid) =>
        new JwtSecurityTokenHandler().WriteToken(new JwtSecurityToken(claims: [new Claim("oid", oid)]));

    private static void AssertPayload(W365StartSessionRequest request, string kind, string oid, int width, int height, string browser)
    {
        var arguments = request.ToArguments();
        Assert.Equal(new[] { "sessionKind", "sessionProperties", "userObjectId" }, arguments.Keys.Order());
        var payload = JsonSerializer.SerializeToElement(arguments);
        Assert.Equal(kind, payload.GetProperty("sessionKind").GetString());
        Assert.Equal(oid, payload.GetProperty("userObjectId").GetString());
        var properties = payload.GetProperty("sessionProperties");
        Assert.Equal(3, properties.EnumerateObject().Count());
        Assert.Equal(width, properties.GetProperty("displayWidth").GetInt32());
        Assert.Equal(height, properties.GetProperty("displayHeight").GetInt32());
        Assert.Equal(browser, properties.GetProperty("browserMode").GetString());
    }
}
