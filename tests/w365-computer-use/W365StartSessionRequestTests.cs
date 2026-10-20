// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text.Json;
using System.Text.Json.Nodes;
using W365ComputerUseSample.ComputerUse;
using Xunit;

namespace W365ComputerUseSample.Tests;

public sealed class W365StartSessionRequestTests
{
    private const string CallerOid = "11111111-1111-4111-8111-111111111111";

    [Fact]
    public void SuppliedValuesAreSentExactlyWithTokenIdentity()
    {
        var value = new
        {
            sessionKind = "local",
            sessionProperties = new
            {
                displayWidth = 1280,
                displayHeight = 720,
                browserMode = "chromium"
            }
        };

        var request = W365StartSessionRequest.FromActivityValue(value, Token(CallerOid));

        AssertPayload(
            request,
            """{"sessionKind":"local","userObjectId":"11111111-1111-4111-8111-111111111111","sessionProperties":{"displayWidth":1280,"displayHeight":720,"browserMode":"chromium"}}""");
        Assert.Equal(1280, request.DisplayWidth);
        Assert.Equal(720, request.DisplayHeight);
    }

    [Fact]
    public void OmittedSessionPropertiesAreNotSentOrDefaulted()
    {
        var request = W365StartSessionRequest.FromActivityValue(new { sessionKind = "cloud" }, Token(CallerOid));

        AssertPayload(request, """{"sessionKind":"cloud","userObjectId":"11111111-1111-4111-8111-111111111111"}""");
        Assert.Null(request.DisplayWidth);
        Assert.Null(request.DisplayHeight);
    }

    [Fact]
    public void PartialSessionPropertiesAreSentWithoutFillingMissingFields()
    {
        var request = W365StartSessionRequest.FromActivityValue(
            new { sessionKind = "local", sessionProperties = new { displayHeight = 800 } },
            Token(CallerOid));

        AssertPayload(
            request,
            """{"sessionKind":"local","userObjectId":"11111111-1111-4111-8111-111111111111","sessionProperties":{"displayHeight":800}}""");
        Assert.Null(request.DisplayWidth);
        Assert.Equal(800, request.DisplayHeight);
    }

    [Fact]
    public void EmptyAndAdditionalSessionPropertiesArePassedThrough()
    {
        var empty = W365StartSessionRequest.FromActivityValue(
            new { sessionKind = "local", sessionProperties = new { } },
            Token(CallerOid));
        var additional = W365StartSessionRequest.FromActivityValue(
            """{"sessionKind":"local","sessionProperties":{"locale":"en-US","persist":true}}""",
            Token(CallerOid));

        AssertPayload(empty, """{"sessionKind":"local","userObjectId":"11111111-1111-4111-8111-111111111111","sessionProperties":{}}""");
        AssertPayload(
            additional,
            """{"sessionKind":"local","userObjectId":"11111111-1111-4111-8111-111111111111","sessionProperties":{"locale":"en-US","persist":true}}""");
    }

    [Fact]
    public void SuppliedIdentityCannotOverrideMcpTokenIdentity()
    {
        var request = W365StartSessionRequest.FromActivityValue(
            new { sessionKind = "local", userObjectId = "22222222-2222-4222-8222-222222222222" },
            Token(CallerOid));

        AssertPayload(request, """{"sessionKind":"local","userObjectId":"11111111-1111-4111-8111-111111111111"}""");
    }

    [Fact]
    public void JsonElementAndJsonNodeValuesAreAccepted()
    {
        var element = JsonSerializer.Deserialize<JsonElement>("""{"sessionKind":"element-kind"}""");
        var node = JsonNode.Parse("""{"sessionKind":"node-kind","sessionProperties":{"browserMode":"edge"}}""");

        AssertPayload(
            W365StartSessionRequest.FromActivityValue(element, Token(CallerOid)),
            """{"sessionKind":"element-kind","userObjectId":"11111111-1111-4111-8111-111111111111"}""");
        AssertPayload(
            W365StartSessionRequest.FromActivityValue(node, Token(CallerOid)),
            """{"sessionKind":"node-kind","userObjectId":"11111111-1111-4111-8111-111111111111","sessionProperties":{"browserMode":"edge"}}""");
    }

    [Theory]
    [InlineData("[]")]
    [InlineData("null")]
    [InlineData("{")]
    [InlineData("""{"sessionKind":""}""")]
    [InlineData("""{"sessionKind":" "}""")]
    [InlineData("""{"sessionKind":null}""")]
    [InlineData("""{"sessionKind":1}""")]
    [InlineData("""{"sessionKind":"local","sessionProperties":null}""")]
    [InlineData("""{"sessionKind":"local","sessionProperties":[]}""")]
    [InlineData("""{"sessionKind":"local","sessionProperties":{"browserMode":" "}}""")]
    [InlineData("""{"sessionKind":"local","sessionProperties":{"browserMode":1}}""")]
    [InlineData("""{"sessionKind":"local","sessionProperties":{"displayWidth":0}}""")]
    [InlineData("""{"sessionKind":"local","sessionProperties":{"displayHeight":-1}}""")]
    [InlineData("""{"sessionKind":"local","sessionProperties":{"displayWidth":"1280"}}""")]
    [InlineData("""{"sessionKind":"local","sessionProperties":{"displayWidth":1.5}}""")]
    [InlineData("""{"sessionKind":"local","sessionProperties":{"displayWidth":2147483648}}""")]
    public void InvalidSuppliedOptionsAreRejected(string value)
    {
        Assert.Throws<ArgumentException>(() => W365StartSessionRequest.FromActivityValue(value, Token(CallerOid)));
    }

    [Theory]
    [InlineData(null)]
    [InlineData("{}")]
    [InlineData("""{"other":"ignored"}""")]
    public void NoSessionOptionsSendsNeitherSessionKindNorProperties(string? value)
    {
        var request = W365StartSessionRequest.FromActivityValue(value, Token(CallerOid));

        AssertPayload(request, """{"userObjectId":"11111111-1111-4111-8111-111111111111"}""");
        Assert.Null(request.SessionKind);
        Assert.Null(request.DisplayWidth);
        Assert.Null(request.DisplayHeight);
    }

    [Fact]
    public void SessionKindIsForwardedExactlyWithoutTrimming()
    {
        var request = W365StartSessionRequest.FromActivityValue("""{"sessionKind":"agentic "}""", Token(CallerOid));

        AssertPayload(request, """{"sessionKind":"agentic ","userObjectId":"11111111-1111-4111-8111-111111111111"}""");
    }

    [Fact]
    public void SessionPropertiesWithoutSessionKindAreForwardedForServerValidation()
    {
        var request = W365StartSessionRequest.FromActivityValue("""{"sessionProperties":{"displayWidth":1280}}""", Token(CallerOid));

        AssertPayload(
            request,
            """{"userObjectId":"11111111-1111-4111-8111-111111111111","sessionProperties":{"displayWidth":1280}}""");
        Assert.Equal(1280, request.DisplayWidth);
    }

    [Theory]
    [InlineData("")]
    [InlineData("not-a-guid")]
    [InlineData("00000000-0000-0000-0000-000000000000")]
    public void InvalidTokenOidIsRejected(string oid)
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            W365StartSessionRequest.FromActivityValue(new { sessionKind = "local" }, Token(oid)));

        Assert.Equal("The W365 MCP access token must contain exactly one nonempty GUID oid claim.", exception.Message);
    }

    [Fact]
    public void MissingTokenOidDoesNotFallBackToMessageIdentity()
    {
        var token = new JwtSecurityTokenHandler().WriteToken(new JwtSecurityToken(claims: [new Claim("sub", CallerOid)]));

        Assert.Throws<InvalidOperationException>(() =>
            W365StartSessionRequest.FromActivityValue(new { sessionKind = "local", userObjectId = CallerOid }, token));
    }

    [Fact]
    public void NonJwtTokenIsRejected()
    {
        var exception = Assert.Throws<InvalidOperationException>(() =>
            W365StartSessionRequest.FromActivityValue(new { sessionKind = "local" }, "opaque-token"));

        Assert.Equal("The W365 MCP access token must be a JWT with an oid claim.", exception.Message);
    }

    [Fact]
    public void ReturnedArgumentsCannotMutateRequestOrLaterCalls()
    {
        var request = W365StartSessionRequest.FromActivityValue(
            new { sessionKind = "local", sessionProperties = new { displayWidth = 1280 } },
            Token(CallerOid));
        var arguments = request.ToArguments();
        arguments["userObjectId"] = "spoofed";
        var properties = Assert.IsType<Dictionary<string, object?>>(arguments["sessionProperties"]);
        properties["displayWidth"] = 1;

        AssertPayload(
            request,
            """{"sessionKind":"local","userObjectId":"11111111-1111-4111-8111-111111111111","sessionProperties":{"displayWidth":1280}}""");
    }

    private static string Token(string oid) =>
        new JwtSecurityTokenHandler().WriteToken(new JwtSecurityToken(claims: [new Claim("oid", oid)]));

    private static void AssertPayload(W365StartSessionRequest request, string expectedJson)
    {
        var actual = JsonNode.Parse(JsonSerializer.Serialize(request.ToArguments()));
        Assert.True(
            JsonNode.DeepEquals(JsonNode.Parse(expectedJson), actual),
            $"Expected {expectedJson} but was {actual!.ToJsonString()}");
    }
}
