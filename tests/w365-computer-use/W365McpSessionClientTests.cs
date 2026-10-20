// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.IdentityModel.Tokens.Jwt;
using System.Reflection;
using System.Security.Claims;
using System.Text.Json;
using System.Text.Json.Nodes;
using ModelContextProtocol.Client;
using ModelContextProtocol.Protocol;
using W365ComputerUseSample.ComputerUse;
using Xunit;

namespace W365ComputerUseSample.Tests;

public sealed class W365McpSessionClientTests
{
    [Fact]
    public async Task PrestartSendsResolvedPayloadBeforeListingToolsWithReturnedSessionId()
    {
        var mcpClient = DispatchProxy.Create<IMcpClient, RecordingMcpClient>();
        var recording = (RecordingMcpClient)mcpClient;
        var client = new W365McpSessionClient(mcpClient);
        var token = new JwtSecurityTokenHandler().WriteToken(new JwtSecurityToken(
            claims: [new Claim("oid", "11111111-1111-4111-8111-111111111111")]));
        var request = W365StartSessionRequest.FromActivityValue(
            new { sessionKind = "test-kind", sessionProperties = new { displayWidth = 1280, displayHeight = 720, browserMode = "test-browser" } },
            token);

        var result = await client.StartSessionAndListToolsAsync(request, CancellationToken.None);

        Assert.Equal("desktop-1", result.SessionId);
        Assert.Equal("https://example.test/screenshare", result.ScreenShareUrl);
        Assert.Equal(2, recording.Requests.Count);
        var call = recording.Requests[0];
        Assert.Equal("tools/call", call.Method);
        Assert.Equal("mcp_W365ComputerUse_StartSession", call.Params!["name"]!.GetValue<string>());
        var arguments = call.Params["arguments"]!.AsObject();
        Assert.Equal(new[] { "sessionKind", "sessionProperties", "userObjectId" }, arguments.Select(property => property.Key).Order());
        Assert.Equal("test-kind", arguments["sessionKind"]!.GetValue<string>());
        Assert.Equal("11111111-1111-4111-8111-111111111111", arguments["userObjectId"]!.GetValue<string>());
        var properties = arguments["sessionProperties"]!.AsObject();
        Assert.Equal(3, properties.Count);
        Assert.Equal(1280, properties["displayWidth"]!.GetValue<int>());
        Assert.Equal(720, properties["displayHeight"]!.GetValue<int>());
        Assert.Equal("test-browser", properties["browserMode"]!.GetValue<string>());
        Assert.Equal("tools/list", recording.Requests[1].Method);
        Assert.Equal("desktop-1", recording.Requests[1].Params!["_meta"]!["sessionId"]!.GetValue<string>());
    }

    public class RecordingMcpClient : DispatchProxy
    {
        public List<JsonRpcRequest> Requests { get; } = [];

        protected override object? Invoke(MethodInfo? targetMethod, object?[]? args)
        {
            if (targetMethod?.Name != "SendRequestAsync" || args?[0] is not JsonRpcRequest request)
            {
                throw new InvalidOperationException($"Unexpected MCP invocation: {targetMethod?.Name}");
            }
            Requests.Add(request);
            var result = request.Method switch
            {
                "tools/call" => JsonSerializer.SerializeToNode(new
                {
                    content = new[] { new { type = "text", text = """{"sessionId":"desktop-1","screenShareUrl":"https://example.test/screenshare"}""" } }
                }),
                "tools/list" => JsonNode.Parse("""{"tools":[]}"""),
                _ => throw new InvalidOperationException($"Unexpected MCP method: {request.Method}")
            };
            return Task.FromResult(new JsonRpcResponse { Id = request.Id, Result = result });
        }
    }
}
