// Copyright (c) Microsoft Corporation.
// Licensed under the MIT License.

using System.IdentityModel.Tokens.Jwt;
using System.Security.Claims;
using System.Text.Json;
using Microsoft.Extensions.AI;
using Microsoft.Extensions.Configuration;
using Microsoft.Extensions.Logging.Abstractions;
using W365ComputerUseSample.ComputerUse;
using Xunit;

namespace W365ComputerUseSample.Tests;

public sealed class StartSessionOrchestrationTests
{
    [Fact]
    public async Task LazyStartupSendsResolvedRequestAndMatchingModelDimensions()
    {
        var model = new ScriptedModel(
            """{"output":[{"type":"function_call","call_id":"title-1","name":"browser_get_title","arguments":"{}"}]}""",
            """{"output":[{"type":"message","content":[{"type":"output_text","text":"done"}]}]}""");
        using var httpClient = new HttpClient();
        var orchestrator = CreateOrchestrator(model, httpClient);
        var starts = new List<JsonElement>();
        IList<AITool> tools =
        [
            AIFunctionFactory.Create((AIFunctionArguments arguments) =>
            {
                starts.Add(JsonSerializer.SerializeToElement(arguments));
                return """{"sessionId":"desktop-1"}""";
            }, "mcp_W365ComputerUse_StartSession"),
            AIFunctionFactory.Create(() => "Sample title", "browser_get_title")
        ];
        var request = Request("11111111-1111-4111-8111-111111111111", 1280, 720);

        var result = await orchestrator.RunAsync("conversation-1", "Read the browser title", tools, startSessionRequest: request);

        Assert.Equal("done", result);
        AssertStartPayload(Assert.Single(starts), "11111111-1111-4111-8111-111111111111", 1280, 720);
        var computerTool = model.Requests[0].GetProperty("tools")[0];
        Assert.Equal(1280, computerTool.GetProperty("display_width").GetInt32());
        Assert.Equal(720, computerTool.GetProperty("display_height").GetInt32());
    }

    [Fact]
    public async Task RecoveryUsesCurrentTurnOptionsRatherThanPriorTurnOrModelOptions()
    {
        var model = new ScriptedModel(
            """{"output":[{"type":"function_call","call_id":"title-1","name":"browser_get_title","arguments":"{}"}]}""",
            """{"output":[{"type":"message","content":[{"type":"output_text","text":"first"}]}]}""",
            """{"output":[{"type":"function_call","call_id":"title-2","name":"browser_get_title","arguments":"{}"}]}""",
            """{"output":[{"type":"message","content":[{"type":"output_text","text":"second"}]}]}""");
        using var httpClient = new HttpClient();
        var orchestrator = CreateOrchestrator(model, httpClient);
        var starts = new List<JsonElement>();
        var titleCalls = 0;
        IList<AITool> tools =
        [
            AIFunctionFactory.Create((AIFunctionArguments arguments) =>
            {
                starts.Add(JsonSerializer.SerializeToElement(arguments));
                return JsonSerializer.Serialize(new { sessionId = $"desktop-{starts.Count}" });
            }, "mcp_W365ComputerUse_StartSession"),
            AIFunctionFactory.Create(() => ++titleCalls == 2 ? "session expired" : "Sample title", "browser_get_title"),
            AIFunctionFactory.Create(() => "ended", "mcp_W365ComputerUse_EndSession")
        ];
        const string oid = "11111111-1111-4111-8111-111111111111";

        Assert.Equal("first", await orchestrator.RunAsync("conversation", "first task", tools, startSessionRequest: Request(oid, 1024, 768)));
        var result = await orchestrator.RunAsync("conversation", "second task", tools, startSessionRequest: Request(oid, 1600, 900));

        Assert.Equal("second", result);
        Assert.Equal(2, starts.Count);
        AssertStartPayload(starts[0], oid, 1024, 768);
        AssertStartPayload(starts[1], oid, 1600, 900);
        Assert.Equal(3, titleCalls);
        Assert.Equal(1024, model.Requests[2].GetProperty("tools")[0].GetProperty("display_width").GetInt32());
        Assert.Equal(768, model.Requests[2].GetProperty("tools")[0].GetProperty("display_height").GetInt32());
        Assert.Equal(1600, model.Requests[3].GetProperty("tools")[0].GetProperty("display_width").GetInt32());
        Assert.Equal(900, model.Requests[3].GetProperty("tools")[0].GetProperty("display_height").GetInt32());
    }

    [Fact]
    public async Task ModelRequestedStartupCannotOverrideResolvedIdentityOrOptions()
    {
        var model = new ScriptedModel(
            """{"output":[{"type":"function_call","call_id":"start-1","name":"mcp_W365ComputerUse_StartSession","arguments":"{\"sessionKind\":\"spoofed\",\"userObjectId\":\"spoofed\",\"sessionProperties\":{\"displayWidth\":1}}"}]}""",
            """{"output":[{"type":"message","content":[{"type":"output_text","text":"done"}]}]}""");
        using var httpClient = new HttpClient();
        var orchestrator = CreateOrchestrator(model, httpClient);
        var starts = new List<JsonElement>();
        IList<AITool> tools =
        [
            AIFunctionFactory.Create((AIFunctionArguments arguments) =>
            {
                starts.Add(JsonSerializer.SerializeToElement(arguments));
                return """{"sessionId":"desktop-1"}""";
            }, "mcp_W365ComputerUse_StartSession")
        ];

        var result = await orchestrator.RunAsync("conversation", "start a desktop", [], additionalTools: tools,
            startSessionRequest: Request("11111111-1111-4111-8111-111111111111", 1280, 720));

        Assert.Equal("done", result);
        AssertStartPayload(Assert.Single(starts), "11111111-1111-4111-8111-111111111111", 1280, 720);
    }

    [Fact]
    public async Task StartupWithoutSessionPropertiesOmitsThemAndKeepsConfiguredModelToolSize()
    {
        var model = new ScriptedModel(
            """{"output":[{"type":"function_call","call_id":"title-1","name":"browser_get_title","arguments":"{}"}]}""",
            """{"output":[{"type":"message","content":[{"type":"output_text","text":"done"}]}]}""");
        using var httpClient = new HttpClient();
        var orchestrator = CreateOrchestrator(model, httpClient);
        var starts = new List<JsonElement>();
        IList<AITool> tools =
        [
            AIFunctionFactory.Create((AIFunctionArguments arguments) =>
            {
                starts.Add(JsonSerializer.SerializeToElement(arguments));
                return """{"sessionId":"desktop-1"}""";
            }, "mcp_W365ComputerUse_StartSession"),
            AIFunctionFactory.Create(() => "Sample title", "browser_get_title")
        ];
        var token = new JwtSecurityTokenHandler().WriteToken(new JwtSecurityToken(
            claims: [new Claim("oid", "11111111-1111-4111-8111-111111111111")]));
        var request = W365StartSessionRequest.FromActivityValue(new { sessionKind = "cloud" }, token);

        var result = await orchestrator.RunAsync("conversation", "Read the browser title", tools, startSessionRequest: request);

        Assert.Equal("done", result);
        var start = Assert.Single(starts);
        Assert.Equal(new[] { "sessionKind", "userObjectId" }, start.EnumerateObject().Select(property => property.Name).Order());
        Assert.Equal("cloud", start.GetProperty("sessionKind").GetString());
        Assert.Equal("11111111-1111-4111-8111-111111111111", start.GetProperty("userObjectId").GetString());
        var computerTool = model.Requests[0].GetProperty("tools")[0];
        Assert.Equal(1024, computerTool.GetProperty("display_width").GetInt32());
        Assert.Equal(768, computerTool.GetProperty("display_height").GetInt32());
    }

    [Fact]
    public async Task NonComputerUseTurnDoesNotRequireSessionContext()
    {
        var model = new ScriptedModel("""{"output":[{"type":"message","content":[{"type":"output_text","text":"hello"}]}]}""");
        using var httpClient = new HttpClient();
        var orchestrator = CreateOrchestrator(model, httpClient);

        var result = await orchestrator.RunAsync("conversation", "hello", [], includeCuaTool: false);

        Assert.Equal("hello", result);
        Assert.Empty(model.Requests[0].GetProperty("tools").EnumerateArray());
    }

    [Fact]
    public async Task ComputerUseWithoutRequestContextFailsBeforeModelInvocation()
    {
        var model = new ScriptedModel();
        using var httpClient = new HttpClient();
        var orchestrator = CreateOrchestrator(model, httpClient);

        var exception = await Assert.ThrowsAsync<InvalidOperationException>(() =>
            orchestrator.RunAsync("conversation", "use desktop", []));

        Assert.Equal("W365 session request context is required for computer use.", exception.Message);
        Assert.Empty(model.Requests);
    }

    private static ComputerUseOrchestrator CreateOrchestrator(ICuaModelProvider model, HttpClient httpClient) =>
        new(model, new TestHttpClientFactory(httpClient), new ConfigurationBuilder().Build(), NullLogger<ComputerUseOrchestrator>.Instance);

    private static W365StartSessionRequest Request(string oid, int width, int height)
    {
        var token = new JwtSecurityTokenHandler().WriteToken(new JwtSecurityToken(claims: [new Claim("oid", oid)]));
        return W365StartSessionRequest.FromActivityValue(
            new { sessionKind = "local", sessionProperties = new { displayWidth = width, displayHeight = height, browserMode = "chromium" } },
            token);
    }

    private static void AssertStartPayload(JsonElement arguments, string oid, int width, int height)
    {
        Assert.Equal(new[] { "sessionKind", "sessionProperties", "userObjectId" },
            arguments.EnumerateObject().Select(property => property.Name).Order());
        Assert.Equal("local", arguments.GetProperty("sessionKind").GetString());
        Assert.Equal(oid, arguments.GetProperty("userObjectId").GetString());
        var properties = arguments.GetProperty("sessionProperties");
        Assert.Equal(3, properties.EnumerateObject().Count());
        Assert.Equal(width, properties.GetProperty("displayWidth").GetInt32());
        Assert.Equal(height, properties.GetProperty("displayHeight").GetInt32());
        Assert.Equal("chromium", properties.GetProperty("browserMode").GetString());
    }

    private sealed class TestHttpClientFactory(HttpClient httpClient) : IHttpClientFactory
    {
        public HttpClient CreateClient(string name) => httpClient;
    }

    private sealed class ScriptedModel(params string[] responses) : ICuaModelProvider
    {
        private readonly Queue<string> _responses = new(responses);
        public string ModelName => "computer-use-preview";
        public List<JsonElement> Requests { get; } = [];

        public Task<string> SendAsync(string requestBody, CancellationToken cancellationToken)
        {
            cancellationToken.ThrowIfCancellationRequested();
            Requests.Add(JsonSerializer.Deserialize<JsonElement>(requestBody));
            return Task.FromResult(_responses.Dequeue());
        }
    }
}
