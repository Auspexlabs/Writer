using System.ComponentModel;
using System.Text;
using ModelContextProtocol.Protocol;
using ModelContextProtocol.Server;
using Writer.Core;

namespace Writer.Cli;

/// <summary>`writer mcp`: a stdio MCP server exposing one tool that runs writer commands.
/// One contract for agents and the CLI: the tool's output is exactly what the command line prints.</summary>
public static class Mcp
{
    public const string ToolName = Assistant.ToolName;
    public const string ToolDescription = Assistant.ToolDescription;
    public const string Instructions = Assistant.Instructions;

    public static async Task ServeAsync(CancellationToken cancellationToken = default)
    {
        var options = new McpServerOptions
        {
            ServerInfo = new Implementation { Name = "writer", Version = Runner.Version },
            ServerInstructions = Instructions,
            ToolCollection = new McpServerPrimitiveCollection<McpServerTool>
            {
                McpServerTool.Create((Func<string, CallToolResult>)Call, new McpServerToolCreateOptions { Name = ToolName, Description = ToolDescription }),
            },
        };
        await using var transport = new StdioServerTransport("writer", null);
        await using var server = McpServer.Create(transport, options);
        await server.RunAsync(cancellationToken);
    }

    /// <summary>Runs one command line and packages its output as a tool result.</summary>
    public static CallToolResult Call([Description("The writer command line, without the program name.")] string command)
    {
        var stdout = new StringWriter();
        var stderr = new StringWriter();
        int code;
        try
        {
            var argv = Tokenize(command);
            if (argv.Length > 0 && argv[0] == "writer") argv = argv[1..];
            if (argv.Length > 0 && argv[0] == "mcp")
                throw new WriterException(ErrorCode.Usage, "The server is already running", "Run a document command instead.");
            code = Runner.Run(argv, stdout, stderr);
        }
        catch (WriterException ex)
        {
            stderr.WriteLine(Runner.ErrorJson(ex));
            code = ex.ExitCode;
        }
        var text = code == 0 ? stdout.ToString() : stderr.ToString();
        return new CallToolResult { Content = [new TextContentBlock { Text = text.TrimEnd('\n') }], IsError = code != 0 };
    }

    /// <summary>Shell-style splitting (Args.Tokenize), kept here for the callers that know it by this name.</summary>
    public static string[] Tokenize(string command) => Args.Tokenize(command);
}
