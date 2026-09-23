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
    public const string ToolName = "writer";

    public const string ToolDescription =
        "Run a writer command against a document (docx, xlsx, pptx, md read/write; pdf read). "
        + "Same syntax as the CLI without the program name, e.g. \"view report.docx outline\", "
        + "\"set report.docx /body/paragraph[2] --prop text=\\\"New text\\\"\", \"help docx paragraph\". "
        + "Returns the command's JSON or text. Start with \"help\" and \"view <file> outline\".";

    public const string Instructions =
        "Workflow: `view <file> outline` to see every element with its path, `query` or `get` to inspect, "
        + "`add`/`set`/`remove`/`move`/`copy` to change, `view <file> html` to check the result, `export` to convert. "
        + "When unsure about a property, run `help <format> <element>` instead of guessing.";

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

    /// <summary>Shell-style splitting: whitespace separates, single and double quotes group, backslash escapes inside double quotes.</summary>
    public static string[] Tokenize(string command)
    {
        var args = new List<string>();
        var current = new StringBuilder();
        var pending = false;
        char? quote = null;
        for (var i = 0; i < command.Length; i++)
        {
            var c = command[i];
            if (quote is null)
            {
                if (char.IsWhiteSpace(c))
                {
                    if (pending) args.Add(current.ToString());
                    current.Clear();
                    pending = false;
                    continue;
                }
                if (c is '"' or '\'')
                {
                    quote = c;
                    pending = true;
                    continue;
                }
                if (c == '\\' && i + 1 < command.Length)
                {
                    current.Append(command[++i]);
                    pending = true;
                    continue;
                }
                current.Append(c);
                pending = true;
                continue;
            }
            if (c == quote)
            {
                quote = null;
                continue;
            }
            if (quote == '"' && c == '\\' && i + 1 < command.Length && command[i + 1] is '"' or '\\')
            {
                current.Append(command[++i]);
                continue;
            }
            current.Append(c);
        }
        if (quote is not null)
            throw new WriterException(ErrorCode.Usage, "Unterminated quote in command", "Close the quote, e.g. --prop text=\"Hello\"");
        if (pending) args.Add(current.ToString());
        return args.ToArray();
    }
}
