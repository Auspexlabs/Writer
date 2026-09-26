using Writer.Core;
using Writer.Formats;

namespace Writer.Cli;

/// <summary>Entry point shared by the executable, the tests and the MCP server.</summary>
public static class Runner
{
    public const string Version = "0.1.5";

    /// <summary>Runs one command. Output goes to stdout, errors as JSON to stderr; returns the exit code.</summary>
    public static int Run(string[] argv, TextWriter stdout, TextWriter stderr)
    {
        var debug = argv.Contains("--debug");
        try
        {
            stdout.Write(Commands.Execute(Args.Parse(argv)));
            return 0;
        }
        catch (WriterException ex)
        {
            stderr.WriteLine(ErrorJson(ex));
            if (debug) stderr.WriteLine(ex.ToString());
            return ex.ExitCode;
        }
        catch (Exception ex)
        {
            var wrapped = new WriterException(ErrorCode.Internal, ex.Message, "This is a bug in writer. Re-run with --debug and report the stack trace.");
            stderr.WriteLine(ErrorJson(wrapped));
            if (debug) stderr.WriteLine(ex.ToString());
            return wrapped.ExitCode;
        }
    }

    public static string ErrorJson(WriterException ex) => NodeJson.Write(w =>
    {
        w.WriteStartObject();
        w.WriteStartObject("error");
        w.WriteString("code", ex.CodeName);
        w.WriteString("message", ex.Message);
        w.WriteString("hint", ex.Hint);
        w.WriteEndObject();
        w.WriteEndObject();
    });
}

/// <summary>Opening and saving files with the right error codes. Saves are atomic.</summary>
static class Files
{
    public static Document Open(string path)
    {
        if (!File.Exists(path))
            throw new WriterException(ErrorCode.FileNotFound, $"{path} not found", "Check the path, or create the file with 'writer create'.");
        var adapter = Adapters.ForPath(path);
        try
        {
            using var stream = File.OpenRead(path);
            var doc = adapter.Open(stream);
            doc.SourcePath = path;
            return doc;
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            throw new WriterException(ErrorCode.Io, ex.Message, "Is the file open in another program, or the folder unreadable?");
        }
    }

    /// <summary>Read-only formats fail early with one clear error instead of a property-level one.</summary>
    public static void EnsureWritable(Document doc)
    {
        if (Adapters.ForName(doc.Format) is { CanWrite: false } adapter)
            throw new WriterException(ErrorCode.FormatReadonly, $"{doc.Format} files are read-only",
                $"Export first, then edit the copy: writer export <file> --to <file>.{(adapter as Writer.Formats.Compat.CompatAdapter)?.Target ?? "md"}");
    }

    /// <summary>Replaces the file atomically (temp file beside it, then rename). Where the folder is not writable but the file is
    /// (the App Store sandbox grants a chosen file, not its folder), the new bytes are built in the temp folder and copied over it.</summary>
    public static void ReplaceAtomic(string path, Action<Stream> write)
    {
        var full = Path.GetFullPath(path);
        var tmp = Path.Combine(Path.GetDirectoryName(full)!, $".{Path.GetFileName(full)}.{Guid.NewGuid():N}.tmp");
        try
        {
            try
            {
                using (var stream = File.Create(tmp)) write(stream);
            }
            catch (UnauthorizedAccessException) when (File.Exists(full))
            {
                tmp = Path.Combine(Path.GetTempPath(), $"writer-{Guid.NewGuid():N}.tmp");
                using (var stream = File.Create(tmp)) write(stream);
                File.Copy(tmp, full, overwrite: true);
                File.Delete(tmp);
                return;
            }
            File.Move(tmp, full, overwrite: true);
        }
        catch (Exception ex) when (ex is IOException or UnauthorizedAccessException)
        {
            if (File.Exists(tmp)) File.Delete(tmp);
            throw new WriterException(ErrorCode.Io, ex.Message, "Check that the folder is writable and the file is not open in another program.");
        }
    }

    public static void WriteTextAtomic(string path, string text) => ReplaceAtomic(path, s => s.Write(new System.Text.UTF8Encoding(false).GetBytes(text)));
    public static void SaveAtomic(Document doc, string path) => ReplaceAtomic(path, doc.Save);
}
