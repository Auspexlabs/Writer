namespace Writer.Core;

public enum ErrorCode
{
    Usage,
    FileNotFound,
    PathNotFound,
    PathAmbiguous,
    Validation,
    UnsupportedKind,
    FormatReadonly,
    FormatError,
    UnknownFormat,
    Io,
    Conflict,
    Internal,
}

/// <summary>The one exception type that reaches users. Every instance carries a next-step hint for the agent.</summary>
public sealed class WriterException(ErrorCode code, string message, string hint) : Exception(message)
{
    public ErrorCode Code { get; } = code;
    public string Hint { get; } = hint;

    public string CodeName => Code switch
    {
        ErrorCode.Usage => "USAGE",
        ErrorCode.FileNotFound => "FILE_NOT_FOUND",
        ErrorCode.PathNotFound => "PATH_NOT_FOUND",
        ErrorCode.PathAmbiguous => "PATH_AMBIGUOUS",
        ErrorCode.Validation => "VALIDATION",
        ErrorCode.UnsupportedKind => "UNSUPPORTED_KIND",
        ErrorCode.FormatReadonly => "FORMAT_READONLY",
        ErrorCode.FormatError => "FORMAT_ERROR",
        ErrorCode.UnknownFormat => "UNKNOWN_FORMAT",
        ErrorCode.Io => "IO",
        ErrorCode.Conflict => "CONFLICT",
        _ => "INTERNAL",
    };

    public int ExitCode => Code switch
    {
        ErrorCode.Usage => 1,
        ErrorCode.FileNotFound or ErrorCode.PathNotFound or ErrorCode.PathAmbiguous => 2,
        ErrorCode.Validation or ErrorCode.UnsupportedKind or ErrorCode.FormatReadonly => 3,
        ErrorCode.FormatError or ErrorCode.UnknownFormat => 4,
        ErrorCode.Io => 5,
        ErrorCode.Conflict => 6,
        _ => 70,
    };
}
