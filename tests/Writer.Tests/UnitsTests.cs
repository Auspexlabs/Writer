using Writer.Core;

namespace Writer.Tests;

public class UnitsTests
{
    [Theory]
    [InlineData("1in", 914400)]
    [InlineData("2cm", 720000)]
    [InlineData("20mm", 720000)]
    [InlineData("72pt", 914400)]
    [InlineData("96px", 914400)]
    [InlineData("720000", 720000)]
    [InlineData("720000emu", 720000)]
    [InlineData(" 2.5 CM ", 900000)]
    [InlineData("-1cm", -360000)]
    public void ParseLength_accepts_every_unit(string input, long emu) => Assert.Equal(emu, Units.ParseLength(input));

    [Theory]
    [InlineData("abc")]
    [InlineData("")]
    [InlineData("2 lightyears")]
    public void ParseLength_rejects_garbage(string input)
    {
        var ex = Assert.Throws<WriterException>(() => Units.ParseLength(input));
        Assert.Equal(ErrorCode.Validation, ex.Code);
        Assert.Contains("2cm", ex.Hint);
    }

    [Theory]
    [InlineData(720000, "2cm")]
    [InlineData(914400, "2.54cm")]
    [InlineData(0, "0cm")]
    [InlineData(152400, "0.423cm")]
    public void FormatLength_prints_centimetres(long emu, string expected) => Assert.Equal(expected, Units.FormatLength(emu));

    [Fact]
    public void Length_round_trips_through_display() => Assert.Equal(720000, Units.ParseLength(Units.FormatLength(720000)));

    [Theory]
    [InlineData("12", "12")]
    [InlineData("12pt", "12")]
    [InlineData("10.5PT", "10.5")]
    [InlineData("10.554", "10.55")]
    public void ParsePoints_normalises(string input, string canonical) => Assert.Equal(canonical, Units.ParsePoints(input));

    [Theory]
    [InlineData("0")]
    [InlineData("-3")]
    [InlineData("big")]
    public void ParsePoints_rejects(string input) =>
        Assert.Equal(ErrorCode.Validation, Assert.Throws<WriterException>(() => Units.ParsePoints(input)).Code);

    [Fact]
    public void FormatPoints_appends_unit() => Assert.Equal("12pt", Units.FormatPoints("12"));

    [Theory]
    [InlineData("FF0000", "FF0000")]
    [InlineData("#ff0000", "FF0000")]
    [InlineData("red", "FF0000")]
    [InlineData("Navy", "000080")]
    public void ParseColor_normalises(string input, string canonical) => Assert.Equal(canonical, Units.ParseColor(input));

    [Theory]
    [InlineData("#12345")]
    [InlineData("GGGGGG")]
    [InlineData("reddish")]
    public void ParseColor_rejects(string input) =>
        Assert.Equal(ErrorCode.Validation, Assert.Throws<WriterException>(() => Units.ParseColor(input)).Code);
}

public class ErrorTests
{
    [Theory]
    [InlineData(ErrorCode.Usage, "USAGE", 1)]
    [InlineData(ErrorCode.PathNotFound, "PATH_NOT_FOUND", 2)]
    [InlineData(ErrorCode.Validation, "VALIDATION", 3)]
    [InlineData(ErrorCode.FormatError, "FORMAT_ERROR", 4)]
    [InlineData(ErrorCode.Io, "IO", 5)]
    [InlineData(ErrorCode.Internal, "INTERNAL", 70)]
    public void Codes_map_to_names_and_exit_codes(ErrorCode code, string name, int exit)
    {
        var ex = new WriterException(code, "m", "h");
        Assert.Equal(name, ex.CodeName);
        Assert.Equal(exit, ex.ExitCode);
        Assert.Equal("m", ex.Message);
        Assert.Equal("h", ex.Hint);
    }
}
