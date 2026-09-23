using System.Buffers.Binary;
using Writer.Core;

namespace Writer.Formats.Common;

/// <summary>Pixel size and content type read from the header of a PNG, JPEG, GIF or BMP file.</summary>
public sealed record ImageInfo(int Width, int Height, string ContentType, string Extension)
{
    /// <summary>Image bytes and a file name from a path or a data: URL.</summary>
    public static (byte[] Bytes, string Name) Load(string src)
    {
        if (src.StartsWith("data:", StringComparison.OrdinalIgnoreCase))
        {
            var comma = src.IndexOf(',');
            if (comma < 0) throw new WriterException(ErrorCode.Validation, "Malformed data URL", "Use data:image/png;base64,....");
            var header = src[5..comma];
            byte[] bytes;
            try
            {
                bytes = header.EndsWith(";base64", StringComparison.OrdinalIgnoreCase)
                    ? Convert.FromBase64String(src[(comma + 1)..])
                    : System.Text.Encoding.Latin1.GetBytes(Uri.UnescapeDataString(src[(comma + 1)..]));
            }
            catch (FormatException)
            {
                throw new WriterException(ErrorCode.Validation, "Malformed data URL", "The base64 payload is not valid.");
            }
            return (bytes, "image." + Read(bytes).Extension);
        }
        if (!File.Exists(src))
            throw new WriterException(ErrorCode.FileNotFound, $"Image {src} not found", "Pass the path of a PNG, JPEG, GIF or BMP file, or a data: URL.");
        return (File.ReadAllBytes(src), Path.GetFileName(src));
    }

    public static ImageInfo Read(byte[] bytes)
    {
        var s = bytes.AsSpan();
        if (s.Length >= 24 && s[..8].SequenceEqual((ReadOnlySpan<byte>)[0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A]))
            return new(BinaryPrimitives.ReadInt32BigEndian(s[16..]), BinaryPrimitives.ReadInt32BigEndian(s[20..]), "image/png", "png");
        if (s.Length >= 10 && s[..4].SequenceEqual("GIF8"u8))
            return new(BinaryPrimitives.ReadUInt16LittleEndian(s[6..]), BinaryPrimitives.ReadUInt16LittleEndian(s[8..]), "image/gif", "gif");
        if (s.Length >= 26 && s[0] == 'B' && s[1] == 'M')
            return new(BinaryPrimitives.ReadInt32LittleEndian(s[18..]), Math.Abs(BinaryPrimitives.ReadInt32LittleEndian(s[22..])), "image/bmp", "bmp");
        if (s.Length >= 4 && s[0] == 0xFF && s[1] == 0xD8)
        {
            var i = 2;
            while (i + 9 < s.Length && s[i] == 0xFF)
            {
                var marker = s[i + 1];
                var length = BinaryPrimitives.ReadUInt16BigEndian(s[(i + 2)..]);
                if (marker is >= 0xC0 and <= 0xCF and not 0xC4 and not 0xC8 and not 0xCC)
                    return new(BinaryPrimitives.ReadUInt16BigEndian(s[(i + 7)..]), BinaryPrimitives.ReadUInt16BigEndian(s[(i + 5)..]), "image/jpeg", "jpeg");
                i += 2 + length;
            }
        }
        throw new WriterException(ErrorCode.Validation, "Unsupported image data", "Use a PNG, JPEG, GIF or BMP file.");
    }
}
