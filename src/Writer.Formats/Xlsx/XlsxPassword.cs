using System.Buffers.Binary;
using System.Security.Cryptography;
using System.Text;
using System.Text.Json.Nodes;
using Writer.Core;

namespace Writer.Formats.Xlsx;

/// <summary>Worksheet editing passwords (not package encryption). Plaintext is
/// used only for this calculation and never becomes part of a document/history.</summary>
public static class XlsxPassword
{
    static WriterException Invalid() => new(ErrorCode.Validation, "Invalid worksheet password verifier", "Use a supported SHA password verifier or a legacy Excel password hash.");
    public static JsonObject Create(string password)
    {
        CheckPassword(password);
        var salt = RandomNumberGenerator.GetBytes(16);
        return new JsonObject { ["algorithmName"] = "SHA-512", ["saltValue"] = Convert.ToBase64String(salt), ["spinCount"] = 100000,
            ["hashValue"] = Convert.ToBase64String(Hash(password, "SHA-512", salt, 100000)) };
    }

    public static bool Verify(string password, JsonObject verifier)
    {
        CheckPassword(password);
        Validate(verifier);
        if (verifier["hashValue"] is not null)
            return CryptographicOperations.FixedTimeEquals(Convert.FromBase64String(verifier["hashValue"]!.GetValue<string>()),
                Hash(password, verifier["algorithmName"]!.GetValue<string>(), Convert.FromBase64String(verifier["saltValue"]!.GetValue<string>()), verifier["spinCount"]!.GetValue<int>()));
        return ushort.Parse(verifier["password"]!.GetValue<string>(), System.Globalization.NumberStyles.HexNumber) == Legacy(password);
    }

    public static void Validate(JsonObject verifier)
    {
        try
        {
            if (verifier.Any(p => p.Key is not ("password" or "algorithmName" or "hashValue" or "saltValue" or "spinCount"))) throw Invalid();
            if (verifier["password"] is { } legacy && (legacy.GetValue<string>().Length != 4 || !ushort.TryParse(legacy.GetValue<string>(), System.Globalization.NumberStyles.HexNumber, null, out _))) throw Invalid();
            if (verifier["hashValue"] is null)
            {
                if (verifier.Count != 1 || verifier["password"] is null) throw Invalid();
                return;
            }
            var algorithm = verifier["algorithmName"]!.GetValue<string>();
            var size = algorithm switch { "SHA-1" => 20, "SHA-256" => 32, "SHA-384" => 48, "SHA-512" => 64, _ => throw Invalid() };
            if (Convert.FromBase64String(verifier["hashValue"]!.GetValue<string>()).Length != size ||
                Convert.FromBase64String(verifier["saltValue"]!.GetValue<string>()).Length is < 1 or > 1024 ||
                verifier["spinCount"]!.GetValue<int>() is < 0 or > 10000000) throw Invalid();
        }
        catch (Exception e) when (e is not WriterException) { throw Invalid(); }
    }

    static void CheckPassword(string password)
    {
        if (password.Length > 255) throw new WriterException(ErrorCode.Validation, "Worksheet password is too long", "Use at most 255 characters.");
    }
    // MS-OFFCRYPTO 2.4.2.4: salt + UTF-16LE; then hash + little-endian iteration.
    static byte[] Hash(string password, string algorithm, byte[] salt, int count)
    {
        using var hash = IncrementalHash.CreateHash(new HashAlgorithmName(algorithm.Replace("-", "")));
        var input = Encoding.Unicode.GetBytes(password);
        try
        {
            hash.AppendData(salt); hash.AppendData(input);
            var value = hash.GetHashAndReset();
            Span<byte> iteration = stackalloc byte[4];
            for (var i = 0; i < count; i++)
            {
                BinaryPrimitives.WriteInt32LittleEndian(iteration, i);
                hash.AppendData(value); hash.AppendData(iteration);
                hash.GetHashAndReset(value);
            }
            return value;
        }
        finally { CryptographicOperations.ZeroMemory(input); }
    }
    // MS-OFFCRYPTO 2.3.7.1: rotate the 15-bit verifier; one byte per UTF-16 unit.
    static ushort Legacy(string password)
    {
        var result = 0;
        for (var i = password.Length - 1; i >= 0; i--)
        {
            var c = password[i]; var b = (c & 255) == 0 ? c >> 8 : c & 255;
            result = ((result << 1) & 0x7FFF | result >> 14 & 1) ^ b;
        }
        return (ushort)(((result << 1) & 0x7FFF | result >> 14 & 1) ^ password.Length ^ 0xCE4B);
    }
}
