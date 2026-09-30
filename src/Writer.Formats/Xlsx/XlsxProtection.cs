using System.Text.Json.Nodes;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Spreadsheet;
using Writer.Core;

namespace Writer.Formats.Xlsx;

static class XlsxProtection
{
    // OOXML flags mean forbidden; the editor exposes positive permissions.
    static readonly string[] Flags = ["formatCells", "formatColumns", "formatRows", "insertColumns", "insertRows", "insertHyperlinks", "deleteColumns", "deleteRows", "sort", "autoFilter", "pivotTables", "objects", "scenarios", "selectLockedCells", "selectUnlockedCells"];
    internal static string Read(Worksheet sheet)
    {
        var result = new JsonObject();
        if (sheet.GetFirstChild<SheetProtection>() is not { } protection) return result.ToJsonString();
        foreach (var flag in Flags)
        {
            var value = protection.GetAttributes().FirstOrDefault(a => a.LocalName == flag && a.NamespaceUri == "").Value;
            var defaultBlocked = flag is not ("objects" or "scenarios" or "selectLockedCells" or "selectUnlockedCells");
            result[flag] = value is null or "" ? !defaultBlocked : value is "0" or "false";
        }
        result["passwordProtected"] = protection.Password is not null || protection.HashValue is not null;
        var verifier = new JsonObject();
        foreach (var key in new[] { "password", "algorithmName", "hashValue", "saltValue", "spinCount" })
        {
            var value = protection.GetAttributes().FirstOrDefault(a => a.LocalName == key && a.NamespaceUri == "").Value;
            if (!string.IsNullOrEmpty(value)) verifier[key] = key == "spinCount" ? JsonValue.Create(int.Parse(value, System.Globalization.CultureInfo.InvariantCulture)) : JsonValue.Create(value);
        }
        if (verifier.Count > 0) result["verifier"] = verifier;
        return result.ToJsonString();
    }

    internal static void Write(Worksheet sheet, string json)
    {
        var input = JsonNode.Parse(json) as JsonObject ?? throw new WriterException(ErrorCode.Validation, "Protection permissions must be an object", "Use permission names with boolean values.");
        foreach (var (key, value) in input)
            if (key != "verifier" && ((!Flags.Contains(key) && key != "passwordProtected") || value is not JsonValue j || !j.TryGetValue<bool>(out _)))
                throw new WriterException(ErrorCode.Validation, "Invalid protection permission: " + key, "Use boolean permission values.");
        var sameVerifier = JsonNode.DeepEquals(input["verifier"], JsonNode.Parse(Read(sheet))?["verifier"]);
        if (input["verifier"] is { } verifier && !sameVerifier)
        {
            if (verifier is not JsonObject o) throw new WriterException(ErrorCode.Validation, "Invalid worksheet password verifier", "Use a verifier object.");
            XlsxPassword.Validate(o);
        }
        var protection = sheet.GetFirstChild<SheetProtection>() ?? new SheetProtection { Sheet = false, Objects = true, Scenarios = true };
        if (input.ContainsKey("verifier") && !sameVerifier)
        {
            foreach (var key in new[] { "password", "algorithmName", "hashValue", "saltValue", "spinCount" }) protection.RemoveAttribute(key, "");
            if (input["verifier"] is JsonObject values)
                foreach (var (key, value) in values) protection.SetAttribute(new OpenXmlAttribute(key, "", value!.ToString()));
        }
        // Keep password hashes and flags not edited by the client, including future extensions.
        foreach (var flag in Flags)
            if (input[flag] is { } value) protection.SetAttribute(new OpenXmlAttribute(flag, "", value.GetValue<bool>() ? "0" : "1"));
        if (protection.Parent is null) sheet.AddChild(protection, true);
    }
}
