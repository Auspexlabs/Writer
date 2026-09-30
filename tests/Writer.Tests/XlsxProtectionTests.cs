using System.Text.Json.Nodes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Validation;
using Writer.Core;
using Writer.Formats.Xlsx;

namespace Writer.Tests;

public class XlsxProtectionTests
{
    [Fact]
    public void Split_panes_roundtrip_and_switch_to_freeze_without_losing_sheet_content()
    {
        using var doc=new XlsxAdapter().Create();
        Set(doc,"/sheet[1]/cell[D20]",("value","keep"));
        Set(doc,"/sheet[1]",("split","{\"x\":3600,\"y\":2400,\"topLeft\":\"D20\",\"activePane\":\"bottomRight\"}"));
        var bytes=Save(doc);using(var package=SpreadsheetDocument.Open(new MemoryStream(bytes),false))Assert.Empty(new OpenXmlValidator().Validate(package));
        using var reopened=new XlsxAdapter().Open(new MemoryStream(bytes));
        var node=PathResolver.Single(reopened.Root,"/sheet[1]");var split=JsonNode.Parse(node.GetProps()["split"])!;
        Assert.Equal(3600,split["x"]!.GetValue<double>());Assert.Equal("D20",split["topLeft"]!.GetValue<string>());
        Assert.False(node.GetProps().ContainsKey("freeze"));
        Set(reopened,"/sheet[1]",("freeze","C3"),("split","null"));
        Assert.Equal("C3",node.GetProps()["freeze"]);Assert.Equal("null",node.GetProps()["split"]);
        Set(reopened,"/sheet[1]",("freeze","none"),("split",split.ToJsonString()));
        var before=node.GetRaw();Assert.ThrowsAny<Exception>(()=>Set(reopened,"/sheet[1]",("split","{\"x\":-1}")));Assert.Equal(before,node.GetRaw());
        Set(reopened,"/sheet[1]",("split","null"));Assert.Equal("null",node.GetProps()["split"]);
        Assert.Equal("keep",PathResolver.Single(reopened.Root,"/sheet[1]/cell[D20]").GetProps()["value"]);
    }
    [Fact]
    public void Workbook_structure_password_roundtrips_without_touching_revision_restrictions()
    {
        using var doc=new XlsxAdapter().Create();
        var book=((XlsxDocument)doc).Workbook.Workbook!;
        book.AddChild(new WorkbookProtection { LockRevision=true, RevisionsPassword="83AF" },true);
        var verifier=XlsxPassword.Create("structure secret");
        var settings=new JsonObject{["structure"]=true,["verifier"]=verifier}.ToJsonString();
        Set(doc,"/",("workbookProtection",settings));
        var bytes=Save(doc);
        using(var package=SpreadsheetDocument.Open(new MemoryStream(bytes),false))
        {
            Assert.Empty(new OpenXmlValidator().Validate(package));
            var protection=package.WorkbookPart!.Workbook!.GetFirstChild<WorkbookProtection>()!;
            Assert.True(protection.LockStructure!.Value);Assert.True(protection.LockRevision!.Value);
            Assert.Equal("83AF",protection.RevisionsPassword!.Value);
            Assert.Equal("SHA-512",protection.WorkbookAlgorithmName!.Value);
            Assert.DoesNotContain("structure secret",protection.OuterXml);
        }
        using var reopened=new XlsxAdapter().Open(new MemoryStream(bytes));
        var result=JsonNode.Parse(reopened.Root.GetProps()["workbookProtection"])!;
        Assert.True(XlsxPassword.Verify("structure secret",result["verifier"]!.AsObject()));
        Set(reopened,"/",("workbookProtection","{\"structure\":false,\"verifier\":null}"));
        Assert.False(JsonNode.Parse(reopened.Root.GetProps()["workbookProtection"])!["structure"]!.GetValue<bool>());
        Save(reopened);Set(reopened,"/",("workbookProtection",settings));Save(reopened);
        Assert.Equal(verifier["hashValue"]!.ToString(),JsonNode.Parse(reopened.Root.GetProps()["workbookProtection"])!["verifier"]!["hashValue"]!.ToString());
        var prior=reopened.Root.GetProps()["workbookProtection"];
        Assert.ThrowsAny<Exception>(()=>Set(reopened,"/",("workbookProtection","{\"structure\":false,\"verifier\":{\"password\":\"oops\"}}")));
        Assert.Equal(prior,reopened.Root.GetProps()["workbookProtection"]);
    }
    static void Set(Document doc, string path, params (string Key,string Value)[] props) => Mutations.Set(PathResolver.Single(doc.Root,path),props.ToDictionary(x=>x.Key,x=>x.Value));
    static byte[] Save(Document doc) { using var stream=new MemoryStream();doc.Save(stream);return stream.ToArray(); }
    [Theory]
    [InlineData("SHA-1","KHYfIQzXAm53ee6kiJ4jSy/nJWk=")]
    [InlineData("SHA-256","3gTtva52LDNiyAJVfECE5Ai1HEp2nnKkAVGccB9ebE0=")]
    [InlineData("SHA-384","oGKy4GD3kFDTfxBanQarJgQWIesIPKh7AkiO2+x6prql4srlp+zis6HWLDhU/tFY")]
    [InlineData("SHA-512","H1LYurL43cmkvd1Ujc2IDzw2+zz00ZmkgDwZSdRlmioHIv+tyxN9XVhTgo90s1e7Hq3QsH0iKUSgfm1BvwF2uA==")]
    public void Password_hashes_match_independent_ISO_vectors(string algorithm,string hash)
    {
        // Independently calculated with Python hashlib, UTF-16LE and struct.pack('<I', i).
        var v=new JsonObject { ["algorithmName"]=algorithm,["hashValue"]=hash,["saltValue"]="AAECAwQFBgcICQoLDA0ODw==",["spinCount"]=1000 };
        Assert.True(XlsxPassword.Verify("工作表🔑",v));Assert.False(XlsxPassword.Verify("工作表",v));
    }
    [Fact]
    public void Password_creation_roundtrips_and_saved_undo_restores_the_verifier()
    {
        Assert.True(XlsxPassword.Verify("password",new JsonObject{["password"]="83AF"}));
        Assert.False(XlsxPassword.Verify("wrong",new JsonObject{["password"]="83AF"}));
        var v=XlsxPassword.Create("中文 password");Assert.True(XlsxPassword.Verify("中文 password",v));
        Assert.NotEqual(v["saltValue"]!.ToString(),XlsxPassword.Create("中文 password")["saltValue"]!.ToString());
        using var doc=new XlsxAdapter().Create();
        Set(doc,"/sheet[1]",("protected","true"),("protection",new JsonObject{["sort"]=true,["verifier"]=v}.ToJsonString()));
        using var reopened=new XlsxAdapter().Open(new MemoryStream(Save(doc)));
        var saved=PathResolver.Single(reopened.Root,"/sheet[1]").GetProps()["protection"];
        Assert.True(XlsxPassword.Verify("中文 password",JsonNode.Parse(saved)!["verifier"]!.AsObject()));
        Set(reopened,"/sheet[1]",("protected","false"),("protection","{\"verifier\":null}"));
        Assert.False(JsonNode.Parse(PathResolver.Single(reopened.Root,"/sheet[1]").GetProps()["protection"])!["passwordProtected"]!.GetValue<bool>());
        Save(reopened);Set(reopened,"/sheet[1]",("protected","true"),("protection",saved));
        var bytes=Save(reopened);using var package=SpreadsheetDocument.Open(new MemoryStream(bytes),false);
        Assert.Empty(new OpenXmlValidator().Validate(package));
        Assert.Equal(v["hashValue"]!.ToString(),package.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetProtection>()!.HashValue!.Value);
        Assert.DoesNotContain("中文 password",package.WorkbookPart.WorksheetParts.First().Worksheet!.OuterXml);
        var before=PathResolver.Single(reopened.Root,"/sheet[1]").GetRaw();
        Assert.ThrowsAny<Exception>(()=>Set(reopened,"/sheet[1]",("protection","{\"sort\":false,\"verifier\":{\"password\":\"oops\"}}")));
        Assert.Equal(before,PathResolver.Single(reopened.Root,"/sheet[1]").GetRaw());
        v["spinCount"]=10000001;Assert.ThrowsAny<Exception>(()=>XlsxPassword.Verify("x",v));
    }
    [Fact]
    public void Row_column_defaults_inherit_and_shift_without_materializing_cells_or_losing_style()
    {
        using var doc=new XlsxAdapter().Create();
        Set(doc,"/sheet[1]/cell[B2]",("locked","false"),("formulaHidden","true"),("bold","true"),("fill","00FF00"),("format","0.000"));
        var ws=((XlsxDocument)doc).Workbook.WorksheetParts.First().Worksheet!;
        var cell=ws.Descendants<Cell>().Single();var style=cell.StyleIndex!.Value;
        ws.AddChild(new Columns(new Column{Min=2,Max=2,Style=style,Width=18,CustomWidth=true}),true);
        var row=ws.Descendants<Row>().Single();row.StyleIndex=0;row.CustomFormat=true;
        cell.StyleIndex=null; // row style overrides column style, even when it is the normal style
        Assert.False(PathResolver.Single(doc.Root,"/sheet[1]/cell[B2]").GetProps().ContainsKey("locked"));
        Set(doc,"/sheet[1]/cell[B3]",("value","7"));
        var b3=PathResolver.Single(doc.Root,"/sheet[1]/cell[B3]").GetProps();
        Assert.Equal("false",b3["locked"]);Assert.Equal("true",b3["bold"]);Assert.Equal("0.000",b3["format"]);
        var snapshot=JsonNode.Parse(PathResolver.Single(doc.Root,"/sheet[1]").GetProps()["rowColumnStyles"])!.AsObject();
        var original=snapshot.ToJsonString();snapshot["cols"]!["C"]=snapshot["cols"]!["B"]!.DeepClone();snapshot["cols"]!.AsObject().Remove("B");
        snapshot["rows"]!["4"]=snapshot["rows"]!["2"]!.DeepClone();snapshot["rows"]!.AsObject().Remove("2");
        Set(doc,"/sheet[1]",("rowColumnStyles",snapshot.ToJsonString()));
        var bytes=Save(doc);using(var package=SpreadsheetDocument.Open(new MemoryStream(bytes),false)){Assert.Empty(new OpenXmlValidator().Validate(package));Assert.Equal(2,package.WorkbookPart!.WorksheetParts.First().Worksheet!.Descendants<Cell>().Count());}
        using var reopened=new XlsxAdapter().Open(new MemoryStream(bytes));
        Set(reopened,"/sheet[1]/cell[C1000000]",("value","4"));
        Assert.Equal("false",PathResolver.Single(reopened.Root,"/sheet[1]/cell[C1000000]").GetProps()["locked"]);
        Set(reopened,"/sheet[1]",("rowColumnStyles",original));Save(reopened); // saved undo restores definitions after style trimming
        Set(reopened,"/sheet[1]/cell[B900000]",("value","9"));
        Assert.Equal("true",PathResolver.Single(reopened.Root,"/sheet[1]/cell[B900000]").GetProps()["bold"]);
        Assert.Equal("false",PathResolver.Single(reopened.Root,"/sheet[1]/cell[B900000]").GetProps()["locked"]);
        Assert.False(PathResolver.Single(reopened.Root,"/sheet[1]/cell[C1000000]").GetProps().ContainsKey("locked"));
    }
    [Fact]
    public void Cell_locks_hidden_formulas_and_sheet_permissions_roundtrip_without_losing_styles()
    {
        using var doc=new XlsxAdapter().Create();
        Set(doc,"/sheet[1]/cell[A1]",("formula","1+2"),("locked","false"),("formulaHidden","true"),("bold","true"));
        Set(doc,"/sheet[1]/range[B1:B3]",("locked","false"));
        Set(doc,"/sheet[1]",("protected","true"),("protection","{\"formatCells\":true,\"sort\":true,\"autoFilter\":true,\"formatRows\":false,\"objects\":false}"));
        var bytes=Save(doc);
        using(var package=SpreadsheetDocument.Open(new MemoryStream(bytes),false))
        {
            Assert.Empty(new OpenXmlValidator().Validate(package));
            var p=package.WorkbookPart!.WorksheetParts.First().Worksheet!.GetFirstChild<SheetProtection>()!;
            Assert.True(p.Sheet!.Value);Assert.False(p.FormatCells!.Value);Assert.True(p.FormatRows!.Value);Assert.False(p.Sort!.Value);
        }
        using var reopened=new XlsxAdapter().Open(new MemoryStream(bytes));
        var cell=PathResolver.Single(reopened.Root,"/sheet[1]/cell[A1]").GetProps();
        Assert.Equal("false",cell["locked"]);Assert.Equal("true",cell["formulaHidden"]);Assert.Equal("true",cell["bold"]);
        Assert.Equal("false",PathResolver.Single(reopened.Root,"/sheet[1]/cell[B3]").GetProps()["locked"]);
        var permissions=JsonNode.Parse(PathResolver.Single(reopened.Root,"/sheet[1]").GetProps()["protection"])!;
        Assert.True(permissions["sort"]!.GetValue<bool>());Assert.False(permissions["objects"]!.GetValue<bool>());
        Set(reopened,"/sheet[1]/cell[A1]",("locked","true"),("formulaHidden","false"));
        using var restored=new XlsxAdapter().Open(new MemoryStream(Save(reopened)));
        cell=PathResolver.Single(restored.Root,"/sheet[1]/cell[A1]").GetProps();
        Assert.False(cell.ContainsKey("locked"));Assert.False(cell.ContainsKey("formulaHidden"));Assert.Equal("true",cell["bold"]);
    }
    [Fact]
    public void Permission_changes_preserve_existing_password_and_validate_before_writing()
    {
        using var doc=new XlsxAdapter().Create();
        Set(doc,"/sheet[1]",("protected","true"));
        var p=((XlsxDocument)doc).Workbook.WorksheetParts.First().Worksheet!.GetFirstChild<SheetProtection>()!;
        p.Password="ABCD";
        Set(doc,"/sheet[1]",("protection","{\"autoFilter\":true}"));
        var before=p.OuterXml;
        Assert.ThrowsAny<Exception>(()=>Set(doc,"/sheet[1]",("protection","{\"sort\":true,\"formatCells\":\"no\"}")));
        Assert.Equal(before,p.OuterXml);
        using var reopened=new XlsxAdapter().Open(new MemoryStream(Save(doc)));
        var props=PathResolver.Single(reopened.Root,"/sheet[1]").GetProps();
        Assert.True(JsonNode.Parse(props["protection"])!["passwordProtected"]!.GetValue<bool>());
        Assert.Equal("ABCD",((XlsxDocument)reopened).Workbook.WorksheetParts.First().Worksheet!.GetFirstChild<SheetProtection>()!.Password!.Value);
    }
}
