using System.Buffers.Binary;
using System.Text;
using Writer.Core;

namespace Writer.Formats.Compat;

/// <summary>A Compound File Binary (OLE2) container, the envelope of .doc, .xls, .ppt and WPS's .wps, .et and .dps:
/// its streams by name, wherever they sit in the storage tree.</summary>
public sealed class Cfb
{
    const int EndOfChain = -2;
    const int FreeSector = -1;

    public static bool IsCfb(ReadOnlySpan<byte> b) =>
        b.Length >= 8 && b[0] == 0xD0 && b[1] == 0xCF && b[2] == 0x11 && b[3] == 0xE0 && b[4] == 0xA1 && b[5] == 0xB1 && b[6] == 0x1A && b[7] == 0xE1;

    readonly byte[] _data;
    readonly int _sectorSize, _miniSectorSize, _miniCutoff;
    readonly int[] _fat, _miniFat;
    readonly byte[] _miniStream;
    readonly List<(string Name, byte Type, int Start, long Size)> _entries = [];

    public Cfb(byte[] data)
    {
        _data = data;
        if (!IsCfb(data) || data.Length < 512) throw new WriterException(ErrorCode.FormatError, "Not an OLE compound file", "The file is damaged or not an Office 97-2003 document.");
        _sectorSize = 1 << BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(0x1E));
        _miniSectorSize = 1 << BinaryPrimitives.ReadUInt16LittleEndian(data.AsSpan(0x20));
        if (_sectorSize is not (512 or 4096) || _miniSectorSize <= 0 || _miniSectorSize > _sectorSize)
            throw new WriterException(ErrorCode.FormatError, "Damaged OLE header", "The file is damaged.");
        var fatSectors = Int(data, 0x2C);
        var dirStart = Int(data, 0x30);
        _miniCutoff = Int(data, 0x38);
        var miniFatStart = Int(data, 0x3C);
        var miniFatCount = Int(data, 0x40);
        var difatStart = Int(data, 0x44);
        var difatCount = Int(data, 0x48);

        // where the FAT sectors are: 109 entries in the header, the rest in a chain of DIFAT sectors
        var difat = new List<int>();
        for (var i = 0; i < 109; i++) difat.Add(Int(data, 0x4C + i * 4));
        var perSector = _sectorSize / 4 - 1;
        for (int s = difatStart, n = 0; s >= 0 && n < difatCount; n++)
        {
            var off = Offset(s);
            if (off + _sectorSize > data.Length) break;
            for (var i = 0; i < perSector; i++) difat.Add(Int(data, off + i * 4));
            s = Int(data, off + perSector * 4);
        }
        var fat = new List<int>();
        foreach (var s in difat.Take(fatSectors))
        {
            if (s < 0) break;
            var off = Offset(s);
            if (off + _sectorSize > data.Length) break;
            for (var i = 0; i < _sectorSize / 4; i++) fat.Add(Int(data, off + i * 4));
        }
        _fat = fat.ToArray();

        var dir = Chain(dirStart, long.MaxValue);
        for (var off = 0; off + 128 <= dir.Length; off += 128)
        {
            var nameLen = BinaryPrimitives.ReadUInt16LittleEndian(dir.AsSpan(off + 0x40));
            var type = dir[off + 0x42];
            if (type == 0) continue;
            var name = nameLen >= 2 && nameLen <= 64 ? Encoding.Unicode.GetString(dir, off, nameLen - 2) : "";
            var size = _sectorSize == 512 ? BinaryPrimitives.ReadUInt32LittleEndian(dir.AsSpan(off + 0x78)) : BinaryPrimitives.ReadInt64LittleEndian(dir.AsSpan(off + 0x78));
            _entries.Add((name, type, Int(dir, off + 0x74), size));
        }
        if (_entries.Count == 0) throw new WriterException(ErrorCode.FormatError, "Empty OLE directory", "The file is damaged.");
        var root = _entries.First(e => e.Type == 5);
        _miniStream = Chain(root.Start, root.Size);
        var miniFat = Chain(miniFatStart, (long)miniFatCount * _sectorSize);
        _miniFat = new int[miniFat.Length / 4];
        for (var i = 0; i < _miniFat.Length; i++) _miniFat[i] = Int(miniFat, i * 4);
    }

    public IEnumerable<string> StreamNames => _entries.Where(e => e.Type == 2).Select(e => e.Name);

    public bool Has(string name) => Find(name) is not null;

    /// <summary>The stream's bytes, or null when there is no stream of that name (case-insensitive).</summary>
    public byte[]? Stream(string name)
    {
        var e = Find(name);
        if (e is null) return null;
        var (_, _, start, size) = e.Value;
        if (size == 0) return [];
        return size < _miniCutoff ? MiniChain(start, size) : Chain(start, size);
    }

    (string Name, byte Type, int Start, long Size)? Find(string name)
    {
        foreach (var e in _entries)
            if (e.Type == 2 && string.Equals(e.Name, name, StringComparison.OrdinalIgnoreCase)) return e;
        return null;
    }

    int Offset(int sector) => (sector + 1) * _sectorSize;

    static int Int(byte[] b, int off) => BinaryPrimitives.ReadInt32LittleEndian(b.AsSpan(off));

    byte[] Chain(int start, long size)
    {
        var ms = new MemoryStream();
        var seen = 0;
        for (var s = start; s >= 0 && s != EndOfChain && s != FreeSector && ms.Length < size; s = s < _fat.Length ? _fat[s] : EndOfChain)
        {
            var off = Offset(s);
            if (off >= _data.Length || ++seen > _data.Length / _sectorSize + 1) break;
            ms.Write(_data, off, (int)Math.Min(_sectorSize, Math.Min(_data.Length - off, size - ms.Length)));
        }
        return ms.ToArray();
    }

    byte[] MiniChain(int start, long size)
    {
        var ms = new MemoryStream();
        var seen = 0;
        for (var s = start; s >= 0 && s != EndOfChain && ms.Length < size; s = s < _miniFat.Length ? _miniFat[s] : EndOfChain)
        {
            var off = s * _miniSectorSize;
            if (off >= _miniStream.Length || ++seen > _miniStream.Length / _miniSectorSize + 1) break;
            ms.Write(_miniStream, off, (int)Math.Min(_miniSectorSize, Math.Min(_miniStream.Length - off, size - ms.Length)));
        }
        return ms.ToArray();
    }
}
