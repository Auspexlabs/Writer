using Writer.Core;

namespace Writer.Formats.Compat;

/// <summary>PowerPoint 97-2003 binary (.ppt, .pot, .pps, and Kingsoft WPS .dps/.dpt, same layout) → a deck.
/// Slides in show order with their shapes flattened out of groups; masters and their decorations are left out.</summary>
public static partial class PptReader
{
    const int RtDocument = 0x03E8, RtDocumentAtom = 0x03E9, RtSlide = 0x03EE, RtSlideAtom = 0x03EF, RtNotes = 0x03F0,
        RtEnvironment = 0x03F2, RtSlidePersistAtom = 0x03F3, RtMainMaster = 0x03F8, RtColorScheme = 0x07F0, RtFontCollection = 0x07D5,
        RtDrawingGroup = 0x040B, RtDrawing = 0x040C, RtPlaceholder = 0x0BC3, RtSlideListWithText = 0x0FF0, RtUserEditAtom = 0x0FF5,
        RtCurrentUserAtom = 0x0FF6, RtOutlineTextRef = 0x0F9E, RtTextHeader = 0x0F9F, RtTextChars = 0x0FA0, RtStyleTextProp = 0x0FA1,
        RtMasterTextProp = 0x0FA2, RtTextBytes = 0x0FA8, RtFontEntity = 0x0FB7, RtPersistDirectory = 0x1772, RtCryptSession = 0x2F14,
        DggContainer = 0xF000, BStore = 0xF001, DgContainer = 0xF002, SpgrContainer = 0xF003, SpContainer = 0xF004, Bse = 0xF007,
        Spgr = 0xF009, Sp = 0xF00A, Opt = 0xF00B, ClientTextbox = 0xF00D, ChildAnchor = 0xF00F, ClientAnchor = 0xF010, ClientData = 0xF011;

    static WriterException Damaged(string what) => new(ErrorCode.FormatError, what, "The file is damaged or not a PowerPoint 97-2003 presentation.");

    public static DeckModel Read(byte[] file, List<string> warnings)
    {
        var cfb = new Cfb(file);
        var doc = cfb.Stream("PowerPoint Document") ?? throw Damaged("No PowerPoint Document stream");
        var (persist, docId, encrypted) = Persist(doc, cfb.Stream("Current User"));
        if (encrypted || persist.Values.Any(o => TypeAt(doc, o) == RtCryptSession))
            throw new WriterException(ErrorCode.FormatError, "This presentation is password-protected", "Remove the password in PowerPoint, then open it again.");
        if (!persist.TryGetValue(docId, out var docOff) || TypeAt(doc, docOff) != RtDocument)
            docOff = Records(doc, 0, doc.Length).FirstOrDefault(r => r.Type == RtDocument && r.Ver == 0xF) is { Len: > 0 } d ? d.Start - 8 : throw Damaged("No document record");
        var docRec = RecAt(doc, docOff);

        var deck = new DeckModel();
        if (Child(doc, docRec, RtDocumentAtom) is { } atom && atom.Len >= 8)
        {
            int w = I32(doc, atom.Start), h = I32(doc, atom.Start + 4);
            if (w > 0 && h > 0) { deck.WidthCm = Cm(w); deck.HeightCm = Cm(h); }
        }

        var ctx = new Ctx { Fonts = Fonts(doc, docRec), Pictures = cfb.Stream("Pictures") ?? [], Bses = Bses(doc, docRec), Doc = doc, Persist = persist, Warnings = warnings };
        var slides = new List<ListEntry>(); var masters = new List<ListEntry>(); var notes = new List<ListEntry>();
        foreach (var l in Children(doc, docRec).Where(r => r.Type == RtSlideListWithText))
            (l.Inst == 0 ? slides : l.Inst == 1 ? masters : notes).AddRange(ListEntries(doc, l));

        foreach (var entry in slides)
        {
            if (!persist.TryGetValue(entry.PersistId, out var off) || TypeAt(doc, off) != RtSlide) { warnings.Add($"slide {entry.SlideId} is missing from the file"); continue; }
            var slideRec = RecAt(doc, off);
            var slide = new SlideModel();
            var sctx = new SlideCtx(ctx, slide);
            if (Child(doc, slideRec, RtSlideAtom) is { } sa && sa.Len >= 24)
            {
                int masterId = I32(doc, sa.Start + 12), notesId = I32(doc, sa.Start + 16), flags = U16(doc, sa.Start + 20);
                sctx.Scheme = (flags & 0x02) == 0 ? Scheme(doc, slideRec) : null;
                sctx.Scheme ??= masters.FirstOrDefault(m => m.SlideId == masterId) is { } m && persist.TryGetValue(m.PersistId, out var mo) && TypeAt(doc, mo) == RtMainMaster ? Scheme(doc, RecAt(doc, mo)) : null;
                if (notesId != 0 && notes.FirstOrDefault(n => n.SlideId == notesId) is { } ne && persist.TryGetValue(ne.PersistId, out var no) && TypeAt(doc, no) == RtNotes)
                    slide.Notes = NotesText(doc, RecAt(doc, no), ne, sctx);
            }
            foreach (var s in Shapes(doc, slideRec, entry, sctx))
            {
                if (s.Background) slide.Background ??= s.Shape.Fill;
                else slide.Shapes.Add(s.Shape);
            }
            deck.Slides.Add(slide);
        }
        return deck;
    }

    /// <summary>persistId → stream offset from the newest edit backwards (a newer definition wins), the document's persist id,
    /// and whether the file says it is encrypted.</summary>
    static (Dictionary<int, int> Map, int DocId, bool Encrypted) Persist(byte[] doc, byte[]? cu)
    {
        var map = new Dictionary<int, int>();
        var encrypted = false;
        int offEdit = 0, docId = 0;
        if (cu is not null)
            foreach (var r in Records(cu, 0, cu.Length))
                if (r.Type == RtCurrentUserAtom && r.Len >= 12) { encrypted = U32(cu, r.Start + 4) == 0xF3D1C4DF; offEdit = I32(cu, r.Start + 8); break; }
        if (offEdit <= 0 || TypeAt(doc, offEdit) != RtUserEditAtom)
            offEdit = Records(doc, 0, doc.Length).Where(r => r.Type == RtUserEditAtom).Select(r => r.Start - 8).LastOrDefault();
        var seen = new HashSet<int>();
        while (offEdit > 0 && offEdit + 36 <= doc.Length && TypeAt(doc, offEdit) == RtUserEditAtom && seen.Add(offEdit))
        {
            int len = I32(doc, offEdit + 4), p = offEdit + 8;
            int offLast = I32(doc, p + 8), offDir = I32(doc, p + 12);
            if (docId == 0) docId = I32(doc, p + 16);
            if (len >= 32 && I32(doc, p + 28) != 0) encrypted = true;
            if (offDir > 0 && TypeAt(doc, offDir) == RtPersistDirectory)
            {
                var end = Math.Min(doc.Length, offDir + 8 + I32(doc, offDir + 4));
                for (var q = offDir + 8; q + 4 <= end;)
                {
                    var h = U32(doc, q); int id = (int)(h & 0xFFFFF), n = (int)(h >> 20); q += 4;
                    for (var i = 0; i < n && q + 4 <= end; i++, q += 4) map.TryAdd(id + i, I32(doc, q));
                }
            }
            offEdit = offLast;
        }
        return (map, docId, encrypted);
    }

    static List<string> Fonts(byte[] doc, Rec docRec)
    {
        var fonts = new List<string>();
        if (Child(doc, docRec, RtEnvironment) is { } env && Child(doc, env, RtFontCollection) is { } fc)
            foreach (var f in Children(doc, fc).Where(r => r.Type == RtFontEntity && r.Len >= 64))
            {
                var name = System.Text.Encoding.Unicode.GetString(doc, f.Start, 64);
                var nul = name.IndexOf('\0');
                fonts.Add(nul >= 0 ? name[..nul] : name);
            }
        return fonts;
    }

    /// <summary>The colour scheme a slide or master carries (the instance-1 ColorSchemeAtom): 8 colours as RRGGBB.</summary>
    static string[]? Scheme(byte[] doc, Rec rec)
    {
        foreach (var c in Children(doc, rec))
            if (c.Type == RtColorScheme && c.Inst == 1 && c.Len >= 32)
            {
                var s = new string[8];
                for (var i = 0; i < 8; i++) s[i] = Inline.Bgr(U32(doc, c.Start + i * 4));
                return s;
            }
        return null;
    }

    sealed class ListEntry { public int PersistId, SlideId; public List<TextGroup> Texts = []; }
    sealed class TextGroup { public int Type = 4; public string Text = ""; public Rec? Style, Master; }

    /// <summary>The SlidePersistAtoms of a SlideListWithText in order, each with the outline texts that follow it.</summary>
    static List<ListEntry> ListEntries(byte[] doc, Rec list)
    {
        var entries = new List<ListEntry>();
        ListEntry? cur = null; TextGroup? g = null;
        foreach (var r in Children(doc, list))
        {
            switch (r.Type)
            {
                case RtSlidePersistAtom when r.Len >= 20: entries.Add(cur = new ListEntry { PersistId = I32(doc, r.Start), SlideId = I32(doc, r.Start + 12) }); g = null; break;
                case RtTextHeader when cur is not null: cur.Texts.Add(g = new TextGroup { Type = I32(doc, r.Start) }); break;
                case RtTextChars or RtTextBytes when g is not null: g.Text = Text(doc, r); break;
                case RtStyleTextProp when g is not null: g.Style = r; break;
                case RtMasterTextProp when g is not null: g.Master = r; break;
            }
        }
        return entries;
    }

    static string? NotesText(byte[] doc, Rec notesRec, ListEntry entry, SlideCtx sctx)
    {
        var lines = Shapes(doc, notesRec, entry, sctx).Where(s => s.Placeholder == 14 || s.TextType == 2).Select(s => s.Raw.Trim()).Where(t => t.Length > 0).ToList();
        return lines.Count > 0 ? string.Join("\n", lines) : null;
    }

    sealed class Ctx
    {
        public List<string> Fonts = [];
        public byte[] Pictures = [], Doc = [];
        public List<(int FoDelay, Rec? Inline)> Bses = [];
        public Dictionary<int, int> Persist = [];
        public List<string> Warnings = [];
        public string? Font(int i) => i >= 0 && i < Fonts.Count ? Fonts[i] : null;
    }

    sealed class SlideCtx(Ctx ctx, SlideModel slide)
    {
        public Ctx Ctx = ctx;
        public SlideModel Slide = slide;
        public string[]? Scheme;
        public string? SchemeColor(int i) => Scheme is not null && i >= 0 && i < 8 ? Scheme[i] : null;
        /// <summary>A text colour: literal when index is 0xFE, else the slide's scheme colour.</summary>
        public string? Color(byte r, byte g, byte b, byte index) => index == 0xFE ? Inline.Rgb(r, g, b) : SchemeColor(index);
        /// <summary>An OfficeArt colour property: BGR, or a scheme index when the 0x08000000 flag is set.</summary>
        public string? OptColor(uint v) => (v & 0x10000000) != 0 ? null : (v & 0x08000000) != 0 ? SchemeColor((int)(v & 0xFF)) : Inline.Bgr(v & 0xFFFFFF);
    }

    static double Cm(double masterUnits) => masterUnits / 576 * 2.54;
}
