// Crystal Reports (Compass) writes Arial as Identity-H with no Unicode map.
// pdftotext then returns glyph junk. Match each embedded glyf outline to the
// system Arial face, then put characters that share a baseline on one line
// so the existing Compass parsers can read the text.
// WinAnsi / PostScript reports take the literal strings and are left as-is.

using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.IO.Compression;
using System.Text;

namespace CrystalPdf
{
    public static class Decoder
    {
        public static void WriteText(string pdfPath, string outPath)
        {
            PdfFile pdf = PdfFile.Open(pdfPath);
            string text = pdf.Extract();
            Encoding enc = Encoding.GetEncoding(1252,
                new EncoderReplacementFallback("?"),
                new DecoderReplacementFallback("?"));
            File.WriteAllText(outPath, text, enc);
        }
    }

    internal struct Hit
    {
        public int Cp;
        public int Adv;
    }

    internal sealed class GMap
    {
        public char[] Ch;
        public int[] Adv;
        public int Units;
    }

    internal struct Run
    {
        public string Text;
        public double Width;
    }

    internal sealed class Face
    {
        public Dictionary<string, List<Hit>> Index;

        public static Face Load(byte[] data)
        {
            Ttf ttf = Ttf.Open(data);
            if (ttf == null) return null;
            Dictionary<int, List<int>> uni = ttf.UnicodeMap();
            Face face = new Face();
            face.Index = new Dictionary<string, List<Hit>>();
            int n = ttf.NumGlyphs;
            for (int gid = 0; gid < n; gid++)
            {
                string sig = ttf.Signature(gid);
                if (sig == null) continue;
                List<int> cps;
                if (!uni.TryGetValue(gid, out cps) || cps.Count == 0) continue;
                List<Hit> list;
                if (!face.Index.TryGetValue(sig, out list))
                {
                    list = new List<Hit>();
                    face.Index[sig] = list;
                }
                int adv = ttf.Advance(gid);
                for (int i = 0; i < cps.Count; i++)
                {
                    Hit h;
                    h.Cp = cps[i];
                    h.Adv = adv;
                    list.Add(h);
                }
            }
            return face;
        }
    }

    internal sealed class Ttf
    {
        byte[] data;
        Dictionary<string, int> off;
        Dictionary<string, int> len;
        public int NumGlyphs;
        public int UnitsPerEm;
        int glyf;
        int hmtx;
        int nMetrics;
        int[] glyphOfs;

        public static Ttf Open(byte[] data)
        {
            if (data == null || data.Length < 12) return null;
            int num = U16(data, 4);
            if (num <= 0 || num > 200 || 12 + num * 16 > data.Length) return null;
            Ttf t = new Ttf();
            t.data = data;
            t.off = new Dictionary<string, int>();
            t.len = new Dictionary<string, int>();
            int p = 12;
            for (int i = 0; i < num; i++)
            {
                string tag = Encoding.ASCII.GetString(data, p, 4);
                int toff = U32(data, p + 8);
                int tlen = U32(data, p + 12);
                t.off[tag] = toff;
                t.len[tag] = tlen;
                p += 16;
            }
            if (!t.off.ContainsKey("glyf") || !t.off.ContainsKey("loca") ||
                !t.off.ContainsKey("head") || !t.off.ContainsKey("maxp")) return null;
            t.glyf = t.off["glyf"];
            t.NumGlyphs = U16(data, t.off["maxp"] + 4);
            t.UnitsPerEm = U16(data, t.off["head"] + 18);
            if (t.UnitsPerEm < 16) t.UnitsPerEm = 1000;
            int locaFmt = U16(data, t.off["head"] + 50);
            int loca = t.off["loca"];
            int n = t.NumGlyphs;
            if (n <= 0 || n > 65535) return null;
            t.glyphOfs = new int[n + 1];
            if (locaFmt == 0)
            {
                if (loca + (n + 1) * 2 > data.Length) return null;
                for (int i = 0; i <= n; i++) t.glyphOfs[i] = U16(data, loca + i * 2) * 2;
            }
            else
            {
                if (loca + (n + 1) * 4 > data.Length) return null;
                for (int i = 0; i <= n; i++) t.glyphOfs[i] = U32(data, loca + i * 4);
            }
            t.hmtx = t.off.ContainsKey("hmtx") ? t.off["hmtx"] : -1;
            t.nMetrics = 0;
            if (t.off.ContainsKey("hhea")) t.nMetrics = U16(data, t.off["hhea"] + 34);
            return t;
        }

        public int Advance(int gid)
        {
            if (hmtx < 0 || nMetrics <= 0) return 0;
            int g = gid;
            if (g >= nMetrics) g = nMetrics - 1;
            if (g < 0) return 0;
            int o = hmtx + g * 4;
            if (o + 2 > data.Length) return 0;
            return U16(data, o);
        }

        public Dictionary<int, List<int>> UnicodeMap()
        {
            Dictionary<int, List<int>> uni = new Dictionary<int, List<int>>();
            if (!off.ContainsKey("cmap")) return uni;
            int cmap = off["cmap"];
            if (cmap + 4 > data.Length) return uni;
            int nsub = U16(data, cmap + 2);
            for (int i = 0; i < nsub; i++)
            {
                int rec = cmap + 4 + i * 8;
                if (rec + 8 > data.Length) break;
                int plat = U16(data, rec);
                int enc = U16(data, rec + 2);
                int so = cmap + U32(data, rec + 4);
                if (so + 2 > data.Length) continue;
                int fmt = U16(data, so);
                if (plat == 3 && enc == 1 && fmt == 4)
                {
                    ReadFmt4(so, uni);
                    break;
                }
            }
            return uni;
        }

        void ReadFmt4(int so, Dictionary<int, List<int>> uni)
        {
            if (so + 14 > data.Length) return;
            int segc = U16(data, so + 6) / 2;
            if (segc <= 0 || segc > 20000) return;
            int endOff = so + 14;
            int startOff = so + 16 + segc * 2;
            int deltaOff = startOff + segc * 2;
            int roOff = deltaOff + segc * 2;
            if (roOff + segc * 2 > data.Length) return;
            for (int s = 0; s < segc; s++)
            {
                int end = U16(data, endOff + s * 2);
                int start = U16(data, startOff + s * 2);
                int delta = I16(data, deltaOff + s * 2);
                int ro = U16(data, roOff + s * 2);
                for (int cp = start; cp <= end; cp++)
                {
                    if (cp == 0xFFFF) continue;
                    int gid;
                    if (ro == 0)
                    {
                        gid = (cp + delta) & 0xFFFF;
                    }
                    else
                    {
                        int pos = roOff + s * 2 + ro + (cp - start) * 2;
                        if (pos + 2 > data.Length) continue;
                        gid = U16(data, pos);
                        if (gid != 0) gid = (gid + delta) & 0xFFFF;
                    }
                    if (gid == 0) continue;
                    List<int> list;
                    if (!uni.TryGetValue(gid, out list))
                    {
                        list = new List<int>();
                        uni[gid] = list;
                    }
                    list.Add(cp);
                }
            }
        }

        public string Signature(int gid)
        {
            if (gid < 0 || gid >= NumGlyphs) return null;
            int start = glyphOfs[gid];
            int end = glyphOfs[gid + 1];
            if (end <= start || end - start < 10) return null;
            int g0 = glyf + start;
            if (g0 < 0 || glyf + end > data.Length) return null;
            int ncont = I16(data, g0);
            if (ncont <= 0) return null;
            if (ncont > 64) return null;
            int xmin = I16(data, g0 + 2);
            int ymin = I16(data, g0 + 4);
            int xmax = I16(data, g0 + 6);
            int ymax = I16(data, g0 + 8);
            double aspect = Math.Round((xmax - xmin) / (double)Math.Max(ymax - ymin, 1), 2);
            int p = g0 + 10;
            if (p + 2 * ncont + 2 > data.Length) return null;
            int lastEnd = 0;
            for (int c = 0; c < ncont; c++) lastEnd = U16(data, p + c * 2);
            int npt = lastEnd + 1;
            if (npt <= 0 || npt > 4000) return null;
            int inst = U16(data, p + 2 * ncont);
            p = p + 2 * ncont + 2 + inst;
            if (p > data.Length) return null;
            byte[] flags = new byte[npt];
            int i = 0;
            while (i < npt)
            {
                if (p >= data.Length) return null;
                byte f = data[p++];
                flags[i] = f;
                if ((f & 8) != 0)
                {
                    if (p >= data.Length) return null;
                    int rep = data[p++];
                    for (int r = 0; r < rep; r++)
                    {
                        i++;
                        if (i >= npt) return null;
                        flags[i] = f;
                    }
                }
                i++;
            }
            int[] xs = new int[npt];
            int[] ys = new int[npt];
            int x = 0;
            for (int k = 0; k < npt; k++)
            {
                int f = flags[k];
                int dx;
                if ((f & 2) != 0)
                {
                    if (p >= data.Length) return null;
                    dx = data[p++];
                    if ((f & 16) == 0) dx = -dx;
                }
                else if ((f & 16) != 0) dx = 0;
                else
                {
                    if (p + 2 > data.Length) return null;
                    dx = I16(data, p);
                    p += 2;
                }
                x += dx;
                xs[k] = x;
            }
            int y = 0;
            for (int k = 0; k < npt; k++)
            {
                int f = flags[k];
                int dy;
                if ((f & 4) != 0)
                {
                    if (p >= data.Length) return null;
                    dy = data[p++];
                    if ((f & 32) == 0) dy = -dy;
                }
                else if ((f & 32) != 0) dy = 0;
                else
                {
                    if (p + 2 > data.Length) return null;
                    dy = I16(data, p);
                    p += 2;
                }
                y += dy;
                ys[k] = y;
            }
            int minx = xs[0], maxx = xs[0], miny = ys[0], maxy = ys[0];
            for (int k = 1; k < npt; k++)
            {
                if (xs[k] < minx) minx = xs[k];
                if (xs[k] > maxx) maxx = xs[k];
                if (ys[k] < miny) miny = ys[k];
                if (ys[k] > maxy) maxy = ys[k];
            }
            int sx = Math.Max(maxx - minx, 1);
            int sy = Math.Max(maxy - miny, 1);
            StringBuilder sb = new StringBuilder(npt * 8);
            sb.Append(ncont);
            sb.Append('|');
            sb.Append(npt);
            sb.Append('|');
            sb.Append(aspect.ToString("0.00", CultureInfo.InvariantCulture));
            for (int k = 0; k < npt; k++)
            {
                int nx = (int)Math.Round((xs[k] - minx) * 24.0 / sx);
                int ny = (int)Math.Round((ys[k] - miny) * 24.0 / sy);
                int on = (flags[k] & 1) != 0 ? 1 : 0;
                sb.Append('|');
                sb.Append(nx);
                sb.Append(',');
                sb.Append(ny);
                sb.Append(',');
                sb.Append(on);
            }
            return sb.ToString();
        }

        static int U16(byte[] d, int o)
        {
            return (d[o] << 8) | d[o + 1];
        }

        static int U32(byte[] d, int o)
        {
            return (d[o] << 24) | (d[o + 1] << 16) | (d[o + 2] << 8) | d[o + 3];
        }

        static int I16(byte[] d, int o)
        {
            int v = U16(d, o);
            if (v >= 32768) v -= 65536;
            return v;
        }
    }

    internal sealed class PdfFile
    {
        static Dictionary<string, Face> faces;
        Dictionary<int, string> objs = new Dictionary<int, string>();
        Dictionary<int, byte[]> streams = new Dictionary<int, byte[]>();
        Encoding latin1 = Encoding.GetEncoding(28591);

        public static PdfFile Open(string path)
        {
            byte[] bytes = File.ReadAllBytes(path);
            PdfFile pdf = new PdfFile();
            pdf.Parse(bytes);
            return pdf;
        }

        public string Extract()
        {
            Dictionary<string, int> fontFile = new Dictionary<string, int>();
            Dictionary<int, string> baseFont = new Dictionary<int, string>();
            foreach (KeyValuePair<int, string> kv in objs)
            {
                string body = kv.Value;
                if (body == null || body.Length == 0) continue;
                int ff = MatchInt(body, "/FontFile2\\s+(\\d+)\\s+0\\s+R");
                string fname = MatchName(body, "/FontName\\s*/([^\\s/<>\\[\\]]+)");
                if (ff > 0 && fname != null) fontFile[fname] = ff;
                string bf = MatchName(body, "/BaseFont\\s*/([^\\s/<>\\[\\]]+)");
                if (bf != null) baseFont[kv.Key] = bf;
            }

            int root = 0;
            foreach (KeyValuePair<int, string> kv in objs)
            {
                if (kv.Value != null && kv.Value.IndexOf("/Type/Catalog") >= 0)
                {
                    int pages = MatchInt(kv.Value, "/Pages\\s+(\\d+)\\s+0\\s+R");
                    if (pages > 0) { root = pages; break; }
                }
                if (kv.Value != null && kv.Value.IndexOf("/Type /Catalog") >= 0)
                {
                    int pages = MatchInt(kv.Value, "/Pages\\s+(\\d+)\\s+0\\s+R");
                    if (pages > 0) { root = pages; break; }
                }
            }

            List<int> pageIds = new List<int>();
            HashSet<int> seen = new HashSet<int>();
            if (root > 0) WalkPages(root, pageIds, seen);

            Dictionary<int, GMap> gidMaps = new Dictionary<int, GMap>();
            StringBuilder all = new StringBuilder();
            int pageCount = 0;
            int letters = 0;
            if (pageIds.Count == 0)
            {
                foreach (KeyValuePair<int, byte[]> kv in streams)
                {
                    string s = latin1.GetString(kv.Value);
                    if (s.IndexOf("Tj") >= 0 && s.IndexOf("Tm") >= 0)
                        pageIds.Add(kv.Key);
                }
            }

            for (int pi = 0; pi < pageIds.Count; pi++)
            {
                int pageId = pageIds[pi];
                string page;
                if (!objs.TryGetValue(pageId, out page) || page == null) continue;
                if (page.IndexOf("/Contents") < 0 && streams.ContainsKey(pageId))
                {
                    // Fallback path stored a content stream id, not a page id.
                    string one = DecodeContent(new int[] { pageId }, null, fontFile, baseFont, gidMaps);
                    if (one.Length > 0)
                    {
                        if (all.Length > 0) all.Append('\n');
                        all.Append(one);
                        pageCount++;
                    }
                    continue;
                }
                int resId = MatchInt(page, "/Resources\\s+(\\d+)\\s+0\\s+R");
                string resBody = page;
                if (resId > 0 && objs.ContainsKey(resId)) resBody = objs[resId];
                Dictionary<string, GMap> fonts = FontsFor(resBody, fontFile, baseFont, gidMaps);
                int[] contents = ContentIds(page);
                string text = DecodeContent(contents, fonts, fontFile, baseFont, gidMaps);
                if (text.Length > 0)
                {
                    if (all.Length > 0) all.Append('\n');
                    all.Append(text);
                    pageCount++;
                }
            }
            string result = all.ToString();
            for (int i = 0; i < result.Length; i++)
                if (char.IsLetter(result[i])) letters++;
            Console.Error.WriteLine("crystal pages=" + pageCount + " letters=" + letters +
                " summary=" + (result.IndexOf("Summary") >= 0) +
                " plan=" + (result.IndexOf("Plan Sections") >= 0) +
                " sf=" + (result.IndexOf("Separation Factor") >= 0));
            return result;
        }

        void WalkPages(int id, List<int> pageIds, HashSet<int> seen)
        {
            if (seen.Contains(id)) return;
            seen.Add(id);
            string body;
            if (!objs.TryGetValue(id, out body) || body == null) return;
            bool pages = body.IndexOf("/Type/Pages") >= 0 || body.IndexOf("/Type /Pages") >= 0;
            bool page = !pages && (body.IndexOf("/Type/Page") >= 0 || body.IndexOf("/Type /Page") >= 0);
            if (pages || body.IndexOf("/Kids") >= 0)
            {
                int k = body.IndexOf("/Kids");
                if (k >= 0)
                {
                    int lb = body.IndexOf('[', k);
                    int rb = lb >= 0 ? body.IndexOf(']', lb) : -1;
                    if (lb >= 0 && rb > lb)
                    {
                        string inside = body.Substring(lb + 1, rb - lb - 1);
                        int p = 0;
                        while (p < inside.Length)
                        {
                            while (p < inside.Length && !char.IsDigit(inside[p])) p++;
                            int s = p;
                            while (p < inside.Length && char.IsDigit(inside[p])) p++;
                            if (s == p) break;
                            int kid = int.Parse(inside.Substring(s, p - s));
                            while (p < inside.Length && char.IsWhiteSpace(inside[p])) p++;
                            if (p + 3 <= inside.Length && inside.Substring(p, 3) == "0 R")
                            {
                                WalkPages(kid, pageIds, seen);
                                p += 3;
                            }
                        }
                    }
                }
            }
            if (page) pageIds.Add(id);
        }

        Dictionary<string, GMap> FontsFor(string resBody, Dictionary<string, int> fontFile,
            Dictionary<int, string> baseFont, Dictionary<int, GMap> gidMaps)
        {
            Dictionary<string, GMap> map = new Dictionary<string, GMap>();
            if (resBody == null) return map;
            int f = IndexOfFontDict(resBody);
            if (f < 0) return map;
            int depth = 0;
            int i = f;
            int end = resBody.Length;
            while (i < resBody.Length - 1)
            {
                if (resBody[i] == '<' && resBody[i + 1] == '<') { depth++; i += 2; continue; }
                if (resBody[i] == '>' && resBody[i + 1] == '>')
                {
                    depth--;
                    i += 2;
                    if (depth == 0) { end = i; break; }
                    continue;
                }
                i++;
            }
            string fontDict = resBody.Substring(f, end - f);
            int p = 0;
            while (p < fontDict.Length)
            {
                int slash = fontDict.IndexOf('/', p);
                if (slash < 0 || slash + 1 >= fontDict.Length) break;
                if (fontDict[slash + 1] == '<' || fontDict[slash + 1] == '>') { p = slash + 1; continue; }
                int ns = slash + 1;
                int ne = ns;
                while (ne < fontDict.Length && IsNameChar(fontDict[ne])) ne++;
                string rname = fontDict.Substring(ns, ne - ns);
                int id = MatchInt(fontDict.Substring(ne), "^\\s*(\\d+)\\s+0\\s+R");
                p = ne;
                if (id <= 0 || rname.Length == 0) continue;
                if (rname == "Font" || rname == "Type" || rname == "Subtype") continue;
                GMap gids = MapForFont(id, fontFile, baseFont, gidMaps);
                if (gids != null) map[rname] = gids;
            }
            return map;
        }

        static int IndexOfFontDict(string body)
        {
            int p = 0;
            while (p < body.Length)
            {
                int i = body.IndexOf("/Font", p);
                if (i < 0) return -1;
                int n = i + 5;
                if (n < body.Length && IsNameChar(body[n])) { p = n; continue; }
                while (n < body.Length && char.IsWhiteSpace(body[n])) n++;
                if (n + 1 < body.Length && body[n] == '<' && body[n + 1] == '<') return n;
                p = n;
            }
            return -1;
        }

        GMap MapForFont(int fontObj, Dictionary<string, int> fontFile,
            Dictionary<int, string> baseFont, Dictionary<int, GMap> gidMaps)
        {
            string bf;
            if (!baseFont.TryGetValue(fontObj, out bf) || bf == null)
            {
                string body;
                if (objs.TryGetValue(fontObj, out body) && body != null)
                {
                    int desc = MatchInt(body, "/DescendantFonts\\s*\\[\\s*(\\d+)\\s+0\\s+R");
                    if (desc > 0) baseFont.TryGetValue(desc, out bf);
                }
            }
            if (bf == null) return null;
            int fileId = -1;
            if (!fontFile.TryGetValue(bf, out fileId))
            {
                string bare = Bare(bf);
                foreach (KeyValuePair<string, int> kv in fontFile)
                {
                    if (Bare(kv.Key) == bare) { fileId = kv.Value; break; }
                }
            }
            if (fileId <= 0) return null;
            GMap cached;
            if (gidMaps.TryGetValue(fileId, out cached)) return cached;
            byte[] ttfBytes;
            if (!streams.TryGetValue(fileId, out ttfBytes)) return null;
            GMap map = BuildGidMap(bf, ttfBytes);
            if (map != null) gidMaps[fileId] = map;
            return map;
        }

        static string Bare(string name)
        {
            int p = name.LastIndexOf('+');
            if (p >= 0 && p + 1 < name.Length) return name.Substring(p + 1);
            return name;
        }

        static GMap BuildGidMap(string fontName, byte[] ttfBytes)
        {
            Face face = GetFace(fontName);
            Ttf sub = Ttf.Open(ttfBytes);
            if (face == null || sub == null) return null;
            GMap map = new GMap();
            map.Ch = new char[sub.NumGlyphs];
            map.Adv = new int[sub.NumGlyphs];
            map.Units = sub.UnitsPerEm > 16 ? sub.UnitsPerEm : 1000;
            int known = 0;
            int miss = 0;
            for (int gid = 1; gid < sub.NumGlyphs; gid++)
            {
                map.Adv[gid] = sub.Advance(gid);
                string sig = sub.Signature(gid);
                if (sig == null)
                {
                    map.Ch[gid] = ' ';
                    continue;
                }
                List<Hit> hits;
                if (!face.Index.TryGetValue(sig, out hits) || hits.Count == 0)
                {
                    map.Ch[gid] = '?';
                    miss++;
                    continue;
                }
                char ch = Prefer(hits, map.Adv[gid]);
                map.Ch[gid] = ch;
                known++;
            }
            Console.Error.WriteLine("font " + Bare(fontName) + " glyphs=" + (sub.NumGlyphs - 1) +
                " matched=" + known + " miss=" + miss);
            return map;
        }

        static char Prefer(List<Hit> hits, int adv)
        {
            List<int> cps = new List<int>();
            for (int i = 0; i < hits.Count; i++)
            {
                if (hits[i].Adv == adv && hits[i].Cp > 0 && hits[i].Cp != 0xFFFF)
                {
                    if (!cps.Contains(hits[i].Cp)) cps.Add(hits[i].Cp);
                }
            }
            if (cps.Count == 0)
            {
                for (int i = 0; i < hits.Count; i++)
                {
                    if (hits[i].Cp > 0 && hits[i].Cp != 0xFFFF && !cps.Contains(hits[i].Cp))
                        cps.Add(hits[i].Cp);
                }
            }
            if (cps.Count == 0) return ' ';
            char ch = (char)cps[0];
            if (cps.Count > 1)
            {
                for (int i = 0; i < cps.Count; i++)
                {
                    int c = cps[i];
                    if ((c >= '0' && c <= '9') || (c >= 'A' && c <= 'Z') || (c >= 'a' && c <= 'z'))
                    {
                        ch = (char)c;
                        break;
                    }
                }
            }
            return Norm(ch);
        }

        static char Norm(char c)
        {
            if (c == '\u00A0' || c == '\u2007' || c == '\u202F') return ' ';
            UnicodeCategory cat = char.GetUnicodeCategory(c);
            if (cat == UnicodeCategory.DashPunctuation) return '-';
            if (c == '\u2212') return '-';
            return c;
        }

        static Face GetFace(string fontName)
        {
            if (faces == null) faces = new Dictionary<string, Face>();
            string kind = KindOf(fontName);
            Face face;
            if (faces.TryGetValue(kind, out face)) return face;
            string path = FontPath(kind);
            if (!File.Exists(path))
            {
                faces[kind] = null;
                return null;
            }
            face = Face.Load(File.ReadAllBytes(path));
            faces[kind] = face;
            Console.Error.WriteLine("indexed " + kind + " " + path);
            return face;
        }

        static string KindOf(string fontName)
        {
            string n = Bare(fontName).ToUpperInvariant();
            bool bold = n.IndexOf("BOLD") >= 0;
            bool ital = n.IndexOf("ITAL") >= 0 || n.IndexOf("OBLIQUE") >= 0;
            if (n.IndexOf("SANSSERIF") >= 0) return "ms";
            if (bold && ital) return "bi";
            if (bold) return "b";
            if (ital) return "i";
            return "r";
        }

        static string FontPath(string kind)
        {
            string fonts = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Windows), "Fonts");
            string name = "arial.ttf";
            if (kind == "b") name = "arialbd.ttf";
            else if (kind == "i") name = "ariali.ttf";
            else if (kind == "bi") name = "arialbi.ttf";
            else if (kind == "ms") name = "micross.ttf";
            string path = Path.Combine(fonts, name);
            if (!File.Exists(path)) path = Path.Combine(fonts, "arial.ttf");
            return path;
        }

        string DecodeContent(int[] contents, Dictionary<string, GMap> fonts,
            Dictionary<string, int> fontFile, Dictionary<int, string> baseFont,
            Dictionary<int, GMap> gidMaps)
        {
            if (contents == null || contents.Length == 0) return "";
            StringBuilder raw = new StringBuilder();
            for (int i = 0; i < contents.Length; i++)
            {
                byte[] b;
                if (!streams.TryGetValue(contents[i], out b) || b == null) continue;
                raw.Append(latin1.GetString(b));
                raw.Append('\n');
            }
            if (raw.Length == 0) return "";
            return Layout(raw.ToString(), fonts);
        }

        static string Layout(string cs, Dictionary<string, GMap> fonts)
        {
            List<Frag> frags = new List<Frag>();
            Lexer lx = new Lexer(cs);
            List<double> nums = new List<double>();
            string lastName = "";
            Token prev = null;
            double[] ctm = new double[] { 1, 0, 0, 1, 0, 0 };
            Stack<double[]> ctmStack = new Stack<double[]>();
            // Text matrix (linear part). Td / TD moves are in text space, so they
            // are scaled by it: with "7.15 0 0 7.15 x y Tm" a Td of 46.6 is 333 pt.
            double tmA = 1, tmB = 0, tmC = 0, tmD = 1;
            // How far the pen has moved along the line since the last Tm / Td.
            double pen = 0;
            while (true)
            {
                Token t = lx.Next();
                if (t.Kind == Tok.Eof) break;
                if (t.Kind == Tok.Num)
                {
                    nums.Add(t.Num);
                    if (nums.Count > 8) nums.RemoveAt(0);
                    prev = t;
                    continue;
                }
                if (t.Kind == Tok.Name)
                {
                    lastName = t.Text;
                    prev = t;
                    continue;
                }
                if (t.Kind == Tok.Op)
                {
                    // Crystal moves each piece of a line with its own q / cm / Q, so a
                    // Tm x is only meaningful after the current transform is applied.
                    if (t.Text == "q")
                    {
                        ctmStack.Push((double[])ctm.Clone());
                    }
                    else if (t.Text == "Q")
                    {
                        if (ctmStack.Count > 0) ctm = ctmStack.Pop();
                    }
                    else if (t.Text == "cm" && nums.Count >= 6)
                    {
                        int n = nums.Count;
                        double ma = nums[n - 6], mb = nums[n - 5], mc = nums[n - 4];
                        double md = nums[n - 3], me = nums[n - 2], mf = nums[n - 1];
                        ctm = new double[]
                        {
                            ma * ctm[0] + mb * ctm[2],
                            ma * ctm[1] + mb * ctm[3],
                            mc * ctm[0] + md * ctm[2],
                            mc * ctm[1] + md * ctm[3],
                            me * ctm[0] + mf * ctm[2] + ctm[4],
                            me * ctm[1] + mf * ctm[3] + ctm[5]
                        };
                    }
                    else if (t.Text == "BT")
                    {
                        lx.X = 0;
                        lx.Y = 0;
                        tmA = 1; tmB = 0; tmC = 0; tmD = 1;
                        pen = 0;
                        lx.Size = lx.TfSize > 0.01 ? lx.TfSize : 1;
                    }
                    else if (t.Text == "Tf" && nums.Count >= 1)
                    {
                        lx.TfSize = Math.Abs(nums[nums.Count - 1]);
                        if (lx.TfSize < 0.01) lx.TfSize = 1;
                        lx.Size = TextSize(tmA, lx.TfSize);
                    }
                    else if (t.Text == "Tm" && nums.Count >= 6)
                    {
                        int n = nums.Count;
                        tmA = nums[n - 6]; tmB = nums[n - 5]; tmC = nums[n - 4]; tmD = nums[n - 3];
                        lx.X = nums[n - 2];
                        lx.Y = nums[n - 1];
                        pen = 0;
                        lx.Size = TextSize(tmA, lx.TfSize);
                    }
                    else if ((t.Text == "Td" || t.Text == "TD") && nums.Count >= 2)
                    {
                        double tx = nums[nums.Count - 2];
                        double ty = nums[nums.Count - 1];
                        lx.X += tx * tmA + ty * tmC;
                        lx.Y += tx * tmB + ty * tmD;
                        pen = 0;
                    }
                    else if (t.Text == "Tj")
                    {
                        Run run = new Run();
                        if (prev != null && prev.Kind == Tok.Hex) run = DecodeHex(prev.Text, fonts, lx.Font, lx.Size);
                        else if (prev != null && prev.Kind == Tok.Lit) run = RunOf(prev.Text, lx.Size);
                        if (run.Text != null && run.Text.Length > 0)
                            frags.Add(PlaceFrag(ctm, lx, run, frags.Count, pen));
                        pen += run.Width;
                    }
                    else if (t.Text == "TJ" && prev != null && prev.Kind == Tok.Array)
                    {
                        Run run = DecodeArray(prev, fonts, lx.Font, lx.Size);
                        if (run.Text != null && run.Text.Length > 0)
                            frags.Add(PlaceFrag(ctm, lx, run, frags.Count, pen));
                        pen += run.Width;
                    }
                    nums.Clear();
                    prev = t;
                    continue;
                }
                prev = t;
            }
            return JoinFrags(frags);
        }

        static double TextSize(double tmA, double tfSize)
        {
            double scale = Math.Abs(tmA);
            if (scale < 0.01) scale = 1;
            return scale * (tfSize > 0.01 ? tfSize : 1);
        }

        static Frag PlaceFrag(double[] ctm, Lexer lx, Run run, int order, double pen)
        {
            double sc = Math.Sqrt(ctm[0] * ctm[0] + ctm[1] * ctm[1]);
            if (sc < 0.001) sc = 1;
            double size = lx.Size < 0.5 ? 7 : lx.Size;
            double ux = lx.X + pen;
            return new Frag(
                ux * ctm[0] + lx.Y * ctm[2] + ctm[4],
                ux * ctm[1] + lx.Y * ctm[3] + ctm[5],
                run.Text, run.Width * sc, size * sc, order);
        }

        static Run RunOf(string text, double size)
        {
            Run run;
            run.Text = text == null ? "" : text;
            run.Width = TextWidth(run.Text, size);
            return run;
        }

        // TJ numbers are thousandths of a text unit, subtracted from x.
        // A large negative number is the gap between columns, not kerning.
        static Run DecodeArray(Token arr, Dictionary<string, GMap> fonts, string font, double size)
        {
            Run run;
            run.Text = "";
            run.Width = 0;
            if (arr.Items == null) return run;
            if (size < 0.5) size = 7;
            StringBuilder sb = new StringBuilder();
            double w = 0;
            for (int i = 0; i < arr.Items.Count; i++)
            {
                Token t = arr.Items[i];
                if (t.Kind == Tok.Hex)
                {
                    Run part = DecodeHex(t.Text, fonts, font, size);
                    sb.Append(part.Text);
                    w += part.Width;
                }
                else if (t.Kind == Tok.Lit)
                {
                    sb.Append(t.Text);
                    w += TextWidth(t.Text, size);
                }
                else if (t.Kind == Tok.Num)
                {
                    double gap = -t.Num / 1000.0 * size;
                    w += gap;
                    bool splitDec = false;
                    if (i + 1 < arr.Items.Count && arr.Items[i + 1].Kind == Tok.Hex && sb.Length > 0 && char.IsDigit(sb[sb.Length - 1]))
                    {
                        Run peek = DecodeHex(arr.Items[i + 1].Text, fonts, font, size);
                        if (peek.Text.Length > 1 && peek.Text[0] == '.' && char.IsDigit(peek.Text[1])) splitDec = true;
                    }
                    // Column gaps are many points. A decimal Crystal split off the
                    // integer ("1,053" + ".68") must stay one number.
                    if (gap > 6.0 && !splitDec) sb.Append(' ');
                }
            }
            run.Text = sb.ToString();
            run.Width = w;
            return run;
        }

        static Run DecodeHex(string hex, Dictionary<string, GMap> fonts, string font, double size)
        {
            Run run;
            run.Text = "";
            run.Width = 0;
            StringBuilder digits = new StringBuilder(hex.Length);
            for (int i = 0; i < hex.Length; i++)
            {
                char c = hex[i];
                if ((c >= '0' && c <= '9') || (c >= 'a' && c <= 'f') || (c >= 'A' && c <= 'F'))
                    digits.Append(c);
            }
            GMap map = null;
            if (fonts != null && font != null) fonts.TryGetValue(font, out map);
            StringBuilder sb = new StringBuilder();
            if (size < 0.5) size = 7;
            if (map != null && digits.Length >= 4 && (digits.Length % 4) == 0)
            {
                double w = 0;
                for (int i = 0; i < digits.Length; i += 4)
                {
                    int gid = int.Parse(digits.ToString(i, 4), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                    if (gid <= 0 || gid >= map.Ch.Length) continue;
                    char ch = map.Ch[gid];
                    if (ch != '\0') sb.Append(ch);
                    int units = map.Units > 16 ? map.Units : 1000;
                    w += map.Adv[gid] / (double)units * size;
                }
                run.Text = sb.ToString();
                run.Width = w;
                return run;
            }
            if ((digits.Length % 2) != 0) return run;
            for (int i = 0; i < digits.Length; i += 2)
            {
                int b = int.Parse(digits.ToString(i, 2), NumberStyles.HexNumber, CultureInfo.InvariantCulture);
                if (b >= 32) sb.Append((char)b);
            }
            run.Text = sb.ToString();
            run.Width = TextWidth(run.Text, size);
            return run;
        }

        static double TextWidth(string t, double size)
        {
            double w = 0;
            for (int i = 0; i < t.Length; i++)
            {
                char c = t[i];
                if (c == ' ') w += 0.28 * size;
                else if (c == '.' || c == ',' || c == ':' || c == '|' || c == 'l' || c == 'i') w += 0.28 * size;
                else w += 0.56 * size;
            }
            return w;
        }

        static string JoinFrags(List<Frag> frags)
        {
            Dictionary<long, List<Frag>> rows = new Dictionary<long, List<Frag>>();
            List<long> keys = new List<long>();
            for (int i = 0; i < frags.Count; i++)
            {
                long key = (long)Math.Round(frags[i].Y);
                List<Frag> row;
                if (!rows.TryGetValue(key, out row))
                {
                    row = new List<Frag>();
                    rows[key] = row;
                    keys.Add(key);
                }
                row.Add(frags[i]);
            }
            keys.Sort(delegate(long a, long b) { return b.CompareTo(a); });
            StringBuilder sb = new StringBuilder();
            for (int k = 0; k < keys.Count; k++)
            {
                List<Frag> row = rows[keys[k]];
                row.Sort(delegate(Frag a, Frag b)
                {
                    if (a.X < b.X) return -1;
                    if (a.X > b.X) return 1;
                    return a.Order.CompareTo(b.Order);
                });
                StringBuilder line = new StringBuilder();
                double right = 0;
                for (int i = 0; i < row.Count; i++)
                {
                    Frag f = row[i];
                    string t = f.Text;
                    if (t == null || t.Length == 0) continue;
                    // Crystal cuts one number into several shows ("48.5" + "9"). Those
                    // meet exactly where the last one ended. A real column gap is wider.
                    if (line.Length > 0 && line[line.Length - 1] != ' ' && t[0] != ' ' &&
                        (f.X - right) > 0.15 * f.Size)
                        line.Append(' ');
                    bool first = line.Length == 0;
                    line.Append(t);
                    right = first ? f.X + f.Width : Math.Max(right, f.X + f.Width);
                }
                string s = line.ToString().Trim();
                if (s.Length == 0) continue;
                if (sb.Length > 0) sb.Append('\n');
                sb.Append(s);
            }
            return sb.ToString();
        }

        static int[] ContentIds(string page)
        {
            List<int> ids = new List<int>();
            int c = page.IndexOf("/Contents");
            if (c < 0) return ids.ToArray();
            int lb = page.IndexOf('[', c);
            int paren = page.IndexOf('R', c);
            if (lb >= 0 && (paren < 0 || lb < paren))
            {
                int rb = page.IndexOf(']', lb);
                if (rb > lb)
                {
                    string inside = page.Substring(lb + 1, rb - lb - 1);
                    int p = 0;
                    while (p < inside.Length)
                    {
                        while (p < inside.Length && !char.IsDigit(inside[p])) p++;
                        int s = p;
                        while (p < inside.Length && char.IsDigit(inside[p])) p++;
                        if (s == p) break;
                        ids.Add(int.Parse(inside.Substring(s, p - s)));
                    }
                }
            }
            else
            {
                int id = MatchInt(page.Substring(c), "/Contents\\s+(\\d+)\\s+0\\s+R");
                if (id > 0) ids.Add(id);
            }
            return ids.ToArray();
        }

        static int MatchInt(string body, string pattern)
        {
            if (body == null) return -1;
            System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(body, pattern);
            if (!m.Success) return -1;
            return int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
        }

        static string MatchName(string body, string pattern)
        {
            if (body == null) return null;
            System.Text.RegularExpressions.Match m = System.Text.RegularExpressions.Regex.Match(body, pattern);
            if (!m.Success) return null;
            return m.Groups[1].Value;
        }

        static bool IsNameChar(char c)
        {
            if (c == '/' || c == '<' || c == '>' || c == '[' || c == ']' || c == '(' || c == ')') return false;
            if (char.IsWhiteSpace(c)) return false;
            return true;
        }

        void Parse(byte[] bytes)
        {
            string raw = latin1.GetString(bytes);
            int i = 0;
            int guard = 0;
            while (i < raw.Length && guard < 100000)
            {
                guard++;
                int num;
                int after = FindObj(raw, i, out num);
                if (after < 0) break;
                int p = after;
                while (p < raw.Length && char.IsWhiteSpace(raw[p])) p++;
                if (p + 1 < raw.Length && raw[p] == '<' && raw[p + 1] == '<')
                {
                    int dictEnd = SkipDict(raw, p);
                    string dict = raw.Substring(p, dictEnd - p);
                    int q = dictEnd;
                    while (q < raw.Length && char.IsWhiteSpace(raw[q])) q++;
                    if (q + 6 <= raw.Length && raw.Substring(q, 6) == "stream")
                    {
                        q += 6;
                        if (q < raw.Length && raw[q] == '\r') q++;
                        if (q < raw.Length && raw[q] == '\n') q++;
                        int len = DirectLength(dict);
                        int streamEnd;
                        if (len >= 0 && q + len <= bytes.Length)
                        {
                            streamEnd = q + len;
                            int look = raw.IndexOf("endstream", streamEnd, Math.Min(16, raw.Length - streamEnd));
                            if (look >= 0) streamEnd = look;
                        }
                        else
                        {
                            int es = raw.IndexOf("endstream", q);
                            streamEnd = es >= 0 ? es : raw.Length;
                        }
                        byte[] slice = new byte[Math.Max(0, streamEnd - q)];
                        if (slice.Length > 0) Buffer.BlockCopy(bytes, q, slice, 0, slice.Length);
                        byte[] data = slice;
                        if (dict.IndexOf("/FlateDecode") >= 0 || dict.IndexOf("/Fl ") >= 0)
                        {
                            byte[] inf = Inflate(slice);
                            if (inf != null) data = inf;
                        }
                        streams[num] = data;
                        objs[num] = dict;
                        if (dict.IndexOf("/ObjStm") >= 0)
                            ParseObjStm(dict, data);
                        int es2 = raw.IndexOf("endstream", streamEnd);
                        i = es2 >= 0 ? es2 + 9 : streamEnd;
                        continue;
                    }
                    objs[num] = dict;
                    int eo = raw.IndexOf("endobj", dictEnd);
                    i = eo >= 0 ? eo + 6 : dictEnd;
                    continue;
                }
                int eo2 = raw.IndexOf("endobj", p);
                i = eo2 >= 0 ? eo2 + 6 : p + 1;
            }
        }

        void ParseObjStm(string dict, byte[] data)
        {
            int n = MatchInt(dict, "/N\\s+(\\d+)");
            int first = MatchInt(dict, "/First\\s+(\\d+)");
            if (n <= 0 || first < 0 || data == null) return;
            string body = latin1.GetString(data);
            if (first > body.Length) return;
            int p = 0;
            int[] ids = new int[n];
            int[] offs = new int[n];
            try
            {
                for (int k = 0; k < n; k++)
                {
                    ids[k] = ReadInt(body, ref p);
                    offs[k] = ReadInt(body, ref p);
                }
            }
            catch
            {
                return;
            }
            for (int k = 0; k < n; k++)
            {
                int start = first + offs[k];
                int end = (k + 1 < n) ? first + offs[k + 1] : body.Length;
                if (start < 0 || end > body.Length || end < start) continue;
                if (!objs.ContainsKey(ids[k]))
                    objs[ids[k]] = body.Substring(start, end - start);
            }
        }

        static int ReadInt(string s, ref int p)
        {
            while (p < s.Length && (s[p] == ' ' || s[p] == '\r' || s[p] == '\n' || s[p] == '\t')) p++;
            int start = p;
            while (p < s.Length && char.IsDigit(s[p])) p++;
            if (start == p) throw new InvalidDataException("objstm");
            return int.Parse(s.Substring(start, p - start), CultureInfo.InvariantCulture);
        }

        static int DirectLength(string dict)
        {
            System.Text.RegularExpressions.Match ind =
                System.Text.RegularExpressions.Regex.Match(dict, "/Length\\s+\\d+\\s+0\\s+R");
            if (ind.Success) return -1;
            System.Text.RegularExpressions.Match m =
                System.Text.RegularExpressions.Regex.Match(dict, "/Length\\s+(\\d+)");
            if (!m.Success) return -1;
            return int.Parse(m.Groups[1].Value, CultureInfo.InvariantCulture);
        }

        static int FindObj(string s, int from, out int num)
        {
            num = 0;
            int i = from;
            while (i < s.Length)
            {
                int k = s.IndexOf(" 0 obj", i);
                if (k < 0) return -1;
                int j = k - 1;
                while (j >= 0 && char.IsDigit(s[j])) j--;
                if (j < k - 1 && (j < 0 || char.IsWhiteSpace(s[j]) || s[j] == '\0'))
                {
                    num = int.Parse(s.Substring(j + 1, k - (j + 1)), CultureInfo.InvariantCulture);
                    return k + 6;
                }
                i = k + 6;
            }
            return -1;
        }

        static int SkipDict(string s, int i)
        {
            int depth = 0;
            while (i < s.Length - 1)
            {
                if (s[i] == '<' && s[i + 1] == '<') { depth++; i += 2; continue; }
                if (s[i] == '>' && s[i + 1] == '>')
                {
                    depth--;
                    i += 2;
                    if (depth == 0) return i;
                    continue;
                }
                i++;
            }
            return s.Length;
        }

        static byte[] Inflate(byte[] b)
        {
            if (b == null || b.Length < 3) return null;
            int[] skips = ((b[0] & 0x0F) == 8) ? new int[] { 2, 0 } : new int[] { 0, 2 };
            for (int s = 0; s < skips.Length; s++)
            {
                int skip = skips[s];
                if (skip >= b.Length) continue;
                try
                {
                    MemoryStream ms = new MemoryStream(b, skip, b.Length - skip);
                    DeflateStream ds = new DeflateStream(ms, CompressionMode.Decompress);
                    MemoryStream os = new MemoryStream();
                    ds.CopyTo(os);
                    ds.Dispose();
                    ms.Dispose();
                    return os.ToArray();
                }
                catch
                {
                }
            }
            return null;
        }
    }

    internal sealed class Frag
    {
        public double X;
        public double Y;
        public double Width;
        public double Size;
        public int Order;
        public string Text;
        public Frag(double x, double y, string text, double width, double size, int order)
        {
            X = x;
            Y = y;
            Text = text;
            Width = width;
            Size = size;
            Order = order;
        }
    }

    internal static class Tok
    {
        public const int Eof = 0;
        public const int Num = 1;
        public const int Name = 2;
        public const int Op = 3;
        public const int Hex = 4;
        public const int Lit = 5;
        public const int Array = 6;
    }

    internal sealed class Token
    {
        public int Kind;
        public double Num;
        public string Text;
        public List<Token> Items;
    }

    internal sealed class Lexer
    {
        string s;
        int i;
        public double X;
        public double Y;
        public double Size = 7;
        public double TfSize = 1;
        public string Font = "";

        public Lexer(string s)
        {
            this.s = s;
            this.i = 0;
        }

        public Token Next()
        {
            Skip();
            Token t = new Token();
            if (i >= s.Length)
            {
                t.Kind = Tok.Eof;
                return t;
            }
            char c = s[i];
            if (c == '%')
            {
                while (i < s.Length && s[i] != '\n' && s[i] != '\r') i++;
                return Next();
            }
            if (c == '/')
            {
                i++;
                int st = i;
                while (i < s.Length && IsName(s[i])) i++;
                t.Kind = Tok.Name;
                t.Text = s.Substring(st, i - st);
                if (t.Text.Length > 0) FontWait(t.Text);
                return t;
            }
            if (c == '<')
            {
                if (i + 1 < s.Length && s[i + 1] == '<')
                {
                    i = SkipDict(i);
                    return Next();
                }
                i++;
                int st = i;
                while (i < s.Length && s[i] != '>') i++;
                t.Kind = Tok.Hex;
                t.Text = s.Substring(st, i - st);
                if (i < s.Length && s[i] == '>') i++;
                return t;
            }
            if (c == '(')
            {
                t.Kind = Tok.Lit;
                t.Text = ReadLit();
                return t;
            }
            if (c == '[')
            {
                i++;
                t.Kind = Tok.Array;
                t.Items = new List<Token>();
                while (true)
                {
                    Skip();
                    if (i >= s.Length) break;
                    if (s[i] == ']') { i++; break; }
                    Token inner = Next();
                    if (inner.Kind == Tok.Eof) break;
                    t.Items.Add(inner);
                }
                return t;
            }
            if (c == '+' || c == '-' || c == '.' || (c >= '0' && c <= '9'))
            {
                int st = i;
                if (s[i] == '+' || s[i] == '-') i++;
                bool any = false;
                while (i < s.Length && s[i] >= '0' && s[i] <= '9') { any = true; i++; }
                if (i < s.Length && s[i] == '.')
                {
                    i++;
                    while (i < s.Length && s[i] >= '0' && s[i] <= '9') { any = true; i++; }
                }
                if (!any)
                {
                    t.Kind = Tok.Op;
                    t.Text = s.Substring(st, i - st);
                    return t;
                }
                t.Kind = Tok.Num;
                double n;
                if (!double.TryParse(s.Substring(st, i - st), NumberStyles.Float, CultureInfo.InvariantCulture, out n))
                    n = 0;
                t.Num = n;
                return t;
            }
            int os = i;
            while (i < s.Length && IsOp(s[i])) i++;
            if (i == os) { i++; return Next(); }
            t.Kind = Tok.Op;
            t.Text = s.Substring(os, i - os);
            if (t.Text == "Tf") ApplyFont();
            return t;
        }

        string pendingFont;

        void FontWait(string name)
        {
            pendingFont = name;
        }

        void ApplyFont()
        {
            if (pendingFont != null && pendingFont.Length > 0) Font = pendingFont;
        }

        string ReadLit()
        {
            i++;
            StringBuilder sb = new StringBuilder();
            int depth = 1;
            while (i < s.Length && depth > 0)
            {
                char c = s[i++];
                if (c == '\\')
                {
                    if (i >= s.Length) break;
                    char n = s[i++];
                    if (n == 'n') sb.Append('\n');
                    else if (n == 'r') sb.Append('\r');
                    else if (n == 't') sb.Append('\t');
                    else if (n == 'b') sb.Append('\b');
                    else if (n == 'f') sb.Append('\f');
                    else if (n == '(' || n == ')' || n == '\\') sb.Append(n);
                    else if (n == '\r')
                    {
                        if (i < s.Length && s[i] == '\n') i++;
                    }
                    else if (n == '\n') { }
                    else if (n >= '0' && n <= '7')
                    {
                        int val = n - '0';
                        int cnt = 1;
                        while (cnt < 3 && i < s.Length && s[i] >= '0' && s[i] <= '7')
                        {
                            val = val * 8 + (s[i] - '0');
                            i++;
                            cnt++;
                        }
                        sb.Append((char)val);
                    }
                    else sb.Append(n);
                }
                else if (c == '(') { depth++; sb.Append(c); }
                else if (c == ')')
                {
                    depth--;
                    if (depth > 0) sb.Append(c);
                }
                else sb.Append(c);
            }
            return sb.ToString();
        }

        void Skip()
        {
            while (i < s.Length)
            {
                char c = s[i];
                if (c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '\0' || c == '\f') { i++; continue; }
                break;
            }
        }

        int SkipDict(int at)
        {
            int depth = 0;
            int p = at;
            while (p < s.Length - 1)
            {
                if (s[p] == '<' && s[p + 1] == '<') { depth++; p += 2; continue; }
                if (s[p] == '>' && s[p + 1] == '>')
                {
                    depth--;
                    p += 2;
                    if (depth == 0) return p;
                    continue;
                }
                p++;
            }
            return s.Length;
        }

        static bool IsName(char c)
        {
            if (c == '/' || c == '<' || c == '>' || c == '[' || c == ']' || c == '(' || c == ')' || c == '%') return false;
            if (c == ' ' || c == '\t' || c == '\r' || c == '\n' || c == '\0' || c == '\f') return false;
            return true;
        }

        static bool IsOp(char c)
        {
            if (c >= 'A' && c <= 'Z') return true;
            if (c >= 'a' && c <= 'z') return true;
            if (c == '*' || c == '\'' || c == '"') return true;
            return false;
        }
    }
}
