using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Drawing2D;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text.RegularExpressions;
using System.Xml;

namespace XlsImageShrink
{
    public class ImgStats
    {
        public int Total, Compressed, Skipped;
        public long MediaBefore, MediaAfter;
        public List<string> Notes = new List<string>();
    }

    class MediaInfo
    {
        public string Path;
        public List<string> Sources = new List<string>();
        public bool SizeKnown;
        public bool SizeUnknownRef;    // 表示サイズを取得できない参照がある
        public double InchW, InchH;    // ppi を掛ける前の値（最大）
    }

    /// <summary>OOXML パッケージ内の画像だけを差し替える。drawing XML には触れない。</summary>
    public static class ImageOptimizer
    {
        const string RelNs = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string CtNs = "http://schemas.openxmlformats.org/package/2006/content-types";
        static readonly Regex DrawingPart = new Regex(@"^xl/drawings/drawing\d+\.xml$", RegexOptions.IgnoreCase);

        public static void CheckZip(string path)
        {
            using (ZipArchive z = ZipFile.OpenRead(path))
            {
                bool ct = false;
                byte[] buf = new byte[81920];
                foreach (ZipArchiveEntry e in z.Entries)
                {
                    if (e.FullName == "[Content_Types].xml") ct = true;
                    using (Stream s = e.Open()) { while (s.Read(buf, 0, buf.Length) > 0) { } }
                }
                if (!ct) throw new InvalidDataException("[Content_Types].xml がありません");
            }
        }

        public static bool HasVbaProject(string path)
        {
            using (ZipArchive z = ZipFile.OpenRead(path))
                foreach (ZipArchiveEntry e in z.Entries)
                    if (e.FullName.EndsWith("vbaProject.bin", StringComparison.OrdinalIgnoreCase)) return true;
            return false;
        }

        public static long NonImageSize(string path)
        {
            long n = 0;
            using (ZipArchive z = ZipFile.OpenRead(path))
                foreach (ZipArchiveEntry e in z.Entries)
                    if (!e.FullName.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase)) n += e.CompressedLength + 80 + e.FullName.Length * 2;
            return n;
        }

        static byte[] ReadAll(ZipArchiveEntry e)
        {
            using (Stream s = e.Open())
            using (MemoryStream ms = new MemoryStream())
            {
                s.CopyTo(ms);
                return ms.ToArray();
            }
        }

        static string Resolve(string baseDir, string target)
        {
            if (target.StartsWith("/")) return target.TrimStart('/');
            List<string> parts = new List<string>(baseDir.Split(new char[] { '/' }, StringSplitOptions.RemoveEmptyEntries));
            foreach (string seg in target.Split('/'))
            {
                if (seg == "." || seg == "") continue;
                if (seg == "..") { if (parts.Count > 0) parts.RemoveAt(parts.Count - 1); }
                else parts.Add(seg);
            }
            return string.Join("/", parts.ToArray());
        }

        // "xl/drawings/_rels/drawing1.xml.rels" → source "xl/drawings/drawing1.xml", baseDir "xl/drawings"
        static bool RelsSource(string relsPath, out string source, out string baseDir)
        {
            source = null; baseDir = null;
            int i = relsPath.LastIndexOf("_rels/", StringComparison.Ordinal);
            if (i < 0 || !relsPath.EndsWith(".rels")) return false;
            baseDir = relsPath.Substring(0, i).TrimEnd('/');
            string file = relsPath.Substring(i + 6, relsPath.Length - (i + 6) - 5);
            source = baseDir.Length > 0 ? baseDir + "/" + file : file;
            return true;
        }

        static XmlElement Child(XmlElement el, string local)
        {
            if (el == null) return null;
            foreach (XmlNode n in el.ChildNodes)
            {
                XmlElement c = n as XmlElement;
                if (c != null && c.LocalName == local) return c;
            }
            return null;
        }

        static XmlElement Desc(XmlElement el, string local)
        {
            if (el == null) return null;
            foreach (XmlNode n in el.ChildNodes)
            {
                XmlElement c = n as XmlElement;
                if (c == null) continue;
                if (c.LocalName == local) return c;
                XmlElement r = Desc(c, local);
                if (r != null) return r;
            }
            return null;
        }

        static double Attr(XmlElement el, string name, double def)
        {
            if (el == null) return def;
            string v = el.GetAttribute(name);
            double d;
            if (!string.IsNullOrEmpty(v) && double.TryParse(v, System.Globalization.NumberStyles.Float, System.Globalization.CultureInfo.InvariantCulture, out d)) return d;
            return def;
        }

        static void WalkDrawing(XmlElement el, double sx, double sy, string baseDir, Dictionary<string, string> rid, Dictionary<string, MediaInfo> media)
        {
            foreach (XmlNode n in el.ChildNodes)
            {
                XmlElement c = n as XmlElement;
                if (c == null) continue;
                string ln = c.LocalName;
                if (ln == "grpSp")
                {
                    double nsx = sx, nsy = sy;
                    XmlElement xf = Child(Child(c, "grpSpPr"), "xfrm");
                    if (xf != null)
                    {
                        XmlElement ext = Child(xf, "ext"), ch = Child(xf, "chExt");
                        double ex = Attr(ext, "cx", 0), ey = Attr(ext, "cy", 0), cx = Attr(ch, "cx", 0), cy = Attr(ch, "cy", 0);
                        if (ex > 0 && cx > 0) nsx = sx * ex / cx;
                        if (ey > 0 && cy > 0) nsy = sy * ey / cy;
                    }
                    WalkDrawing(c, nsx, nsy, baseDir, rid, media);
                }
                else if (ln == "pic" || ln == "sp" || ln == "cxnSp")
                {
                    XmlElement bf = Desc(c, "blipFill");
                    XmlElement blip = Child(bf, "blip");
                    if (blip == null) continue;
                    string id = blip.GetAttribute("embed", RelNs);
                    string target;
                    if (string.IsNullOrEmpty(id) || !rid.TryGetValue(id, out target)) continue;
                    MediaInfo mi;
                    if (!media.TryGetValue(target, out mi)) continue;
                    XmlElement ext = Child(Child(Child(c, "spPr"), "xfrm"), "ext");
                    double w = Attr(ext, "cx", 0) * sx, h = Attr(ext, "cy", 0) * sy;
                    if (w <= 0 || h <= 0) { mi.SizeUnknownRef = true; continue; }
                    XmlElement sr = Child(bf, "srcRect");
                    double fw = 1 - (Attr(sr, "l", 0) + Attr(sr, "r", 0)) / 100000.0;
                    double fh = 1 - (Attr(sr, "t", 0) + Attr(sr, "b", 0)) / 100000.0;
                    if (fw < 0.02) fw = 0.02; if (fh < 0.02) fh = 0.02;
                    if (fw > 1) fw = 1; if (fh > 1) fh = 1;     // 負のトリミング（余白追加）は 1 とみなす
                    double iw = w / 914400.0 / fw, ih = h / 914400.0 / fh;
                    mi.SizeKnown = true;
                    if (iw > mi.InchW) mi.InchW = iw;
                    if (ih > mi.InchH) mi.InchH = ih;
                }
                else WalkDrawing(c, sx, sy, baseDir, rid, media);
            }
        }

        static ImageCodecInfo JpegCodec()
        {
            foreach (ImageCodecInfo c in ImageCodecInfo.GetImageEncoders()) if (c.FormatID == ImageFormat.Jpeg.Guid) return c;
            return null;
        }

        static bool HasAlpha(Image img)
        {
            if (!Image.IsAlphaPixelFormat(img.PixelFormat) && (img.PixelFormat & PixelFormat.Indexed) == 0) return false;
            using (Bitmap b = new Bitmap(img.Width, img.Height, PixelFormat.Format32bppArgb))
            {
                using (Graphics g = Graphics.FromImage(b))
                {
                    g.CompositingMode = CompositingMode.SourceCopy;
                    g.DrawImage(img, new Rectangle(0, 0, b.Width, b.Height), 0, 0, img.Width, img.Height, GraphicsUnit.Pixel);
                }
                BitmapData bd = b.LockBits(new Rectangle(0, 0, b.Width, b.Height), ImageLockMode.ReadOnly, PixelFormat.Format32bppArgb);
                try
                {
                    byte[] row = new byte[bd.Width * 4];
                    for (int y = 0; y < bd.Height; y++)
                    {
                        Marshal.Copy(new IntPtr(bd.Scan0.ToInt64() + (long)y * bd.Stride), row, 0, row.Length);
                        for (int x = 3; x < row.Length; x += 4) if (row[x] != 255) return true;
                    }
                }
                finally { b.UnlockBits(bd); }
            }
            return false;
        }

        /// <summary>再圧縮した結果を返す。元のまま残す場合は null。</summary>
        static byte[] Recompress(byte[] data, string ext, MediaInfo mi, bool fullMode, double ppi, int quality, int maxLong, bool pngToJpeg, out string newExt, out string note)
        {
            newExt = ext; note = null;
            bool isJpeg = ext == ".jpg" || ext == ".jpeg";
            bool isPng = ext == ".png";
            bool isBmpTif = ext == ".bmp" || ext == ".tif" || ext == ".tiff" || ext == ".dib";
            if (!isJpeg && !isPng && !isBmpTif) { note = "対象外形式(" + ext + ")"; return null; }
            using (MemoryStream ms = new MemoryStream(data))
            using (Image img = Image.FromStream(ms, false, true))
            {
                foreach (Guid g in img.FrameDimensionsList)
                    if (img.GetFrameCount(new FrameDimension(g)) > 1) { note = "複数フレーム"; return null; }
                if (isJpeg)
                {
                    int f = img.Flags;
                    if ((f & (int)ImageFlags.ColorSpaceCmyk) != 0 || (f & (int)ImageFlags.ColorSpaceYcck) != 0) { note = "CMYK のため元のまま"; return null; }
                    foreach (int id in img.PropertyIdList)
                    {
                        if (id == 0x0112)
                        {
                            PropertyItem pi = img.GetPropertyItem(0x0112);
                            if (pi.Value != null && pi.Value.Length >= 2 && BitConverter.ToUInt16(pi.Value, 0) > 1) { note = "EXIF の向き情報があるため元のまま"; return null; }
                        }
                    }
                }
                int w = img.Width, h = img.Height;
                double factor = 1.0;
                if (fullMode && mi.SizeKnown && !mi.SizeUnknownRef)
                    factor = Math.Min(1.0, Math.Max(mi.InchW * ppi / w, mi.InchH * ppi / h));
                if (maxLong > 0) factor = Math.Min(factor, (double)maxLong / Math.Max(w, h));
                if (factor > 1) factor = 1;
                int nw = Math.Max(1, (int)Math.Round(w * factor)), nh = Math.Max(1, (int)Math.Round(h * factor));
                bool resize = nw < w || nh < h;

                bool alpha = isPng && HasAlpha(img);
                bool toJpeg = isJpeg || (!alpha && pngToJpeg && fullMode);
                if (!toJpeg && !resize) return null;
                if (!toJpeg && !isPng) return null;      // BMP/TIFF は JPEG 化しないなら触らない

                using (Bitmap dst = new Bitmap(nw, nh, toJpeg ? PixelFormat.Format24bppRgb : PixelFormat.Format32bppArgb))
                {
                    try { dst.SetResolution(Math.Max(1, img.HorizontalResolution), Math.Max(1, img.VerticalResolution)); } catch { }
                    using (Graphics g = Graphics.FromImage(dst))
                    {
                        if (toJpeg) g.Clear(Color.White); else g.Clear(Color.Transparent);
                        g.InterpolationMode = InterpolationMode.HighQualityBicubic;
                        g.PixelOffsetMode = PixelOffsetMode.HighQuality;
                        g.CompositingQuality = CompositingQuality.HighQuality;
                        using (ImageAttributes ia = new ImageAttributes())
                        {
                            ia.SetWrapMode(WrapMode.TileFlipXY);
                            g.DrawImage(img, new Rectangle(0, 0, nw, nh), 0, 0, w, h, GraphicsUnit.Pixel, ia);
                        }
                    }
                    using (MemoryStream o = new MemoryStream())
                    {
                        if (toJpeg)
                        {
                            using (EncoderParameters ep = new EncoderParameters(1))
                            {
                                ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, (long)quality);
                                dst.Save(o, JpegCodec(), ep);
                            }
                            if (!isJpeg) newExt = ".jpeg";
                        }
                        else dst.Save(o, ImageFormat.Png);
                        if (o.Length >= data.Length) return null;      // 小さくならないなら元のまま
                        return o.ToArray();
                    }
                }
            }
        }

        public static ImgStats Run(string src, string dst, double ppi, int quality, int maxLong, bool pngToJpeg)
        {
            ImgStats st = new ImgStats();
            using (ZipArchive zin = ZipFile.OpenRead(src))
            {
                Dictionary<string, ZipArchiveEntry> byName = new Dictionary<string, ZipArchiveEntry>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, MediaInfo> media = new Dictionary<string, MediaInfo>(StringComparer.OrdinalIgnoreCase);
                foreach (ZipArchiveEntry e in zin.Entries)
                {
                    byName[e.FullName] = e;
                    if (e.FullName.StartsWith("xl/media/", StringComparison.OrdinalIgnoreCase) && e.Length > 0)
                    {
                        MediaInfo mi = new MediaInfo(); mi.Path = e.FullName; media[e.FullName] = mi;
                    }
                }

                // すべての .rels を読み、画像の参照元と、drawing の rId→画像 を集める
                Dictionary<string, Dictionary<string, string>> ridMaps = new Dictionary<string, Dictionary<string, string>>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, List<string>> relsOfMedia = new Dictionary<string, List<string>>(StringComparer.OrdinalIgnoreCase);
                foreach (ZipArchiveEntry e in zin.Entries)
                {
                    string source, baseDir;
                    if (!RelsSource(e.FullName, out source, out baseDir)) continue;
                    XmlDocument doc = new XmlDocument();
                    try { using (Stream s = e.Open()) doc.Load(s); } catch { continue; }
                    Dictionary<string, string> map = new Dictionary<string, string>();
                    foreach (XmlNode n in doc.DocumentElement.ChildNodes)
                    {
                        XmlElement r = n as XmlElement;
                        if (r == null || r.LocalName != "Relationship") continue;
                        if (string.Equals(r.GetAttribute("TargetMode"), "External", StringComparison.OrdinalIgnoreCase)) continue;
                        string t = Resolve(baseDir, r.GetAttribute("Target"));
                        map[r.GetAttribute("Id")] = t;
                        MediaInfo mi;
                        if (media.TryGetValue(t, out mi))
                        {
                            if (!mi.Sources.Contains(source)) mi.Sources.Add(source);
                            List<string> l;
                            if (!relsOfMedia.TryGetValue(mi.Path, out l)) { l = new List<string>(); relsOfMedia[mi.Path] = l; }
                            if (!l.Contains(e.FullName)) l.Add(e.FullName);
                        }
                    }
                    ridMaps[source] = map;
                }

                // drawing XML から表示サイズを取得
                foreach (KeyValuePair<string, Dictionary<string, string>> kv in ridMaps)
                {
                    if (!DrawingPart.IsMatch(kv.Key)) continue;
                    ZipArchiveEntry de;
                    if (!byName.TryGetValue(kv.Key, out de)) continue;
                    XmlDocument doc = new XmlDocument();
                    try { using (Stream s = de.Open()) doc.Load(s); } catch { continue; }
                    WalkDrawing(doc.DocumentElement, 1, 1, "xl/drawings", kv.Value, media);
                }

                // 画像ごとに再圧縮
                Dictionary<string, byte[]> replaced = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                Dictionary<string, string> renamed = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
                HashSet<string> usedNames = new HashSet<string>(byName.Keys, StringComparer.OrdinalIgnoreCase);
                foreach (MediaInfo mi in media.Values)
                {
                    ZipArchiveEntry e = byName[mi.Path];
                    st.Total++;
                    st.MediaBefore += e.Length;
                    string name = Path.GetFileName(mi.Path);
                    if (mi.Sources.Count == 0) { st.Skipped++; st.MediaAfter += e.Length; continue; }
                    bool chart = false, allDrawing = true;
                    foreach (string s in mi.Sources)
                    {
                        if (s.IndexOf("/charts/", StringComparison.OrdinalIgnoreCase) >= 0 || s.IndexOf("/diagrams/", StringComparison.OrdinalIgnoreCase) >= 0) chart = true;
                        if (!DrawingPart.IsMatch(s)) allDrawing = false;
                    }
                    if (chart) { st.Skipped++; st.MediaAfter += e.Length; st.Notes.Add(name + ": グラフ内の画像のため対象外"); continue; }
                    string ext = Path.GetExtension(mi.Path).ToLowerInvariant();
                    byte[] data = ReadAll(e);
                    byte[] res = null; string newExt = ext, note = null;
                    try { res = Recompress(data, ext, mi, allDrawing, ppi, quality, maxLong, pngToJpeg, out newExt, out note); }
                    catch (Exception ex) { note = "読み込めないため元のまま (" + ex.GetType().Name + ")"; res = null; }
                    if (res == null)
                    {
                        st.Skipped++; st.MediaAfter += e.Length;
                        if (note != null) st.Notes.Add(name + ": " + note);
                        continue;
                    }
                    st.Compressed++; st.MediaAfter += res.Length;
                    replaced[mi.Path] = res;
                    if (newExt != ext)
                    {
                        string stem = mi.Path.Substring(0, mi.Path.Length - ext.Length);
                        string np = stem + newExt; int k = 1;
                        while (usedNames.Contains(np)) { np = stem + "_c" + k + newExt; k++; }
                        usedNames.Add(np);
                        renamed[mi.Path] = np;
                    }
                }

                // 拡張子が変わった画像の rels と Content_Types を書き換える
                Dictionary<string, byte[]> xmlOverride = new Dictionary<string, byte[]>(StringComparer.OrdinalIgnoreCase);
                if (renamed.Count > 0)
                {
                    HashSet<string> relsToFix = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
                    foreach (string m in renamed.Keys) foreach (string r in relsOfMedia[m]) relsToFix.Add(r);
                    foreach (string rp in relsToFix)
                    {
                        string source, baseDir;
                        RelsSource(rp, out source, out baseDir);
                        XmlDocument doc = new XmlDocument();
                        doc.PreserveWhitespace = true;
                        using (Stream s = byName[rp].Open()) doc.Load(s);
                        foreach (XmlNode n in doc.DocumentElement.ChildNodes)
                        {
                            XmlElement r = n as XmlElement;
                            if (r == null || r.LocalName != "Relationship") continue;
                            string target = r.GetAttribute("Target");
                            string np;
                            if (renamed.TryGetValue(Resolve(baseDir, target), out np))
                            {
                                int slash = target.LastIndexOf('/');
                                r.SetAttribute("Target", (slash >= 0 ? target.Substring(0, slash + 1) : "") + Path.GetFileName(np));
                            }
                        }
                        using (MemoryStream ms = new MemoryStream()) { doc.Save(ms); xmlOverride[rp] = ms.ToArray(); }
                    }
                    XmlDocument ct = new XmlDocument();
                    ct.PreserveWhitespace = true;
                    using (Stream s = byName["[Content_Types].xml"].Open()) ct.Load(s);
                    bool has = false;
                    foreach (XmlNode n in ct.DocumentElement.ChildNodes)
                    {
                        XmlElement d = n as XmlElement;
                        if (d != null && d.LocalName == "Default" && string.Equals(d.GetAttribute("Extension"), "jpeg", StringComparison.OrdinalIgnoreCase)) has = true;
                    }
                    if (!has)
                    {
                        XmlElement d = ct.CreateElement("Default", CtNs);
                        d.SetAttribute("Extension", "jpeg");
                        d.SetAttribute("ContentType", "image/jpeg");
                        ct.DocumentElement.InsertBefore(d, ct.DocumentElement.FirstChild);
                    }
                    using (MemoryStream ms = new MemoryStream()) { ct.Save(ms); xmlOverride["[Content_Types].xml"] = ms.ToArray(); }
                }

                // 元のエントリ順のまま新しい ZIP を作る
                if (File.Exists(dst)) File.Delete(dst);
                using (FileStream fs = new FileStream(dst, FileMode.CreateNew))
                using (ZipArchive zout = new ZipArchive(fs, ZipArchiveMode.Create))
                {
                    foreach (ZipArchiveEntry e in zin.Entries)
                    {
                        string name = e.FullName;
                        byte[] bytes = null;
                        string outName = name;
                        if (replaced.TryGetValue(name, out bytes)) { string np; if (renamed.TryGetValue(name, out np)) outName = np; }
                        else if (!xmlOverride.TryGetValue(name, out bytes)) bytes = null;
                        ZipArchiveEntry ne = zout.CreateEntry(outName, CompressionLevel.Optimal);
                        try { ne.LastWriteTime = e.LastWriteTime; } catch { }
                        using (Stream os = ne.Open())
                        {
                            if (bytes != null) os.Write(bytes, 0, bytes.Length);
                            else using (Stream ins = e.Open()) ins.CopyTo(os);
                        }
                    }
                }
            }
            return st;
        }
    }

    public class SizeResult
    {
        public string Path;
        public long Size;
        public double Ppi; public int Quality;
        public ImgStats Stats;
        public bool Changed;       // 画像を圧縮した結果を採用したか
        public bool Unachieved;    // サイズ上限に届かなかった
    }

    public static class SizeController
    {
        static int seq = 0;

        static SizeResult Try(string src, string wd, double ppi, int q, Settings s, Logger log)
        {
            string dst = Path.Combine(wd, "c" + (++seq) + Path.GetExtension(src));
            ImgStats st = ImageOptimizer.Run(src, dst, ppi, q, s.MaxLongSide, s.PngToJpeg);
            SizeResult r = new SizeResult();
            r.Path = dst; r.Size = new FileInfo(dst).Length; r.Ppi = ppi; r.Quality = q; r.Stats = st; r.Changed = true;
            if (log != null) log.Info("  圧縮試行 ppi=" + ppi.ToString("0") + " 品質=" + q + " → " + Util.FormatSize(r.Size) + " (" + r.Size + " bytes)");
            return r;
        }

        static SizeResult NoChange(string src, SizeResult tried)
        {
            SizeResult r = new SizeResult();
            r.Path = src; r.Size = new FileInfo(src).Length; r.Changed = false;
            r.Stats = tried != null ? tried.Stats : new ImgStats();
            if (tried != null) { r.Stats.Skipped = r.Stats.Total; r.Stats.Compressed = 0; }
            return r;
        }

        public static SizeResult Optimize(string src, string wd, Settings s, Logger log)
        {
            long s0 = new FileInfo(src).Length;
            double p0 = s.Ppi; int q0 = s.JpegQuality;
            if (!s.SizeTargetEnabled)
            {
                SizeResult r = Try(src, wd, p0, q0, s, log);
                return Better(r, s0) ? r : NoChange(src, r);
            }
            long limit = (long)(s.SizeLimitMB * 1024 * 1024);
            if (s0 <= limit)
            {
                if (s.SkipIfUnderLimit) { if (log != null) log.Info("  既に上限以下のため圧縮しません"); return NoChange(src, null); }
                SizeResult r = Try(src, wd, p0, q0, s, log);
                return Better(r, s0) ? r : NoChange(src, r);
            }
            if (ImageOptimizer.NonImageSize(src) > limit)
            {
                if (log != null) log.Warn("  画像以外だけで上限を超えるため、上限には届きません");
                SizeResult r = Try(src, wd, p0, q0, s, log);
                if (!Better(r, s0)) r = NoChange(src, r);
                r.Unachieved = true;
                return Unachieved(src, r, s);
            }
            double pmin = Math.Min(s.MinPpi, p0); int qmin = Math.Min(s.MinQuality, q0);
            List<double[]> stages = new List<double[]>();
            AddStage(stages, p0, q0, pmin, qmin, p0, q0);
            AddStage(stages, p0, 75, pmin, qmin, p0, q0);
            AddStage(stages, p0, 65, pmin, qmin, p0, q0);
            AddStage(stages, 0.8 * p0, 65, pmin, qmin, p0, q0);
            AddStage(stages, 0.8 * p0, 55, pmin, qmin, p0, q0);
            AddStage(stages, pmin, 55, pmin, qmin, p0, q0);
            AddStage(stages, pmin, qmin, pmin, qmin, p0, q0);
            SizeResult prev = null, cur = null;
            foreach (double[] st in stages)
            {
                cur = Try(src, wd, st[0], (int)st[1], s, log);
                if (cur.Stats.Compressed == 0) break;      // 圧縮できる画像がない
                if (cur.Size <= limit)
                {
                    if (prev != null) cur = Refine(src, wd, prev, cur, limit, s, log);
                    return cur;
                }
                prev = cur;
            }
            if (!Better(cur, s0)) cur = NoChange(src, cur);
            cur.Unachieved = true;
            return Unachieved(src, cur, s);
        }

        // 画像を 1 枚以上圧縮でき、かつ小さくなったときだけ採用する
        static bool Better(SizeResult r, long s0) { return r.Stats.Compressed > 0 && r.Size < s0; }

        static SizeResult Unachieved(string src, SizeResult r, Settings s)
        {
            if (s.UnachievedMode == "convertOnly") { SizeResult n = NoChange(src, r); n.Unachieved = true; return n; }
            return r;
        }

        static void AddStage(List<double[]> stages, double p, double q, double pmin, int qmin, double p0, int q0)
        {
            p = Math.Round(Math.Max(pmin, Math.Min(p0, p)));
            q = Math.Max(qmin, Math.Min(q0, q));
            if (stages.Count > 0)
            {
                double[] last = stages[stages.Count - 1];
                if (p > last[0]) p = last[0];
                if (q > last[1]) q = last[1];
                if (p == last[0] && q == last[1]) return;
            }
            stages.Add(new double[] { p, q });
        }

        // over（上限超過）と under（上限以下）の間を二分探索し、上限以下で最も画質の高い結果を返す
        static SizeResult Refine(string src, string wd, SizeResult over, SizeResult under, long limit, Settings s, Logger log)
        {
            SizeResult lo = under, hi = over;
            for (int i = 0; i < 4; i++)
            {
                double mp = Math.Round((lo.Ppi + hi.Ppi) / 2.0);
                int mq = (lo.Quality + hi.Quality) / 2;
                if ((mp == lo.Ppi && mq == lo.Quality) || (mp == hi.Ppi && mq == hi.Quality)) break;
                SizeResult t = Try(src, wd, mp, mq, s, log);
                if (t.Size <= limit) lo = t; else hi = t;
            }
            return lo;
        }
    }
}
