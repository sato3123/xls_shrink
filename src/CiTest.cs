using System;
using System.Collections.Generic;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Security.Cryptography;
using System.Text;

namespace XlsImageShrink
{
    /// <summary>
    /// Excel を使わずに実行できるテスト（GitHub Actions など Excel のない環境用）。
    /// XlsImageShrink.exe --citest フォルダ
    /// 画像圧縮・サイズ制御・スキャン・形式判定・圧縮のみのジョブ・元ファイル処理・復元・設定を確認する。
    /// </summary>
    public static class CiTest
    {
        static string root, repPath;
        static int pass, fail;

        const string NsMain = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
        const string NsR = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
        const string NsPkg = "http://schemas.openxmlformats.org/package/2006/relationships";
        const string NsXdr = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing";
        const string NsA = "http://schemas.openxmlformats.org/drawingml/2006/main";
        const long Emu = 914400;

        static void Out(string line) { File.AppendAllText(repPath, line + "\r\n", new UTF8Encoding(false)); }
        static void Check(string id, string desc, bool ok, string detail)
        {
            if (ok) pass++; else fail++;
            Out((ok ? "PASS " : "FAIL ") + id + " " + desc + (string.IsNullOrEmpty(detail) ? "" : "  [" + detail + "]"));
        }

        public static int Run(string dir)
        {
            root = Path.GetFullPath(dir);
            if (Directory.Exists(root))
            {
                foreach (string f in Directory.GetFiles(root, "*", SearchOption.AllDirectories)) { try { File.SetAttributes(f, FileAttributes.Normal); } catch { } }
                Directory.Delete(root, true);
            }
            Directory.CreateDirectory(root);
            repPath = Path.Combine(root, "citest_report.txt");
            AppPaths.Init(Path.Combine(root, "app"));
            Util.NoRecycle = true;
            Out("XlsImageShrink CI テスト（Excel 不要）  " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            Out("OS: " + Environment.OSVersion + "  .NET CLR " + Environment.Version + "  64bit プロセス=" + Environment.Is64BitProcess);
            Step("画像圧縮", TestImages);
            Step("サイズ目標", TestSize);
            Step("形式判定・スキャン", TestScan);
            Step("圧縮のみジョブ", TestCompressJob);
            Step("移動と復元", TestRestore);
            Step("設定", TestSettings);
            Step("画面（フォルダ選択で自動検索）", TestUi);
            Out("");
            Out("合計: PASS " + pass + " / FAIL " + fail);
            Out("END");
            return fail == 0 ? 0 : 1;
        }

        static void Step(string name, Action a)
        {
            Out("--- " + name);
            try { a(); }
            catch (Exception ex) { fail++; Out("FAIL " + name + " 例外: " + ex); }
        }

        // ================= テスト用 xlsx の生成 =================
        static byte[] NoiseImage(int w, int h, int seed, ImageFormat fmt, bool alpha)
        {
            using (Bitmap b = new Bitmap(w, h, alpha ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb))
            {
                int bpp = alpha ? 4 : 3;
                BitmapData bd = b.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, b.PixelFormat);
                byte[] row = new byte[bd.Stride];
                byte[] noise = new byte[w * bpp];
                Random rnd = new Random(seed);
                for (int y = 0; y < h; y++)
                {
                    rnd.NextBytes(noise);
                    for (int x = 0; x < w; x++)
                    {
                        int i = x * bpp, g = (x * 255 / w + y * 255 / h) / 2;
                        row[i] = (byte)((g + noise[i]) / 2);
                        row[i + 1] = (byte)((255 - g + noise[i + 1]) / 2);
                        row[i + 2] = (byte)((x * 255 / w + noise[i + 2]) / 2);
                        if (alpha) row[i + 3] = (byte)(x < w / 3 ? 0 : 255);
                    }
                    Marshal.Copy(row, 0, new IntPtr(bd.Scan0.ToInt64() + (long)y * bd.Stride), bd.Stride);
                }
                b.UnlockBits(bd);
                using (MemoryStream ms = new MemoryStream())
                {
                    if (fmt.Guid == ImageFormat.Jpeg.Guid)
                    {
                        ImageCodecInfo codec = null;
                        foreach (ImageCodecInfo c in ImageCodecInfo.GetImageEncoders()) if (c.FormatID == ImageFormat.Jpeg.Guid) codec = c;
                        using (EncoderParameters ep = new EncoderParameters(1))
                        {
                            ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 92L);
                            b.Save(ms, codec, ep);
                        }
                    }
                    else b.Save(ms, fmt);
                    return ms.ToArray();
                }
            }
        }

        static string Pic(int id, string rid, long cx, long cy, string srcRect)
        {
            return "<xdr:pic><xdr:nvPicPr><xdr:cNvPr id=\"" + id + "\" name=\"Picture " + id + "\"/><xdr:cNvPicPr/></xdr:nvPicPr>"
                + "<xdr:blipFill><a:blip r:embed=\"" + rid + "\"/>" + (srcRect ?? "") + "<a:stretch><a:fillRect/></a:stretch></xdr:blipFill>"
                + "<xdr:spPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" + cx + "\" cy=\"" + cy + "\"/></a:xfrm><a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>";
        }
        static string Anchor(int row, string body)
        {
            return "<xdr:twoCellAnchor editAs=\"oneCell\"><xdr:from><xdr:col>0</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" + row + "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>"
                + "<xdr:to><xdr:col>5</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>" + (row + 10) + "</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:to>" + body + "<xdr:clientData/></xdr:twoCellAnchor>";
        }

        /// <summary>
        /// 画像の表示幅（インチ）:
        /// image1.jpg 3000x2000 を 4 インチ → 150ppi で 600px
        /// image2.png 透過なし 1500x1000 を 3 インチ → JPEG 化して 450px
        /// image3.png 透過あり 900x600 を 2.5 インチ → PNG のまま 375px
        /// image4.jpg 2000x1000 を左右 25% トリミングして 2 インチ → 見える部分 2 インチ = 全体 4 インチ相当 → 600px
        /// image5.jpg 2000x1000 を 0.5 倍のグループ内で 4 インチ → 実際は 2 インチ → 300px
        /// image6.jpg 2000x1000 を 1 インチと 3 インチの 2 か所で使用 → 大きい方 → 450px
        /// image7.gif は対象外形式 → 不変
        /// </summary>
        static void MakeXlsx(string path, bool withVba)
        {
            StringBuilder dr = new StringBuilder();
            dr.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>");
            dr.Append("<xdr:wsDr xmlns:xdr=\"" + NsXdr + "\" xmlns:a=\"" + NsA + "\" xmlns:r=\"" + NsR + "\">");
            dr.Append(Anchor(0, Pic(2, "rId1", 4 * Emu, 4 * Emu * 2 / 3, null)));
            dr.Append(Anchor(12, Pic(3, "rId2", 3 * Emu, 2 * Emu, null)));
            dr.Append(Anchor(24, Pic(4, "rId3", 5 * Emu / 2, 5 * Emu / 3, null)));
            dr.Append(Anchor(36, Pic(5, "rId4", 2 * Emu, 2 * Emu, "<a:srcRect l=\"25000\" r=\"25000\"/>")));
            dr.Append(Anchor(48, "<xdr:grpSp><xdr:nvGrpSpPr><xdr:cNvPr id=\"20\" name=\"Group 20\"/><xdr:cNvGrpSpPr/></xdr:nvGrpSpPr>"
                + "<xdr:grpSpPr><a:xfrm><a:off x=\"0\" y=\"0\"/><a:ext cx=\"" + (2 * Emu) + "\" cy=\"" + Emu + "\"/><a:chOff x=\"0\" y=\"0\"/><a:chExt cx=\"" + (4 * Emu) + "\" cy=\"" + (2 * Emu) + "\"/></a:xfrm></xdr:grpSpPr>"
                + Pic(6, "rId5", 4 * Emu, 2 * Emu, null) + "</xdr:grpSp>"));
            dr.Append(Anchor(60, Pic(7, "rId6", Emu, Emu / 2, null)));
            dr.Append(Anchor(72, Pic(8, "rId6", 3 * Emu, 3 * Emu / 2, null)));
            dr.Append(Anchor(84, Pic(9, "rId7", Emu, Emu, null)));
            dr.Append("</xdr:wsDr>");

            StringBuilder drRels = new StringBuilder();
            drRels.Append("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + NsPkg + "\">");
            string[] media = { "image1.jpg", "image2.png", "image3.png", "image4.jpg", "image5.jpg", "image6.jpg", "image7.gif" };
            for (int i = 0; i < media.Length; i++)
                drRels.Append("<Relationship Id=\"rId" + (i + 1) + "\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/image\" Target=\"../media/" + media[i] + "\"/>");
            drRels.Append("</Relationships>");

            string ct = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\">"
                + "<Default Extension=\"rels\" ContentType=\"application/vnd.openxmlformats-package.relationships+xml\"/>"
                + "<Default Extension=\"xml\" ContentType=\"application/xml\"/>"
                + "<Default Extension=\"jpg\" ContentType=\"image/jpeg\"/>"
                + "<Default Extension=\"png\" ContentType=\"image/png\"/>"
                + "<Default Extension=\"gif\" ContentType=\"image/gif\"/>"
                + (withVba ? "<Default Extension=\"bin\" ContentType=\"application/vnd.ms-office.vbaProject\"/>" : "")
                + "<Override PartName=\"/xl/workbook.xml\" ContentType=\"application/vnd." + (withVba ? "ms-excel.sheet.macroEnabled.main+xml" : "openxmlformats-officedocument.spreadsheetml.sheet.main+xml") + "\"/>"
                + "<Override PartName=\"/xl/worksheets/sheet1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml\"/>"
                + "<Override PartName=\"/xl/drawings/drawing1.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.drawing+xml\"/>"
                + "</Types>";

            Dictionary<string, byte[]> parts = new Dictionary<string, byte[]>();
            List<string> order = new List<string>();
            Action<string, byte[]> add = delegate(string n, byte[] b) { parts[n] = b; order.Add(n); };
            Encoding u8 = new UTF8Encoding(false);
            add("[Content_Types].xml", u8.GetBytes(ct));
            add("_rels/.rels", u8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + NsPkg + "\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument\" Target=\"xl/workbook.xml\"/></Relationships>"));
            add("xl/workbook.xml", u8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><workbook xmlns=\"" + NsMain + "\" xmlns:r=\"" + NsR + "\"><sheets><sheet name=\"写真\" sheetId=\"1\" r:id=\"rId1\"/></sheets></workbook>"));
            add("xl/_rels/workbook.xml.rels", u8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + NsPkg + "\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet\" Target=\"worksheets/sheet1.xml\"/>"
                + (withVba ? "<Relationship Id=\"rId9\" Type=\"http://schemas.microsoft.com/office/2006/relationships/vbaProject\" Target=\"vbaProject.bin\"/>" : "") + "</Relationships>"));
            add("xl/worksheets/sheet1.xml", u8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><worksheet xmlns=\"" + NsMain + "\" xmlns:r=\"" + NsR + "\"><sheetData><row r=\"1\"><c r=\"A1\" t=\"inlineStr\"><is><t>点検報告書</t></is></c></row></sheetData><drawing r:id=\"rId1\"/></worksheet>"));
            add("xl/worksheets/_rels/sheet1.xml.rels", u8.GetBytes("<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><Relationships xmlns=\"" + NsPkg + "\"><Relationship Id=\"rId1\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing\" Target=\"../drawings/drawing1.xml\"/></Relationships>"));
            add("xl/drawings/drawing1.xml", u8.GetBytes(dr.ToString()));
            add("xl/drawings/_rels/drawing1.xml.rels", u8.GetBytes(drRels.ToString()));
            if (withVba)
            {
                byte[] vba = new byte[5000]; new Random(77).NextBytes(vba);
                add("xl/vbaProject.bin", vba);
            }
            add("xl/media/image1.jpg", NoiseImage(3000, 2000, 1, ImageFormat.Jpeg, false));
            add("xl/media/image2.png", NoiseImage(1500, 1000, 2, ImageFormat.Png, false));
            add("xl/media/image3.png", NoiseImage(900, 600, 3, ImageFormat.Png, true));
            add("xl/media/image4.jpg", NoiseImage(2000, 1000, 4, ImageFormat.Jpeg, false));
            add("xl/media/image5.jpg", NoiseImage(2000, 1000, 5, ImageFormat.Jpeg, false));
            add("xl/media/image6.jpg", NoiseImage(2000, 1000, 6, ImageFormat.Jpeg, false));
            add("xl/media/image7.gif", NoiseImage(300, 200, 7, ImageFormat.Gif, false));

            if (File.Exists(path)) File.Delete(path);
            using (FileStream fs = new FileStream(path, FileMode.CreateNew))
            using (ZipArchive z = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                foreach (string n in order)
                {
                    ZipArchiveEntry e = z.CreateEntry(n, CompressionLevel.Optimal);
                    using (Stream s = e.Open()) s.Write(parts[n], 0, parts[n].Length);
                }
            }
        }

        // ================= 読み取り補助 =================
        static Dictionary<string, byte[]> ReadZip(string path)
        {
            Dictionary<string, byte[]> d = new Dictionary<string, byte[]>(StringComparer.Ordinal);
            using (ZipArchive z = ZipFile.OpenRead(path))
                foreach (ZipArchiveEntry e in z.Entries)
                    using (Stream s = e.Open())
                    using (MemoryStream ms = new MemoryStream()) { s.CopyTo(ms); d[e.FullName] = ms.ToArray(); }
            return d;
        }
        static List<string> EntryOrder(string path)
        {
            List<string> l = new List<string>();
            using (ZipArchive z = ZipFile.OpenRead(path)) foreach (ZipArchiveEntry e in z.Entries) l.Add(e.FullName);
            return l;
        }
        static Size Dim(byte[] img)
        {
            using (MemoryStream ms = new MemoryStream(img))
            using (Image im = Image.FromStream(ms)) return im.Size;
        }
        static bool IsAlphaPng(byte[] img)
        {
            using (MemoryStream ms = new MemoryStream(img))
            using (Image im = Image.FromStream(ms)) return Image.IsAlphaPixelFormat(im.PixelFormat);
        }
        static bool Eq(byte[] a, byte[] b)
        {
            if (a == null || b == null || a.Length != b.Length) return false;
            for (int i = 0; i < a.Length; i++) if (a[i] != b[i]) return false;
            return true;
        }
        static string Txt(byte[] b) { return Encoding.UTF8.GetString(b); }
        static bool Near(int v, int expect) { return Math.Abs(v - expect) <= 2; }
        static string Hash(string p) { return Util.Sha256(p); }

        // ================= テスト =================
        static void TestImages()
        {
            string src = Path.Combine(root, "img_src.xlsx"), dst = Path.Combine(root, "img_dst.xlsx");
            MakeXlsx(src, false);
            ImgStats st = ImageOptimizer.Run(src, dst, 150, 85, 0, true);
            Dictionary<string, byte[]> a = ReadZip(src), b = ReadZip(dst);
            Out("  入力 " + new FileInfo(src).Length + " bytes → 出力 " + new FileInfo(dst).Length + " bytes  画像 " + st.Compressed + "/" + st.Total + " 圧縮");

            Check("T-IMG-18a", "出力 ZIP が読める", true, "");
            bool zipOk = true; string zmsg = "";
            try { ImageOptimizer.CheckZip(dst); } catch (Exception ex) { zipOk = false; zmsg = ex.Message; }
            Check("T-IMG-18", "出力 ZIP の全エントリが展開できる", zipOk, zmsg);
            Check("T-IMG-00", "ファイルサイズが減る", new FileInfo(dst).Length < new FileInfo(src).Length, "");
            Check("T-IMG-STAT", "画像 7 枚中 6 枚を圧縮、1 枚（GIF）は対象外", st.Total == 7 && st.Compressed == 6 && st.Skipped == 1, st.Compressed + "/" + st.Total);

            Size s1 = b.ContainsKey("xl/media/image1.jpg") ? Dim(b["xl/media/image1.jpg"]) : Size.Empty;
            Check("T-IMG-01", "150ppi・表示 4 インチ → 幅 600px", Near(s1.Width, 600), s1.ToString());
            Check("T-IMG-02", "縦横比を保つ（3000x2000 → 600x400）", Near(s1.Height, 400), s1.ToString());

            Check("T-IMG-09", "透過なし PNG は JPEG 化（image2.png → image2.jpeg）", b.ContainsKey("xl/media/image2.jpeg") && !b.ContainsKey("xl/media/image2.png"), "");
            if (b.ContainsKey("xl/media/image2.jpeg"))
            {
                Size s2 = Dim(b["xl/media/image2.jpeg"]);
                Check("T-IMG-09b", "JPEG 化した画像も 150ppi（3 インチ → 450px）", Near(s2.Width, 450), s2.ToString());
            }
            string rels = Txt(b["xl/drawings/_rels/drawing1.xml.rels"]);
            Check("T-IMG-09c", "rels の参照先が新しい拡張子に更新される", rels.IndexOf("../media/image2.jpeg") >= 0 && rels.IndexOf("../media/image2.png") < 0, "");
            string ct = Txt(b["[Content_Types].xml"]);
            Check("T-IMG-09d", "Content_Types に jpeg の Default が追加される", ct.IndexOf("Extension=\"jpeg\"") >= 0, "");

            Check("T-IMG-08", "透過 PNG は PNG のまま、透過を保持", b.ContainsKey("xl/media/image3.png") && IsAlphaPng(b["xl/media/image3.png"]), "");
            if (b.ContainsKey("xl/media/image3.png"))
            {
                Size s3 = Dim(b["xl/media/image3.png"]);
                Check("T-IMG-08b", "透過 PNG も縮小（2.5 インチ → 375px）", Near(s3.Width, 375), s3.ToString());
            }
            Size s4 = Dim(b["xl/media/image4.jpg"]);
            Check("T-IMG-05", "トリミング画像は見える部分で ppi を満たす（600px）", Near(s4.Width, 600), s4.ToString());
            Size s5 = Dim(b["xl/media/image5.jpg"]);
            Check("T-IMG-GRP", "グループ内の画像はグループの縮尺を考慮（300px）", Near(s5.Width, 300), s5.ToString());
            Size s6 = Dim(b["xl/media/image6.jpg"]);
            Check("T-IMG-07", "共有画像は最も大きい表示に合わせる（450px）", Near(s6.Width, 450), s6.ToString());
            Check("T-IMG-11", "対象外形式（GIF）はバイト単位で不変", Eq(a["xl/media/image7.gif"], b["xl/media/image7.gif"]), "");

            Check("T-LAY-15", "drawing XML はバイト単位で不変", Eq(a["xl/drawings/drawing1.xml"], b["xl/drawings/drawing1.xml"]), "");
            bool others = true; string diff = "";
            foreach (string k in a.Keys)
            {
                if (k.StartsWith("xl/media/") || k == "[Content_Types].xml" || k == "xl/drawings/_rels/drawing1.xml.rels") continue;
                if (!b.ContainsKey(k) || !Eq(a[k], b[k])) { others = false; diff += k + " "; }
            }
            Check("T-CMP-03", "画像・rels・Content_Types 以外のパーツはバイト単位で不変", others, diff);
            List<string> oa = EntryOrder(src), ob = EntryOrder(dst);
            Check("T-IMG-ORDER", "エントリ順を保ち [Content_Types].xml が先頭", ob.Count == oa.Count && ob[0] == "[Content_Types].xml", string.Join(",", ob.ToArray()));

            // 拡大しない
            string dst2 = Path.Combine(root, "img_dst_hi.xlsx");
            ImageOptimizer.Run(src, dst2, 600, 95, 0, false);
            Dictionary<string, byte[]> c = ReadZip(dst2);
            Size h3 = Dim(c["xl/media/image3.png"]);
            Check("T-IMG-03", "必要画素数より小さい画像は拡大しない（900px のまま）", h3.Width == 900 && Eq(a["xl/media/image3.png"], c["xl/media/image3.png"]), h3.ToString());
            Check("T-IMG-10", "JPEG 化 OFF では PNG は PNG のまま", c.ContainsKey("xl/media/image2.png") && !c.ContainsKey("xl/media/image2.jpeg"), "");

            // 長辺上限
            string dst3 = Path.Combine(root, "img_dst_long.xlsx");
            ImageOptimizer.Run(src, dst3, 600, 85, 400, true);
            Size l1 = Dim(ReadZip(dst3)["xl/media/image1.jpg"]);
            Check("T-IMG-04", "長辺上限 400px", l1.Width <= 401 && l1.Height <= 401, l1.ToString());

            // 二重劣化防止の前提: 同じ入力から同じ設定なら同じ結果
            string dst4 = Path.Combine(root, "img_dst_again.xlsx");
            ImageOptimizer.Run(src, dst4, 150, 85, 0, true);
            Check("T-IMG-17", "同じ入力・同じ設定なら画像は同一（常に元画像から再圧縮）", Eq(ReadZip(dst4)["xl/media/image1.jpg"], b["xl/media/image1.jpg"]), "");

            // 破損画像
            string bad = Path.Combine(root, "img_bad.xlsx");
            File.Copy(src, bad);
            using (ZipArchive z = ZipFile.Open(bad, ZipArchiveMode.Update))
            {
                z.GetEntry("xl/media/image5.jpg").Delete();
                ZipArchiveEntry e = z.CreateEntry("xl/media/image5.jpg");
                using (Stream s = e.Open()) { byte[] junk = new byte[2000]; new Random(3).NextBytes(junk); junk[0] = 0xFF; junk[1] = 0xD8; s.Write(junk, 0, junk.Length); }
            }
            string badOut = Path.Combine(root, "img_bad_out.xlsx");
            ImgStats st5 = ImageOptimizer.Run(bad, badOut, 150, 85, 0, true);
            Dictionary<string, byte[]> bb = ReadZip(bad), bo = ReadZip(badOut);
            Check("T-IMG-13", "壊れた画像は元のまま、他の画像は圧縮", Eq(bb["xl/media/image5.jpg"], bo["xl/media/image5.jpg"]) && st5.Compressed == 5, "圧縮=" + st5.Compressed + " " + string.Join(" / ", st5.Notes.ToArray()));

            // マクロ（vbaProject.bin）を保持
            string vsrc = Path.Combine(root, "vba_src.xlsm"), vdst = Path.Combine(root, "vba_dst.xlsm");
            MakeXlsx(vsrc, true);
            ImageOptimizer.Run(vsrc, vdst, 150, 85, 0, true);
            Check("T-CMP-04", "xlsm の vbaProject.bin はバイト単位で不変", Eq(ReadZip(vsrc)["xl/vbaProject.bin"], ReadZip(vdst)["xl/vbaProject.bin"]) && ImageOptimizer.HasVbaProject(vdst), "");
            Check("T-MAC-12c", "xlsm の ContentType（macroEnabled）を保持", Txt(ReadZip(vdst)["[Content_Types].xml"]).IndexOf("macroEnabled") >= 0, "");
        }

        static void TestSize()
        {
            string src = Path.Combine(root, "size_src.xlsx");
            MakeXlsx(src, false);
            long s0 = new FileInfo(src).Length;
            string wd = Path.Combine(root, "size_wd"); Directory.CreateDirectory(wd);

            Settings s = new Settings();
            SizeResult r0 = SizeController.Optimize(src, wd, s, null);
            Check("T-SIZ-08", "サイズ目標なし: 指定 ppi・品質を 1 回だけ適用", r0.Changed && r0.Ppi == 150 && r0.Quality == 85 && !r0.Unachieved, "ppi=" + r0.Ppi + " q=" + r0.Quality + " " + r0.Size);

            s.SizeTargetEnabled = true; s.SkipIfUnderLimit = true;
            long limit = (long)(r0.Size * 0.75);
            s.SizeLimitMB = limit / (1024.0 * 1024.0);
            SizeResult r1 = SizeController.Optimize(src, wd, s, null);
            Check("T-SIZ-02", "上限以下になるまで段階的に下げる", r1.Changed && r1.Size <= limit && (r1.Quality < 85 || r1.Ppi < 150), "上限=" + limit + " 結果=" + r1.Size + " ppi=" + r1.Ppi + " q=" + r1.Quality);
            Check("T-SIZ-03", "上限に近い画質を採用（上限の 60% 以上）", r1.Size >= limit * 0.6, (r1.Size * 100 / Math.Max(1, limit)) + "%");

            s.SizeLimitMB = 0.001;
            SizeResult r2 = SizeController.Optimize(src, wd, s, null);
            Check("T-SIZ-04", "届かない上限は目標未達、最低 ppi・品質より下げない", r2.Unachieved && r2.Changed && r2.Ppi >= 96 && r2.Quality >= 50, "ppi=" + r2.Ppi + " q=" + r2.Quality);
            s.UnachievedMode = "convertOnly";
            SizeResult r3 = SizeController.Optimize(src, wd, s, null);
            Check("T-SIZ-05", "目標未達時「圧縮せず出力」", r3.Unachieved && !r3.Changed && r3.Path == src, "");
            s.UnachievedMode = "best";

            s.SizeLimitMB = (s0 + 1000) / (1024.0 * 1024.0);
            SizeResult r4 = SizeController.Optimize(src, wd, s, null);
            Check("T-SIZ-06", "既に上限以下なら圧縮しない（ON）", !r4.Changed && r4.Path == src, "");
            s.SkipIfUnderLimit = false;
            SizeResult r5 = SizeController.Optimize(src, wd, s, null);
            Check("T-SIZ-06b", "上限以下でも初期設定で 1 回圧縮（OFF）", r5.Changed && r5.Ppi == 150 && r5.Quality == 85, "");

            // 画像以外が大きいブック
            string big = Path.Combine(root, "size_big.xlsx");
            File.Copy(src, big);
            using (ZipArchive z = ZipFile.Open(big, ZipArchiveMode.Update))
            {
                ZipArchiveEntry e = z.CreateEntry("xl/sharedStrings.xml", CompressionLevel.NoCompression);
                using (Stream st = e.Open()) { byte[] junk = new byte[3 * 1024 * 1024]; new Random(1).NextBytes(junk); st.Write(junk, 0, junk.Length); }
            }
            s.SkipIfUnderLimit = true; s.SizeLimitMB = 2;
            SizeResult r6 = SizeController.Optimize(big, wd, s, null);
            Check("T-SIZ-07", "画像以外だけで上限を超えるときは早期に目標未達と判定", r6.Unachieved && r6.Ppi == 150 && r6.Quality == 85, "ppi=" + r6.Ppi + " q=" + r6.Quality);
        }

        static byte[] OleBytes(bool vba) { return OleBytes(vba, false); }
        static byte[] OleBytes(bool vba, bool encrypted)
        {
            byte[] b = new byte[4096];
            byte[] sig = { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            Buffer.BlockCopy(sig, 0, b, 0, 8);
            if (vba)
            {
                byte[] name = Encoding.Unicode.GetBytes("_VBA_PROJECT_CUR");
                Buffer.BlockCopy(name, 0, b, 1000, name.Length);
            }
            if (encrypted)
            {
                byte[] name = Encoding.Unicode.GetBytes("EncryptedPackage");
                Buffer.BlockCopy(name, 0, b, 2000, name.Length);
            }
            return b;
        }

        static void TestScan()
        {
            string d = Path.Combine(root, "scan");
            Directory.CreateDirectory(Path.Combine(d, "sub", "deep"));
            Directory.CreateDirectory(Path.Combine(d, "old"));
            Directory.CreateDirectory(Path.Combine(d, "backup"));
            File.WriteAllBytes(Path.Combine(d, "a.xls"), OleBytes(false));
            File.WriteAllBytes(Path.Combine(d, "macro.xls"), OleBytes(true));
            File.WriteAllBytes(Path.Combine(d, "sub", "UPPER.XLS"), OleBytes(false));
            File.WriteAllBytes(Path.Combine(d, "sub", "deep", "deep.xls"), OleBytes(false));
            File.WriteAllBytes(Path.Combine(d, "old", "skip.xls"), OleBytes(false));
            File.WriteAllBytes(Path.Combine(d, "backup", "moved.xls"), OleBytes(false));
            File.WriteAllBytes(Path.Combine(d, "~$lock.xls"), OleBytes(false));
            File.WriteAllText(Path.Combine(d, "fake.xls"), "<html><table></table></html>");
            File.WriteAllBytes(Path.Combine(d, "empty.xls"), new byte[0]);
            File.WriteAllBytes(Path.Combine(d, "hidden.xls"), OleBytes(false));
            File.SetAttributes(Path.Combine(d, "hidden.xls"), FileAttributes.Hidden);
            File.WriteAllText(Path.Combine(d, "note.xls.bak"), "x");
            MakeXlsx(Path.Combine(d, "done.xlsx"), false);
            File.Copy(Path.Combine(d, "done.xlsx"), Path.Combine(d, "done_圧縮.xlsx"));
            File.WriteAllBytes(Path.Combine(d, "enc.xlsx"), OleBytes(false, true));
            File.Copy(Path.Combine(d, "done.xlsx"), Path.Combine(d, "done_圧縮(1).xlsx"));

            ProbeResult p1 = FormatProbe.Probe(Path.Combine(d, "a.xls"));
            ProbeResult p2 = FormatProbe.Probe(Path.Combine(d, "macro.xls"));
            ProbeResult p3 = FormatProbe.Probe(Path.Combine(d, "fake.xls"));
            ProbeResult p4 = FormatProbe.Probe(Path.Combine(d, "done.xlsx"));
            ProbeResult p5 = FormatProbe.Probe(Path.Combine(d, "empty.xls"));
            Check("T-MAC-PROBE1", "OLE 形式を判定", p1.Format == FileFormat.Ole && !p1.VbaStorage, "");
            Check("T-MAC-PROBE2", "OLE 内の _VBA_PROJECT_CUR を検出", p2.VbaStorage, "");
            Check("T-MAC-PROBE3", "HTML 形式の偽 xls を判定", p3.Format == FileFormat.Markup, p3.Format.ToString());
            Check("T-MAC-PROBE4", "ZIP（OOXML）を判定", p4.Format == FileFormat.Zip, "");
            Check("T-MAC-PROBE5", "0 バイトを判定", p5.Format == FileFormat.Empty, "");

            Settings s = new Settings(); s.TargetFolder = d; s.Recursive = true;
            List<FileItem> a = FileScanner.Scan(s, null, null);
            List<string> names = new List<string>();
            foreach (FileItem it in a) names.Add(Path.GetFileName(it.Path));
            string ns = string.Join(",", names.ToArray());
            Check("T-SCN-01", "再帰 ON で全階層の .xls を列挙", names.Contains("deep.xls") && names.Contains("a.xls"), ns);
            Check("T-SCN-03", "大文字拡張子 .XLS も対象", names.Contains("UPPER.XLS"), "");
            Check("T-SCN-04", "既定では .xlsx・.xls.bak を列挙しない", !names.Contains("done.xlsx") && !names.Contains("note.xls.bak"), "");
            Check("T-SCN-05", "~$ 一時ファイルを除外", !names.Contains("~$lock.xls"), "");
            Check("T-SCN-06", "隠しファイルを除外", !names.Contains("hidden.xls"), "");
            FileItem fk = null, mc = null;
            foreach (FileItem it in a) { if (it.Path.EndsWith("fake.xls")) fk = it; if (it.Path.EndsWith("macro.xls")) mc = it; }
            Check("T-SCN-14", "形式不一致は事前にチェックを外す、マクロありを表示", fk != null && !fk.Checked && fk.PreNote == "形式不一致" && mc != null && mc.PreNote == "マクロあり", "");

            s.Recursive = false;
            List<FileItem> b = FileScanner.Scan(s, null, null);
            bool flat = true; foreach (FileItem it in b) if (Path.GetDirectoryName(it.Path) != d) flat = false;
            Check("T-SCN-02", "再帰 OFF は直下のみ", flat && b.Count > 0, b.Count + " 件");

            s.Recursive = true; s.ExcludeFolders = "old; deep";
            List<FileItem> c = FileScanner.Scan(s, null, null);
            bool ex = true; foreach (FileItem it in c) if (it.Path.IndexOf("\\old\\") >= 0 || it.Path.IndexOf("\\deep\\") >= 0) ex = false;
            Check("T-SCN-07", "除外フォルダ名（; 区切り）", ex, "");

            s.ExcludeFolders = ""; s.MoveRoot = Path.Combine(d, "backup");
            List<FileItem> e2 = FileScanner.Scan(s, null, null);
            bool mv = true; foreach (FileItem it in e2) if (it.Path.EndsWith("moved.xls")) mv = false;
            Check("T-SCN-08", "対象フォルダ内の移動先はスキャンしない", mv, "");

            s.MoveRoot = ""; s.IncludeConverted = true;
            List<FileItem> f = FileScanner.Scan(s, null, null);
            FileItem done = null, sfx = null, enc = null;
            foreach (FileItem it in f)
            {
                string n = Path.GetFileName(it.Path);
                if (n == "done.xlsx") done = it; if (n == "done_圧縮.xlsx") sfx = it; if (n == "enc.xlsx") enc = it;
            }
            Check("T-CMP-02", "「変換済みも対象」ON で .xlsx を「圧縮のみ」で列挙", done != null && done.Kind == ItemKind.CompressOnly && done.Checked, "");
            bool numbered = false; foreach (FileItem it in f) if (Path.GetFileName(it.Path) == "done_圧縮(1).xlsx") numbered = true;
            Check("T-CMP-06", "接尾辞「_圧縮」「_圧縮(1)」付きは対象外", sfx == null && !numbered, "");
            Check("T-CMP-ENC", "暗号化 xlsx（実体が OLE）はチェックを外す", enc != null && !enc.Checked && enc.PreNote == "暗号化", enc == null ? "" : enc.PreNote);
        }

        static Settings JobSettings(string d)
        {
            Settings s = new Settings();
            s.TargetFolder = d; s.Recursive = true; s.IncludeConverted = true; s.VerifyLayout = false;
            return s;
        }
        static List<FileItem> RunJob(Settings s, bool dry, out JobRunner jr)
        {
            List<FileItem> items = FileScanner.Scan(s, null, null);
            foreach (FileItem it in items) it.Checked = it.Kind == ItemKind.CompressOnly && it.Checked;
            jr = new JobRunner(s); jr.DryRun = dry;
            jr.Run(items);
            return items;
        }
        static FileItem Find(List<FileItem> l, string n)
        {
            foreach (FileItem it in l) if (string.Equals(Path.GetFileName(it.Path), n, StringComparison.OrdinalIgnoreCase)) return it;
            return null;
        }
        static string St(FileItem it) { return it == null ? "(なし)" : it.Status + " " + it.ErrorCode + " " + it.Message; }

        static string Prepare(string name, out string master)
        {
            string d = Path.Combine(root, name);
            Directory.CreateDirectory(Path.Combine(d, "sub"));
            MakeXlsx(Path.Combine(d, "写真 帳.xlsx"), false);
            MakeXlsx(Path.Combine(d, "sub", "macro.xlsm"), true);
            // 画像なし（小さくならない）
            string tiny = Path.Combine(d, "plain.xlsx");
            using (FileStream fs = new FileStream(tiny, FileMode.CreateNew))
            using (ZipArchive z = new ZipArchive(fs, ZipArchiveMode.Create))
            {
                ZipArchiveEntry e = z.CreateEntry("[Content_Types].xml");
                using (StreamWriter w = new StreamWriter(e.Open())) w.Write("<?xml version=\"1.0\"?><Types xmlns=\"http://schemas.openxmlformats.org/package/2006/content-types\"/>");
            }
            master = Hash(Path.Combine(d, "写真 帳.xlsx"));
            return d;
        }

        static void TestCompressJob()
        {
            string master;
            JobRunner jr;

            // 残す
            string d = Prepare("job_keep", out master);
            DateTime lw = File.GetLastWriteTimeUtc(Path.Combine(d, "写真 帳.xlsx"));
            List<FileItem> items = RunJob(JobSettings(d), false, out jr);
            FileItem p = Find(items, "写真 帳.xlsx");
            Check("T-CMP-08", "残す: 「_圧縮」付きで出力し元は不変", p.Status == "完了" && File.Exists(Path.Combine(d, "写真 帳_圧縮.xlsx")) && Hash(Path.Combine(d, "写真 帳.xlsx")) == master && File.GetLastWriteTimeUtc(Path.Combine(d, "写真 帳.xlsx")) == lw, St(p));
            Check("T-CMP-08b", "出力サイズが元より小さい", p.FinalSize > 0 && p.FinalSize < p.Size && new FileInfo(Path.Combine(d, "写真 帳_圧縮.xlsx")).Length == p.FinalSize, p.Size + " → " + p.FinalSize);
            FileItem m = Find(items, "macro.xlsm");
            Check("T-CMP-04b", "xlsm は .xlsm のまま出力しマクロを保持", m.Status == "完了" && File.Exists(Path.Combine(d, "sub", "macro_圧縮.xlsm")) && ImageOptimizer.HasVbaProject(Path.Combine(d, "sub", "macro_圧縮.xlsm")) && m.Macro == "VBA", St(m));
            FileItem pl = Find(items, "plain.xlsx");
            Check("T-CMP-07", "小さくならないファイルは「変更なし」で出力しない", pl.Status == "変更なし" && !File.Exists(Path.Combine(d, "plain_圧縮.xlsx")), St(pl));
            Check("T-CNV-15", "一時ファイル .~tmp が残らない", Directory.GetFiles(d, "*.~tmp", SearchOption.AllDirectories).Length == 0, "");
            Check("T-ERR-13", "作業フォルダが残らない", !Directory.Exists(Path.Combine(AppPaths.TempRoot, jr.JobId)), "");
            byte[] rb = File.ReadAllBytes(jr.ReportPath);
            string[] lines = File.ReadAllLines(jr.ReportPath, Encoding.UTF8);
            Check("T-LOG-01", "レポート CSV（UTF-8 BOM 付き、全件 1 行ずつ）", rb.Length > 3 && rb[0] == 0xEF && rb[1] == 0xBB && rb[2] == 0xBF && lines.Length == 4, "行数=" + lines.Length);
            Check("T-LOG-03", "ログファイルが出力される", File.Exists(jr.Log.FilePath) && new FileInfo(jr.Log.FilePath).Length > 0, "");

            // 2 回目: 同名出力 → 連番、スキップ
            RunJob(JobSettings(d), false, out jr);
            Check("T-CNV-12", "同名出力があるとき連番を付ける", File.Exists(Path.Combine(d, "写真 帳_圧縮(1).xlsx")), "");
            Settings sk = JobSettings(d); sk.Conflict = "skip";
            items = RunJob(sk, false, out jr);
            Check("T-CNV-10", "同名出力があるときスキップ", Find(items, "写真 帳.xlsx").Status == "スキップ" && !File.Exists(Path.Combine(d, "写真 帳_圧縮(2).xlsx")), St(Find(items, "写真 帳.xlsx")));

            // ドライラン
            string dd = Prepare("job_dry", out master);
            Settings ds = JobSettings(dd); ds.OriginalMode = "delete"; ds.DeleteToRecycle = false;
            int before = Directory.GetFiles(dd, "*", SearchOption.AllDirectories).Length;
            items = RunJob(ds, true, out jr);
            p = Find(items, "写真 帳.xlsx");
            Check("T-GUI-13", "ドライラン: ファイルを変更せず予測サイズを出す", p.Status == "ドライラン" && p.FinalSize < p.Size && Hash(Path.Combine(dd, "写真 帳.xlsx")) == master && Directory.GetFiles(dd, "*", SearchOption.AllDirectories).Length == before, St(p));

            // 使用中
            string dl = Prepare("job_lock", out master);
            using (new FileStream(Path.Combine(dl, "写真 帳.xlsx"), FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                items = RunJob(JobSettings(dl), false, out jr);
            }
            FileItem lk = Find(items, "写真 帳.xlsx"), ok = Find(items, "macro.xlsm");
            Check("T-ERR-02", "使用中のファイルはスキップして次へ", lk.Status == "スキップ" && lk.ErrorCode == "E101" && ok.Status == "完了", St(lk) + " / " + St(ok));

            // 削除: 同名で置き換え
            string dx = Prepare("job_del", out master);
            Settings xs = JobSettings(dx); xs.OriginalMode = "delete"; xs.DeleteToRecycle = false;
            items = RunJob(xs, false, out jr);
            p = Find(items, "写真 帳.xlsx");
            string here = Path.Combine(dx, "写真 帳.xlsx");
            Check("T-CMP-09", "削除: 元と同じ名前で圧縮版に置き換え", p.Status == "完了" && File.Exists(here) && Hash(here) != master && new FileInfo(here).Length == p.FinalSize && !File.Exists(Path.Combine(dx, "写真 帳_圧縮.xlsx")), St(p));
            Check("T-CMP-09b", "置き換え後のファイルが ZIP として正常", Valid(here), "");
            Check("T-ORG-08", "「変更なし」のファイルは削除しない", File.Exists(Path.Combine(dx, "plain.xlsx")), "");

            // ネットワーク判定はローカルでは false
            Check("T-ORG-03a", "ローカルパスはネットワーク扱いしない", !Util.IsNetworkPath(here), "");
        }

        static bool Valid(string p) { try { ImageOptimizer.CheckZip(p); return true; } catch { return false; } }

        static void TestRestore()
        {
            string master;
            JobRunner jr;

            // 圧縮のみ・移動 → 復元
            string d = Prepare("job_move", out master);
            string macroHash = Hash(Path.Combine(d, "sub", "macro.xlsm"));
            DateTime lw = File.GetLastWriteTimeUtc(Path.Combine(d, "写真 帳.xlsx"));
            Settings s = JobSettings(d); s.OriginalMode = "move"; s.MoveRoot = Path.Combine(root, "backup1");
            List<FileItem> items = RunJob(s, false, out jr);
            string broot = jr.Man.BackupRoot;
            string here = Path.Combine(d, "写真 帳.xlsx"), there = Path.Combine(broot, "写真 帳.xlsx");
            FileItem p = Find(items, "写真 帳.xlsx");
            Check("T-CMP-10", "移動: 元は移動先へ、同名で圧縮版に置き換え", p.Status == "完了" && File.Exists(there) && Hash(there) == master && File.Exists(here) && Hash(here) != master, St(p));
            Check("T-ORG-05", "移動先でフォルダ構造を保つ", File.Exists(Path.Combine(broot, "sub", "macro.xlsm")) && Hash(Path.Combine(broot, "sub", "macro.xlsm")) == macroHash, broot);
            Check("T-ORG-05b", "移動しても元の更新日時を保持", File.GetLastWriteTimeUtc(there) == lw, "");
            Check("T-RST-00", "マニフェストを履歴と移動先の両方に保存", File.Exists(jr.Man.FilePath) && File.Exists(Path.Combine(broot, "_manifest.json")), "");
            Check("T-ORG-08b", "「変更なし」のファイルは移動しない", File.Exists(Path.Combine(d, "plain.xlsx")) && !File.Exists(Path.Combine(broot, "plain.xlsx")), "");

            Manifest man = Manifest.Load(jr.Man.FilePath);
            ManEntry sub = null;
            foreach (ManEntry e in man.Entries) if (e.Original.EndsWith("macro.xlsm")) sub = e;
            RestoreResult r1 = RestoreService.Restore(man, new ManEntry[] { sub }, true, false, null);
            Check("T-RST-02", "選択復元: 1 件だけ戻る", r1.Restored == 1 && Hash(Path.Combine(d, "sub", "macro.xlsm")) == macroHash && Hash(here) != master, string.Join(" | ", r1.Messages.ToArray()));

            man = Manifest.Load(jr.Man.FilePath);
            RestoreResult r2 = RestoreService.Restore(man, man.Entries, true, false, null);
            Check("T-CMP-10b", "全体復元: 元ファイルが同じ名前・同じ内容で戻る", r2.Restored == 1 && Hash(here) == master && !File.Exists(there), string.Join(" | ", r2.Messages.ToArray()));
            Check("T-RST-01b", "復元後の更新日時が元と同じ", File.GetLastWriteTimeUtc(here) == lw, "");
            man = Manifest.Load(jr.Man.FilePath);
            RestoreResult r3 = RestoreService.Restore(man, man.Entries, true, false, null);
            Check("T-RST-09", "二重復元しても何も変わらない", r3.Restored == 0 && Hash(here) == master, "");
            Check("T-RST-CLEAN", "全件復元後、移動先の空フォルダを片付ける", !Directory.Exists(Path.Combine(broot, "sub")), "");

            // 圧縮版を編集してから復元 → 「_復元」
            string d2 = Prepare("job_move_edit", out master);
            Settings s2 = JobSettings(d2); s2.OriginalMode = "move"; s2.MoveRoot = Path.Combine(root, "backup2");
            RunJob(s2, false, out jr);
            string h2 = Path.Combine(d2, "写真 帳.xlsx");
            File.AppendAllText(h2, "edited");
            string editedHash = Hash(h2);
            man = Manifest.Load(jr.Man.FilePath);
            RestoreResult r4 = RestoreService.Restore(man, man.Entries, true, false, null);
            Check("T-CMP-11", "圧縮版が編集済みなら残し、元は「_復元」で戻す", Hash(h2) == editedHash && File.Exists(Path.Combine(d2, "写真 帳_復元.xlsx")) && Hash(Path.Combine(d2, "写真 帳_復元.xlsx")) == master, string.Join(" | ", r4.Messages.ToArray()));

            // 移動先で変更されたファイル → 戻さない
            string d3 = Prepare("job_move_mod", out master);
            Settings s3 = JobSettings(d3); s3.OriginalMode = "move"; s3.MoveRoot = Path.Combine(root, "backup3");
            RunJob(s3, false, out jr);
            File.AppendAllText(Path.Combine(jr.Man.BackupRoot, "写真 帳.xlsx"), "x");
            string compressedHash = Hash(Path.Combine(d3, "写真 帳.xlsx"));
            man = Manifest.Load(jr.Man.FilePath);
            RestoreResult r5 = RestoreService.Restore(man, man.Entries, true, false, null);
            Check("T-RST-08", "移動先で変更されたファイルは復元しない", r5.Failed >= 1 && Hash(Path.Combine(d3, "写真 帳.xlsx")) == compressedHash, string.Join(" | ", r5.Messages.ToArray()));

            // 変換ジョブ相当（kind=convert）の復元: 元の場所に同名があればスキップ、生成ファイル削除
            string d4 = Path.Combine(root, "conv_restore"); Directory.CreateDirectory(d4);
            string orig = Path.Combine(d4, "a.xls"), outp = Path.Combine(d4, "a.xlsx"), bk = Path.Combine(root, "backup4", "a.xls");
            string orig2 = Path.Combine(d4, "b.xls"), outp2 = Path.Combine(d4, "b.xlsx"), bk2 = Path.Combine(root, "backup4", "b.xls");
            File.WriteAllBytes(orig, OleBytes(false)); File.WriteAllText(outp, "generated");
            File.WriteAllBytes(orig2, OleBytes(true)); File.WriteAllText(outp2, "generated2");
            Manifest mm = new Manifest(); mm.JobId = "test"; mm.Mode = "move"; mm.BackupRoot = Path.Combine(root, "backup4");
            mm.FilePath = Path.Combine(AppPaths.HistoryDir, "29991231_000000_test.manifest.json");
            foreach (string[] t in new string[][] { new string[] { orig, outp, bk }, new string[] { orig2, outp2, bk2 } })
            {
                ManEntry e = new ManEntry();
                e.Original = t[0]; e.Output = t[1]; e.Backup = t[2];
                e.OriginalSha256 = Hash(t[0]); e.OutputSha256 = Hash(t[1]);
                e.Modified = File.GetLastWriteTimeUtc(t[0]).ToString("o");
                OriginalHandler.Move(t[0], t[2], e.OriginalSha256);
                e.State = "moved"; mm.Entries.Add(e);
            }
            mm.Save();
            File.WriteAllText(orig2, "someone else");
            RestoreResult r6 = RestoreService.Restore(mm, mm.Entries, true, false, null);
            Check("T-RST-01", "変換ジョブの復元: 元が戻り生成ファイルは削除", File.Exists(orig) && FormatProbe.Probe(orig).Format == FileFormat.Ole && !File.Exists(outp), string.Join(" | ", r6.Messages.ToArray()));
            Check("T-RST-06", "元の場所に同名があれば上書きせずスキップ", File.ReadAllText(orig2) == "someone else" && File.Exists(bk2), "");

            // 壊れたマニフェスト
            string bad = Path.Combine(AppPaths.HistoryDir, "00000000_000000_bad.manifest.json");
            File.WriteAllText(bad, "{ broken");
            List<string> errs = new List<string>();
            List<Manifest> all = Manifest.LoadAll(errs);
            Check("T-RST-12", "壊れたマニフェストがあっても他の履歴を読める", errs.Count == 1 && all.Count >= 3, "エラー=" + errs.Count + " 読込=" + all.Count);
            File.Delete(bad);

            // 移動先のマニフェストから読み込み
            Manifest fromBackup = Manifest.Load(Path.Combine(root, "backup4", "_manifest.json"));
            Check("T-RST-11", "移動先の _manifest.json から読み込める", fromBackup.Entries.Count == 2, "");

            // 中断された移動の修復
            Manifest rm = new Manifest(); rm.FilePath = Path.Combine(root, "repair.manifest.json");
            ManEntry mv1 = new ManEntry(); mv1.Original = Path.Combine(d4, "gone.xls"); mv1.Backup = bk2; mv1.State = "moving";
            ManEntry mv2 = new ManEntry(); mv2.Original = orig; mv2.Backup = Path.Combine(root, "nothing.xls"); mv2.State = "moving";
            rm.Entries.Add(mv1); rm.Entries.Add(mv2);
            int n = rm.Repair();
            Check("T-ORG-09", "「移動中」で止まったエントリを実体の場所に合わせて修復", n == 2 && mv1.State == "moved" && mv2.State == "kept", mv1.State + "/" + mv2.State);

            // 別ドライブ相当: 同一ボリューム判定
            Check("T-ORG-06a", "同一ボリューム判定", Util.SameVolume(@"C:\a\b.xls", @"C:\x\y.xls") && !Util.SameVolume(@"C:\a.xls", @"D:\a.xls"), "");
        }

        static void TestUi()
        {
            string d = Path.Combine(root, "ui");
            Directory.CreateDirectory(Path.Combine(d, "sub"));
            File.WriteAllBytes(Path.Combine(d, "a.xls"), OleBytes(false));
            File.WriteAllBytes(Path.Combine(d, "sub", "b.xls"), OleBytes(true));
            MakeXlsx(Path.Combine(d, "sub", "c.xlsx"), false);
            bool excel = ExcelSession.IsInstalled();
            using (MainForm f = new MainForm())
            {
                Check("T-GUI-01b", "フォルダ未選択では「実行」を押せない", !f.TestRunEnabled(), "");
                f.TestSetFolder(d);
                Check("T-GUI-20", "フォルダを選ぶと自動で検索し一覧に表示", f.TestListCount() == 2, "件数=" + f.TestListCount() + " " + f.TestHint());
                Check("T-GUI-21", "検索後は「実行」を押せる（スキャンボタン不要）", f.TestRunEnabled(), "");
                Check("T-GUI-22", "検索結果の件数を画面に表示", f.TestHint().IndexOf("2 件見つかりました") >= 0, f.TestHint());
                if (!excel) Check("T-ERR-01b", "Excel がない環境では .xls のチェックを外す", f.TestCheckedCount() == 0, "チェック=" + f.TestCheckedCount());
                f.TestSetIncludeConverted(true);
                f.TestSetFolder(d);
                Check("T-GUI-23", "「変換済みも対象」を ON にすると .xlsx も一覧に出る", f.TestListCount() == 3, "件数=" + f.TestListCount());
                if (!excel) Check("T-ERR-01c", "Excel がなくても変換済みファイルは処理対象", f.TestCheckedCount() == 1, "チェック=" + f.TestCheckedCount());
                f.TestSetRecursive(false);
                f.TestSetFolder(d);
                Check("T-GUI-24", "サブフォルダ OFF で直下のみ", f.TestListCount() == 1, "件数=" + f.TestListCount());
                f.TestSetFolder(Path.Combine(d, "nothing_here"));
                Check("T-GUI-25", "存在しないフォルダでは一覧を空にし「実行」を押せない", f.TestListCount() == 0 && !f.TestRunEnabled() && f.TestHint().IndexOf("見つかりません") >= 0, f.TestHint());
                string empty = Path.Combine(root, "ui_empty"); Directory.CreateDirectory(empty);
                f.TestSetFolder(empty);
                Check("T-GUI-26", "対象ファイルがないときはそう表示", f.TestListCount() == 0 && f.TestHint().IndexOf("対象ファイルが見つかりません") >= 0, f.TestHint());
            }
        }

        static void TestSettings()
        {
            string sf = Path.Combine(root, "settings_test.json");
            Settings w = new Settings(); w.Ppi = 123; w.CompressSuffix = "_小"; w.KeepOnLayoutDiff = false; w.Save(sf);
            Settings r = Settings.Load(sf);
            Check("T-GUI-02", "設定の保存と読込（日本語を含む）", r.Ppi == 123 && r.CompressSuffix == "_小" && !r.KeepOnLayoutDiff, "");
            File.WriteAllText(sf, "{ broken");
            Settings r2 = Settings.Load(sf);
            Check("T-LOG-05", "壊れた設定は既定値で起動し .bak に退避", r2.Ppi == 150 && File.Exists(sf + ".bak"), "");
            Settings d = new Settings();
            Check("T-CMP-01", "「変換済みも対象」の初期値は OFF", !d.IncludeConverted, "");
            Check("T-CMP-12a", "「差異があるときは削除・移動しない」の初期値は ON", d.KeepOnLayoutDiff, "");
            Check("T-GUI-01", "元ファイルの初期値は「残す」、既定値は入力チェックを通る", d.OriginalMode == "keep" && d.Validate() == null, d.Validate());

            Settings v;
            v = new Settings(); v.SizeTargetEnabled = true; v.SizeLimitMB = 0;
            Check("T-SIZ-10", "上限 0 MB は不可", v.Validate() != null, v.Validate());
            v = new Settings(); v.SizeTargetEnabled = false; v.SizeLimitMB = 0;
            Check("T-SIZ-10b", "サイズ目標が無効なら上限は問わない", v.Validate() == null, v.Validate());
            v = new Settings(); v.MinQuality = 90; v.JpegQuality = 80;
            Check("T-GUI-04a", "最低品質 > JPEG 品質 は不可", v.Validate() != null, "");
            v = new Settings(); v.MinPpi = 200; v.Ppi = 150;
            Check("T-GUI-04b", "最低 ppi > 解像度 は不可", v.Validate() != null, "");
            v = new Settings(); v.MaxLongSide = 100;
            Check("T-GUI-04c", "長辺上限 100px は不可（0 または 320〜10000）", v.Validate() != null, "");
            v = new Settings(); v.OriginalMode = "move"; v.MoveRoot = "";
            Check("T-GUI-05", "移動で移動先未入力は不可", v.Validate() != null, "");
            v = new Settings(); v.CompressSuffix = "a/b";
            Check("T-GUI-04d", "ファイル名に使えない接尾辞は不可", v.Validate() != null, "");

            Check("T-SIZ-09", "サイズ表示（MB は 1024×1024）", Util.FormatSize(1024 * 1024) == "1.0MB" && Util.FormatSize(1536) == "2KB", Util.FormatSize(1024 * 1024));
            Check("T-LOG-02", "CSV の値に改行・カンマ・引用符があっても 1 行に収まる", Util.Csv("a,b\"c\r\nd") == "\"a,b\"\"c  d\"", Util.Csv("a,b\"c\r\nd"));
            string up = Path.Combine(root, "uniq"); Directory.CreateDirectory(up);
            File.WriteAllText(Path.Combine(up, "x.xlsx"), ""); File.WriteAllText(Path.Combine(up, "x(1).xlsx"), "");
            Check("T-CNV-12b", "連番は空いている番号を使う", Path.GetFileName(Util.UniquePath(up, "x", ".xlsx")) == "x(2).xlsx", "");
            Check("T-RST-REL", "移動先の相対パス計算", Util.RelativeTo(@"C:\data", @"C:\data\A\b.xls") == @"A\b.xls" && Util.RelativeTo(@"C:\data\", @"C:\data\b.xls") == "b.xls", "");
        }
    }
}
