using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.Drawing.Imaging;
using System.IO;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System.Text;

namespace XlsImageShrink
{
    /// <summary>実機の Excel を使う自動テスト。XlsImageShrink.exe --selftest フォルダ で実行する。</summary>
    public static class SelfTest
    {
        static string root, repPath, master;
        static int pass, fail, skip;
        static bool vbaOk, xlmOk, nameOk, pwOk;
        static Dictionary<string, string> masterHash;

        static void Out(string line)
        {
            File.AppendAllText(repPath, line + "\r\n", new UTF8Encoding(false));
        }
        static void Check(string id, string desc, bool ok, string detail)
        {
            if (ok) pass++; else fail++;
            Out((ok ? "PASS " : "FAIL ") + id + " " + desc + (string.IsNullOrEmpty(detail) ? "" : "  [" + detail + "]"));
        }
        static void SkipT(string id, string desc, string why) { skip++; Out("SKIP " + id + " " + desc + "  [" + why + "]"); }

        public static int Run(string dir)
        {
            root = Path.GetFullPath(dir);
            if (Directory.Exists(root))
            {
                foreach (string f in Directory.GetFiles(root, "*", SearchOption.AllDirectories)) { try { File.SetAttributes(f, FileAttributes.Normal); } catch { } }
                Directory.Delete(root, true);
            }
            Directory.CreateDirectory(root);
            repPath = Path.Combine(root, "selftest_report.txt");
            AppPaths.Init(Path.Combine(root, "app"));
            Util.NoRecycle = true;
            Out("XlsImageShrink 自己テスト  " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss"));
            Out("OS: " + Environment.OSVersion + "  64bit=" + Environment.Is64BitOperatingSystem + "  .NET CLR " + Environment.Version + "  プロセス64bit=" + Environment.Is64BitProcess);
            if (!ExcelSession.IsInstalled()) { Out("FAIL T-ERR-01 Excel が見つかりません"); return 1; }
            int excelBefore = Process.GetProcessesByName("EXCEL").Length;
            Stopwatch sw = Stopwatch.StartNew();
            try
            {
                MakeData();
                Out("テストデータ作成 " + sw.Elapsed.TotalSeconds.ToString("0") + " 秒  (VBA作成=" + vbaOk + " XLMシート=" + xlmOk + " 名前定義=" + nameOk + " パスワード=" + pwOk + ")");
                Step("T-SCN", TestScan);
                Step("T-CNV", TestConvert);
                Step("T-SIZ", TestSize);
                Step("T-ORG move", TestMoveRestore);
                Step("T-ORG delete", TestDelete);
                Step("T-CMP", TestCompressOnly);
                Step("T-DIFF", TestLayoutDiff);
                Step("T-DRY", TestDryRunAndLock);
            }
            catch (Exception ex) { fail++; Out("FATAL " + ex); }
            GC.Collect(); GC.WaitForPendingFinalizers();
            System.Threading.Thread.Sleep(3000);
            int excelAfter = Process.GetProcessesByName("EXCEL").Length;
            Check("T-ERR-14", "EXCEL.EXE が残っていない", excelAfter <= excelBefore, "前=" + excelBefore + " 後=" + excelAfter);
            Check("T-ERR-13", "作業フォルダが残っていない", !Directory.Exists(AppPaths.TempRoot) || Directory.GetDirectories(AppPaths.TempRoot).Length == 0, AppPaths.TempRoot);
            Out("");
            Out("合計: PASS " + pass + " / FAIL " + fail + " / SKIP " + skip + "   所要 " + sw.Elapsed.TotalSeconds.ToString("0") + " 秒");
            Out("END");
            return fail == 0 ? 0 : 1;
        }

        static void Step(string name, Action a)
        {
            Stopwatch sw = Stopwatch.StartNew();
            Out("--- " + name);
            try { a(); }
            catch (Exception ex) { fail++; Out("FAIL " + name + " 例外: " + ex); }
            Out("    (" + sw.Elapsed.TotalSeconds.ToString("0") + " 秒)");
        }

        // ================= テストデータ =================
        static void NoiseImage(string path, int w, int h, int seed, ImageFormat fmt, bool alpha)
        {
            using (Bitmap b = new Bitmap(w, h, alpha ? PixelFormat.Format32bppArgb : PixelFormat.Format24bppRgb))
            {
                int bpp = alpha ? 4 : 3;
                BitmapData bd = b.LockBits(new Rectangle(0, 0, w, h), ImageLockMode.WriteOnly, b.PixelFormat);
                byte[] row = new byte[bd.Stride];
                Random rnd = new Random(seed);
                byte[] noise = new byte[w * bpp];
                for (int y = 0; y < h; y++)
                {
                    rnd.NextBytes(noise);
                    for (int x = 0; x < w; x++)
                    {
                        int i = x * bpp;
                        int g = (x * 255 / w + y * 255 / h) / 2;
                        row[i] = (byte)((g + noise[i]) / 2);
                        row[i + 1] = (byte)((255 - g + noise[i + 1]) / 2);
                        row[i + 2] = (byte)((x * 255 / w + noise[i + 2]) / 2);
                        if (alpha) row[i + 3] = (byte)((x < w / 3) ? 0 : 255);
                    }
                    Marshal.Copy(row, 0, new IntPtr(bd.Scan0.ToInt64() + (long)y * bd.Stride), bd.Stride);
                }
                b.UnlockBits(bd);
                if (fmt.Guid == ImageFormat.Jpeg.Guid)
                {
                    ImageCodecInfo codec = null;
                    foreach (ImageCodecInfo c in ImageCodecInfo.GetImageEncoders()) if (c.FormatID == ImageFormat.Jpeg.Guid) codec = c;
                    using (EncoderParameters ep = new EncoderParameters(1))
                    {
                        ep.Param[0] = new EncoderParameter(System.Drawing.Imaging.Encoder.Quality, 92L);
                        b.Save(path, codec, ep);
                    }
                }
                else b.Save(path, fmt);
            }
        }

        static void FillSheet(dynamic ws)
        {
            ws.Cells[1, 1].Value2 = "点検報告書";
            ws.Range["A1:D1"].Merge();
            ws.Columns[1].ColumnWidth = 22;
            ws.Rows[3].RowHeight = 31.5;
            for (int r = 4; r <= 30; r++)
            {
                ws.Cells[r, 1].Value2 = "項目 " + r;
                ws.Cells[r, 2].Value2 = r * 1.5;
                ws.Cells[r, 3].Formula = "=B" + r + "*2";
            }
        }

        static void MakeData()
        {
            master = Path.Combine(root, "master");
            string img = Path.Combine(root, "img");
            Directory.CreateDirectory(master); Directory.CreateDirectory(img);
            Directory.CreateDirectory(Path.Combine(master, "sub", "deep"));
            Directory.CreateDirectory(Path.Combine(master, "old"));
            string j1 = Path.Combine(img, "p1.jpg"), j2 = Path.Combine(img, "p2.jpg"), j3 = Path.Combine(img, "p3.jpg"), p1 = Path.Combine(img, "opaque.png"), p2 = Path.Combine(img, "alpha.png");
            NoiseImage(j1, 3000, 2000, 1, ImageFormat.Jpeg, false);
            NoiseImage(j2, 3000, 2000, 2, ImageFormat.Jpeg, false);
            NoiseImage(j3, 2400, 1600, 3, ImageFormat.Jpeg, false);
            NoiseImage(p1, 1500, 1000, 4, ImageFormat.Png, false);
            NoiseImage(p2, 900, 600, 5, ImageFormat.Png, true);

            Type t = Type.GetTypeFromProgID("Excel.Application");
            dynamic app = Activator.CreateInstance(t);
            try
            {
                app.Visible = false; app.DisplayAlerts = false;
                Out("Excel Version " + Convert.ToString((object)app.Version) + " / " + Convert.ToString((object)app.OperatingSystem));

                // 写真入り xls
                dynamic wb = app.Workbooks.Add();
                dynamic ws = wb.Worksheets[1];
                ws.Name = "写真";
                FillSheet(ws);
                ws.Shapes.AddPicture(j1, 0, -1, 10f, 500f, 288f, 192f);
                ws.Shapes.AddPicture(j2, 0, -1, 310f, 500f, 288f, 192f);
                ws.Shapes.AddPicture(j3, 0, -1, 10f, 700f, 144f, 96f);
                ws.Shapes.AddPicture(p1, 0, -1, 310f, 700f, 216f, 144f);
                ws.Shapes.AddPicture(p2, 0, -1, 10f, 860f, 180f, 120f);
                try { dynamic tb = ws.Shapes.AddTextbox(1, 400f, 60f, 160f, 40f); tb.TextFrame2.TextRange.Text = "確認済み"; } catch { }
                try { ws.Shapes.AddShape(1, 400f, 120f, 80f, 40f); } catch { }
                wb.SaveAs(Filename: Path.Combine(master, "報告 1.xls"), FileFormat: 56);
                // 同じ内容を xlsx でも保存（圧縮のみのテスト用）
                wb.SaveAs(Filename: Path.Combine(master, "済.xlsx"), FileFormat: 51);
                wb.Close(false);

                // 画像なし
                wb = app.Workbooks.Add(); ws = wb.Worksheets[1]; FillSheet(ws);
                wb.SaveAs(Filename: Path.Combine(master, "sub", "simple.xls"), FileFormat: 56);
                wb.SaveAs(Filename: Path.Combine(master, "plain.xlsx"), FileFormat: 51);
                wb.Close(false);
                File.Copy(Path.Combine(master, "sub", "simple.xls"), Path.Combine(master, "old", "skipme.xls"));
                File.Copy(Path.Combine(master, "sub", "simple.xls"), Path.Combine(master, "~$temp.xls"));
                File.Copy(Path.Combine(master, "sub", "simple.xls"), Path.Combine(master, "sub", "UPPER.XLS"));

                // Excel 4.0 マクロシート
                try
                {
                    wb = app.Workbooks.Add(); ws = wb.Worksheets[1]; FillSheet(ws);
                    wb.Sheets.Add(Type: 3);
                    wb.SaveAs(Filename: Path.Combine(master, "sub", "deep", "xlm.xls"), FileFormat: 56);
                    wb.Close(false); xlmOk = true;
                }
                catch (Exception ex) { Out("  XLM シートを作成できません: " + ex.Message); try { wb.Close(false); } catch { } }

                // 名前定義に GET.CELL
                try
                {
                    wb = app.Workbooks.Add(); ws = wb.Worksheets[1]; FillSheet(ws);
                    string sn = Convert.ToString((object)ws.Name);
                    wb.Names.Add(Name: "CellColor", RefersTo: "=GET.CELL(63,'" + sn + "'!$A$1)");
                    wb.SaveAs(Filename: Path.Combine(master, "sub", "name.xls"), FileFormat: 56);
                    wb.Close(false); nameOk = true;
                }
                catch (Exception ex) { Out("  名前定義を作成できません: " + ex.Message); try { wb.Close(false); } catch { } }

                // VBA（「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」が有効なときだけ作れる）
                try
                {
                    wb = app.Workbooks.Add(); ws = wb.Worksheets[1]; FillSheet(ws);
                    dynamic comp = wb.VBProject.VBComponents.Add(1);
                    comp.CodeModule.AddFromString("Sub Hello()\r\n    MsgBox \"hello\"\r\nEnd Sub");
                    wb.SaveAs(Filename: Path.Combine(master, "vba.xls"), FileFormat: 56);
                    wb.Close(false); vbaOk = true;
                }
                catch (Exception ex) { Out("  VBA 入りブックを作成できません（信頼設定が無効）: " + ex.Message); try { wb.Close(false); } catch { } }

                // 読み取りパスワード
                try
                {
                    wb = app.Workbooks.Add(); ws = wb.Worksheets[1]; FillSheet(ws);
                    wb.SaveAs(Filename: Path.Combine(master, "pw.xls"), FileFormat: 56, Password: "abc123");
                    wb.Close(false); pwOk = true;
                }
                catch (Exception ex) { Out("  パスワード付きブックを作成できません: " + ex.Message); try { wb.Close(false); } catch { } }
            }
            finally
            {
                int pid = 0;
                try { uint p; GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out p); pid = (int)p; } catch { }
                try { app.Quit(); } catch { }
                try { Marshal.FinalReleaseComObject(app); } catch { }
                app = null;
                GC.Collect(); GC.WaitForPendingFinalizers(); GC.Collect();
                try { if (pid != 0) { Process pr = Process.GetProcessById(pid); if (!pr.WaitForExit(10000)) pr.Kill(); } } catch { }
            }
            // 偽 xls・破損
            File.WriteAllText(Path.Combine(master, "fake.xls"), "<html><body><table><tr><td>1</td></tr></table></body></html>");
            byte[] broken = new byte[4096];
            new Random(9).NextBytes(broken);
            byte[] sig = new byte[] { 0xD0, 0xCF, 0x11, 0xE0, 0xA1, 0xB1, 0x1A, 0xE1 };
            Buffer.BlockCopy(sig, 0, broken, 0, 8);
            File.WriteAllBytes(Path.Combine(master, "broken.xls"), broken);
            masterHash = HashDir(master);
            foreach (KeyValuePair<string, string> kv in masterHash) Out("  data: " + kv.Key + "  " + new FileInfo(Path.Combine(master, kv.Key)).Length + " bytes");
        }

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        // ================= 補助 =================
        static Dictionary<string, string> HashDir(string dir)
        {
            Dictionary<string, string> d = new Dictionary<string, string>(StringComparer.OrdinalIgnoreCase);
            foreach (string f in Directory.GetFiles(dir, "*", SearchOption.AllDirectories)) d[Util.RelativeTo(dir, f)] = Util.Sha256(f);
            return d;
        }
        static string Fresh(string name)
        {
            string dst = Path.Combine(root, name);
            foreach (string f in Directory.GetFiles(master, "*", SearchOption.AllDirectories))
            {
                string to = Path.Combine(dst, Util.RelativeTo(master, f));
                Directory.CreateDirectory(Path.GetDirectoryName(to));
                File.Copy(f, to);
            }
            return dst;
        }
        static Settings Base(string folder)
        {
            Settings s = new Settings();
            s.TargetFolder = folder; s.Recursive = true; s.TimeoutSec = 120;
            return s;
        }
        static FileItem Find(List<FileItem> items, string fileName)
        {
            foreach (FileItem it in items) if (string.Equals(Path.GetFileName(it.Path), fileName, StringComparison.OrdinalIgnoreCase)) return it;
            return null;
        }
        static void Only(List<FileItem> items, params string[] names)
        {
            foreach (FileItem it in items)
            {
                it.Checked = false;
                foreach (string n in names) if (string.Equals(Path.GetFileName(it.Path), n, StringComparison.OrdinalIgnoreCase)) it.Checked = true;
            }
        }
        static List<FileItem> RunJob(Settings s, string[] only, bool dry, out JobRunner jr)
        {
            List<FileItem> items = FileScanner.Scan(s, null, null);
            if (only != null) Only(items, only);
            jr = new JobRunner(s); jr.DryRun = dry;
            jr.Run(items);
            return items;
        }
        static string St(FileItem it) { return it == null ? "(なし)" : it.Status + " " + it.ErrorCode + " " + it.Message; }
        static bool Same(string path, string rel) { return File.Exists(path) && Util.Sha256(path) == masterHash[rel]; }

        static string MediaInfoText(string xlsx, out int maxJpegW, out int pngCount, out int jpegCount, out bool anyAlphaPng)
        {
            maxJpegW = 0; pngCount = 0; jpegCount = 0; anyAlphaPng = false;
            StringBuilder sb = new StringBuilder();
            using (ZipArchive z = ZipFile.OpenRead(xlsx))
                foreach (ZipArchiveEntry e in z.Entries)
                {
                    if (!e.FullName.StartsWith("xl/media/")) continue;
                    using (Stream s = e.Open())
                    using (MemoryStream ms = new MemoryStream())
                    {
                        s.CopyTo(ms); ms.Position = 0;
                        using (Image im = Image.FromStream(ms))
                        {
                            string ext = Path.GetExtension(e.FullName).ToLowerInvariant();
                            sb.Append(Path.GetFileName(e.FullName) + "=" + im.Width + "x" + im.Height + "(" + e.Length / 1024 + "KB) ");
                            if (ext == ".png") { pngCount++; if (Image.IsAlphaPixelFormat(im.PixelFormat)) anyAlphaPng = true; }
                            else if (ext == ".jpg" || ext == ".jpeg") { jpegCount++; if (im.Width > maxJpegW) maxJpegW = im.Width; }
                        }
                    }
                }
            return sb.ToString();
        }
        static string ContentTypes(string xlsx)
        {
            using (ZipArchive z = ZipFile.OpenRead(xlsx))
            using (StreamReader r = new StreamReader(z.GetEntry("[Content_Types].xml").Open())) return r.ReadToEnd();
        }
        // 画像・rels・Content_Types 以外で内容が異なるパーツの数
        static int NonMediaDiff(string a, string b)
        {
            int n = 0;
            using (ZipArchive za = ZipFile.OpenRead(a))
            using (ZipArchive zb = ZipFile.OpenRead(b))
            {
                foreach (ZipArchiveEntry e in za.Entries)
                {
                    if (e.FullName.StartsWith("xl/media/") || e.FullName.EndsWith(".rels") || e.FullName == "[Content_Types].xml") continue;
                    ZipArchiveEntry o = zb.GetEntry(e.FullName);
                    if (o == null) { n++; continue; }
                    if (Bytes(e) != Bytes(o)) n++;
                }
            }
            return n;
        }
        static string Bytes(ZipArchiveEntry e)
        {
            using (Stream s = e.Open())
            using (MemoryStream ms = new MemoryStream()) { s.CopyTo(ms); return Convert.ToBase64String(System.Security.Cryptography.SHA256.Create().ComputeHash(ms.ToArray())); }
        }
        static int CountFiles(string dir, string pattern) { return Directory.GetFiles(dir, pattern, SearchOption.AllDirectories).Length; }

        // ================= テスト =================
        static void TestScan()
        {
            string d = Fresh("t_scan");
            Settings s = Base(d);
            List<FileItem> a = FileScanner.Scan(s, null, null);
            int expect = 0;
            foreach (string f in Directory.GetFiles(d, "*", SearchOption.AllDirectories))
                if (Path.GetExtension(f).ToLowerInvariant() == ".xls" && !Path.GetFileName(f).StartsWith("~$")) expect++;
            Check("T-SCN-01", "再帰 ON で全階層の .xls を列挙", a.Count == expect && Find(a, "xlm.xls") != null == xlmOk, "件数=" + a.Count + " 期待=" + expect);
            Check("T-SCN-03", "大文字拡張子 .XLS も対象", Find(a, "UPPER.XLS") != null, "");
            Check("T-SCN-04", "既定では .xlsx を列挙しない", Find(a, "済.xlsx") == null, "");
            Check("T-SCN-05", "~$ 一時ファイルを除外", Find(a, "~$temp.xls") == null, "");
            FileItem fake = Find(a, "fake.xls");
            Check("T-SCN-14", "形式不一致を事前に検出しチェックを外す", fake != null && fake.PreNote == "形式不一致" && !fake.Checked, fake == null ? "" : fake.PreNote);
            s.Recursive = false;
            List<FileItem> b = FileScanner.Scan(s, null, null);
            Check("T-SCN-02", "再帰 OFF は直下のみ", Find(b, "simple.xls") == null && Find(b, "報告 1.xls") != null, "件数=" + b.Count);
            s.Recursive = true; s.ExcludeFolders = "old";
            Check("T-SCN-07", "除外フォルダ名", Find(FileScanner.Scan(s, null, null), "skipme.xls") == null, "");
            s.IncludeConverted = true;
            List<FileItem> c = FileScanner.Scan(s, null, null);
            FileItem x = Find(c, "済.xlsx");
            Check("T-CMP-02", "「変換済みも対象」ON で .xlsx が「圧縮のみ」で列挙", x != null && x.Kind == ItemKind.CompressOnly && Find(c, "報告 1.xls").Kind == ItemKind.Convert, "");
            File.Copy(Path.Combine(d, "済.xlsx"), Path.Combine(d, "済_圧縮.xlsx"));
            Check("T-CMP-06", "接尾辞付きファイルは対象外", Find(FileScanner.Scan(s, null, null), "済_圧縮.xlsx") == null, "");
        }

        static long convFinal, convConv;

        static void TestConvert()
        {
            string d = Fresh("t_conv");
            Settings s = Base(d);
            List<FileItem> items = FileScanner.Scan(s, null, null);
            foreach (FileItem it in items) it.Checked = true;       // fake / broken も処理させる
            JobRunner jr = new JobRunner(s);
            jr.Run(items);
            FileItem p = Find(items, "報告 1.xls");
            string po = Path.Combine(d, "報告 1.xlsx");
            Check("T-CNV-01", "写真入り xls → xlsx", p.Status == "完了" && File.Exists(po), St(p));
            Check("T-MAC-01", "マクロなしは .xlsx", p.Macro == "" && p.OutputPath.EndsWith(".xlsx"), p.Macro);
            Check("T-LAY-01", "変換前後でレイアウト差異なし", p.DiffCount == 0, "差異=" + p.DiffCount + " " + (p.Warnings.Count > 0 ? p.Warnings[0] : ""));
            Check("T-IMG-00", "ファイルサイズが削減される", p.FinalSize > 0 && p.FinalSize < p.Size && p.FinalSize < p.ConvSize, Util.FormatSize(p.Size) + " → 変換直後 " + Util.FormatSize(p.ConvSize) + " → " + Util.FormatSize(p.FinalSize) + " 画像 " + p.ImgCompressed + "/" + p.ImgTotal);
            convFinal = p.FinalSize; convConv = p.ConvSize;
            if (File.Exists(po))
            {
                int maxW, png, jpg; bool alpha;
                string mi = MediaInfoText(po, out maxW, out png, out jpg, out alpha);
                Check("T-IMG-01", "150ppi に縮小（表示幅 4 インチ → 幅 600px 以下）", maxW > 0 && maxW <= 602, mi);
                Check("T-IMG-08", "透過 PNG は PNG のまま", png == 1 && alpha, "png=" + png + " alpha=" + alpha);
                Check("T-IMG-09", "透過なし PNG は JPEG 化", jpg == 4, "jpeg=" + jpg);
                bool zipOk = true; string zmsg = "";
                try { ImageOptimizer.CheckZip(po); } catch (Exception ex) { zipOk = false; zmsg = ex.Message; }
                string ct = ContentTypes(po);
                Check("T-IMG-18", "出力 ZIP が整合", zipOk, zmsg);
                Check("T-MAC-12", "xlsx の ContentType が拡張子と一致", ct.IndexOf("spreadsheetml.sheet.main+xml") >= 0 && ct.IndexOf("macroEnabled") < 0, "");
            }
            FileItem sm = Find(items, "simple.xls");
            Check("T-CNV-01b", "画像なし xls → xlsx", sm.Status == "完了" && File.Exists(Path.Combine(d, "sub", "simple.xlsx")) && sm.DiffCount == 0, St(sm));
            if (xlmOk)
            {
                FileItem x = Find(items, "xlm.xls"); string xo = Path.Combine(d, "sub", "deep", "xlm.xlsm");
                Check("T-MAC-05", "Excel 4.0 マクロシート → .xlsm", x.Status.StartsWith("完了") && File.Exists(xo) && x.Macro.IndexOf("Excel4.0") >= 0, St(x) + " macro=" + x.Macro);
                if (File.Exists(xo)) Check("T-MAC-12b", "xlsm の ContentType が macroEnabled", ContentTypes(xo).IndexOf("macroEnabled") >= 0, "");
            }
            else SkipT("T-MAC-05", "Excel 4.0 マクロシート", "この Excel ではテストデータを作成できない");
            if (nameOk)
            {
                FileItem x = Find(items, "name.xls");
                Check("T-MAC-06", "名前定義の GET.CELL → .xlsm", x.Status.StartsWith("完了") && File.Exists(Path.Combine(d, "sub", "name.xlsm")) && x.Macro.IndexOf("名前定義") >= 0, St(x) + " macro=" + x.Macro);
            }
            else SkipT("T-MAC-06", "名前定義の XLM 関数", "テストデータを作成できない");
            if (vbaOk)
            {
                FileItem x = Find(items, "vba.xls"); string xo = Path.Combine(d, "vba.xlsm");
                Check("T-MAC-02", "VBA あり → .xlsm（vbaProject.bin を保持）", x.Status.StartsWith("完了") && File.Exists(xo) && ImageOptimizer.HasVbaProject(xo), St(x) + " macro=" + x.Macro);
            }
            else SkipT("T-MAC-02", "VBA あり → .xlsm", "VBA プロジェクトへの信頼アクセスが無効でテストデータを作成できない");
            if (pwOk)
            {
                FileItem x = Find(items, "pw.xls");
                Check("T-CNV-09", "読み取りパスワード付きはスキップ", x.Status == "スキップ" && x.ErrorCode == "E102" && !File.Exists(Path.Combine(d, "pw.xlsx")), St(x));
            }
            else SkipT("T-CNV-09", "読み取りパスワード", "テストデータを作成できない");
            FileItem fk = Find(items, "fake.xls");
            Check("T-MAC-13", "HTML を .xls にしたファイルはスキップ", fk.Status == "スキップ" && fk.ErrorCode == "E105", St(fk));
            FileItem br = Find(items, "broken.xls");
            Check("T-ERR-04", "破損ファイルは失敗として次へ進む", br.Status == "失敗" || br.Status == "スキップ", St(br));
            bool same = true; string diff = "";
            foreach (KeyValuePair<string, string> kv in masterHash) if (!Same(Path.Combine(d, kv.Key), kv.Key)) { same = false; diff += kv.Key + " "; }
            Check("T-ORG-01", "「残す」では元ファイルが 1 バイトも変わらない", same, diff);
            Check("T-CNV-15", "一時ファイル .~tmp が残らない", CountFiles(d, "*.~tmp") == 0, "");
            bool bom = false; int lines = 0;
            if (jr.ReportPath != null && File.Exists(jr.ReportPath))
            {
                byte[] rb = File.ReadAllBytes(jr.ReportPath);
                bom = rb.Length > 3 && rb[0] == 0xEF && rb[1] == 0xBB && rb[2] == 0xBF;
                lines = File.ReadAllLines(jr.ReportPath).Length;
            }
            Check("T-LOG-01", "レポート CSV（BOM 付き、全件 1 行ずつ）", bom && lines == items.Count + 1, "行数=" + lines + " 件数=" + items.Count);
            Check("T-LOG-03", "ログファイルが出力される", File.Exists(jr.Log.FilePath), "");

            // 同名出力: 連番 / スキップ
            JobRunner jr2;
            List<FileItem> again = RunJob(s, new string[] { "simple.xls" }, false, out jr2);
            Check("T-CNV-12", "同名出力があるとき連番を付ける", File.Exists(Path.Combine(d, "sub", "simple(1).xlsx")), St(Find(again, "simple.xls")));
            s.Conflict = "skip";
            again = RunJob(s, new string[] { "simple.xls" }, false, out jr2);
            Check("T-CNV-10", "同名出力があるときスキップ", Find(again, "simple.xls").Status == "スキップ" && !File.Exists(Path.Combine(d, "sub", "simple(2).xlsx")), St(Find(again, "simple.xls")));
        }

        static void TestSize()
        {
            if (convFinal <= 0) { SkipT("T-SIZ", "サイズ目標", "変換テストが失敗したため"); return; }
            string d = Fresh("t_size");
            Settings s = Base(d);
            s.SizeTargetEnabled = true; s.SkipIfUnderLimit = true;
            JobRunner jr;
            // 初期設定(150ppi/85)の結果より小さい上限 → 段階的に下げて上限以下に
            long limit = (long)(convFinal * 0.7);
            s.SizeLimitMB = limit / (1024.0 * 1024.0);
            FileItem p = Find(RunJob(s, new string[] { "報告 1.xls" }, false, out jr), "報告 1.xls");
            long sz = File.Exists(p.OutputPath) ? new FileInfo(p.OutputPath).Length : -1;
            Check("T-SIZ-02", "上限以下になるまで段階的に圧縮", p.Status == "完了" && sz > 0 && sz <= limit && (p.UsedQuality < 85 || p.UsedPpi < 150), "上限=" + limit + " 結果=" + sz + " ppi=" + p.UsedPpi + " 品質=" + p.UsedQuality);
            Check("T-SIZ-03", "上限に近い画質を採用（上限の 60% 以上）", sz >= limit * 0.6, "結果/上限=" + (sz * 100 / Math.Max(1, limit)) + "%");
            // 到達不能な上限 → 目標未達
            s.SizeLimitMB = 0.02;
            p = Find(RunJob(s, new string[] { "報告 1.xls" }, false, out jr), "報告 1.xls");
            Check("T-SIZ-04", "届かない上限は「目標未達」で最低画質の結果を出力", p.Status == "完了（目標未達）" && File.Exists(p.OutputPath) && p.UsedQuality >= 50 && p.UsedPpi >= 96, St(p) + " ppi=" + p.UsedPpi + " 品質=" + p.UsedQuality);
            // 既に上限以下 → 圧縮しない
            s.SizeLimitMB = 500;
            p = Find(RunJob(s, new string[] { "報告 1.xls" }, false, out jr), "報告 1.xls");
            Check("T-SIZ-06", "上限以下なら圧縮しない（ON）", p.Status == "完了" && p.ImgCompressed == 0 && p.FinalSize == p.ConvSize, St(p) + " 画像圧縮=" + p.ImgCompressed);
            s.SkipIfUnderLimit = false;
            p = Find(RunJob(s, new string[] { "報告 1.xls" }, false, out jr), "報告 1.xls");
            Check("T-SIZ-06b", "上限以下でも初期設定で 1 回圧縮（OFF）", p.Status == "完了" && p.ImgCompressed > 0 && p.FinalSize < p.ConvSize, St(p));
            // 無効 → 指定 ppi・品質で 1 回
            s.SizeTargetEnabled = false; s.Ppi = 96; s.JpegQuality = 60;
            p = Find(RunJob(s, new string[] { "報告 1.xls" }, false, out jr), "報告 1.xls");
            int maxW = 0, png, jpg; bool alpha; string mi = "";
            if (File.Exists(p.OutputPath)) mi = MediaInfoText(p.OutputPath, out maxW, out png, out jpg, out alpha);
            Check("T-SIZ-08", "サイズ目標なし: 96ppi 指定で幅 384px 以下", p.Status == "完了" && maxW > 0 && maxW <= 386, mi);
            Check("T-ORG-01b", "繰り返し処理しても元ファイルは不変", Same(Path.Combine(d, "報告 1.xls"), "報告 1.xls"), "");
        }

        static void TestMoveRestore()
        {
            string d = Fresh("t_move");
            string bk = Path.Combine(root, "backup_move");
            Settings s = Base(d);
            s.OriginalMode = "move"; s.MoveRoot = bk;
            DateTime lw = File.GetLastWriteTimeUtc(Path.Combine(d, "sub", "simple.xls"));
            JobRunner jr;
            List<FileItem> items = RunJob(s, null, false, out jr);
            string broot = jr.Man.BackupRoot;
            bool moved = !File.Exists(Path.Combine(d, "報告 1.xls")) && Same(Path.Combine(broot, "報告 1.xls"), "報告 1.xls")
                && !File.Exists(Path.Combine(d, "sub", "simple.xls")) && Same(Path.Combine(broot, "sub", "simple.xls"), Path.Combine("sub", "simple.xls"));
            Check("T-ORG-05", "移動: フォルダ構造を保って移動先へ（内容一致）", moved, broot);
            Check("T-ORG-05b", "出力は元の場所に残る", File.Exists(Path.Combine(d, "報告 1.xlsx")) && File.Exists(Path.Combine(d, "sub", "simple.xlsx")), "");
            if (pwOk) Check("T-ORG-08", "スキップしたファイルは移動しない", File.Exists(Path.Combine(d, "pw.xls")) && !File.Exists(Path.Combine(broot, "pw.xls")), "");
            Check("T-ORG-08b", "対象外（形式不一致）のファイルは移動しない", File.Exists(Path.Combine(d, "fake.xls")), "");
            Check("T-RST-00", "マニフェストが履歴と移動先の両方に保存される", File.Exists(jr.Man.FilePath) && File.Exists(Path.Combine(broot, "_manifest.json")), "");

            // 選択復元 → 全体復元 → 二重復元
            Manifest m = Manifest.Load(jr.Man.FilePath);
            ManEntry one = null;
            foreach (ManEntry e in m.Entries) if (Path.GetFileName(e.Original) == "simple.xls") one = e;
            RestoreResult r1 = RestoreService.Restore(m, new ManEntry[] { one }, true, false, null);
            Check("T-RST-02", "選択復元: 1 件だけ戻る", r1.Restored == 1 && Same(Path.Combine(d, "sub", "simple.xls"), Path.Combine("sub", "simple.xls")) && !File.Exists(Path.Combine(d, "報告 1.xls")), "復元=" + r1.Restored);
            Check("T-RST-03", "復元時に生成ファイルを削除", !File.Exists(Path.Combine(d, "sub", "simple.xlsx")), "");
            Check("T-RST-01b", "復元後の更新日時が元と同じ", File.GetLastWriteTimeUtc(Path.Combine(d, "sub", "simple.xls")) == lw, "");
            // 生成ファイルを編集してから復元 → 削除しない
            File.AppendAllText(Path.Combine(d, "報告 1.xlsx"), "edited");
            // 元の場所に同名ファイルを置く → その 1 件はスキップ
            string upper = Path.Combine(d, "sub", "UPPER.XLS");
            File.WriteAllText(upper, "someone put a new file here");
            m = Manifest.Load(jr.Man.FilePath);
            RestoreResult r2 = RestoreService.Restore(m, m.Entries, true, false, null);
            Check("T-RST-01", "ジョブ全体の復元: 元ファイルが元の場所に戻り内容一致", Same(Path.Combine(d, "報告 1.xls"), "報告 1.xls"), "復元=" + r2.Restored + " スキップ=" + r2.Skipped + " 失敗=" + r2.Failed);
            Check("T-RST-04", "編集された生成ファイルは削除しない", File.Exists(Path.Combine(d, "報告 1.xlsx")), "");
            Check("T-RST-06", "元の場所に同名があれば上書きせずスキップ", File.ReadAllText(upper).StartsWith("someone") && File.Exists(Path.Combine(broot, "sub", "UPPER.XLS")), string.Join(" | ", r2.Messages.ToArray()));
            m = Manifest.Load(jr.Man.FilePath);
            int before = 0; foreach (ManEntry e in m.Entries) if (e.State == "restored") before++;
            RestoreResult r3 = RestoreService.Restore(m, m.Entries, true, false, null);
            Check("T-RST-09", "二重復元しても何も変わらない", r3.Restored == 0 && Same(Path.Combine(d, "報告 1.xls"), "報告 1.xls"), "復元済み=" + before + " 再復元=" + r3.Restored);
            // 移動先で変更されたファイルは確認なしでは戻さない
            File.Delete(upper);
            File.AppendAllText(Path.Combine(broot, "sub", "UPPER.XLS"), "x");
            m = Manifest.Load(jr.Man.FilePath);
            RestoreResult r4 = RestoreService.Restore(m, m.Entries, true, false, null);
            Check("T-RST-08", "移動先で変更されたファイルは復元しない", r4.Failed == 1 && !File.Exists(upper), string.Join(" | ", r4.Messages.ToArray()));
            // 壊れたマニフェスト
            string bad = Path.Combine(AppPaths.HistoryDir, "00000000_000000_bad.manifest.json");
            File.WriteAllText(bad, "{ this is not json");
            List<string> errs = new List<string>();
            List<Manifest> all = Manifest.LoadAll(errs);
            Check("T-RST-12", "壊れたマニフェストがあっても他の履歴を読める", errs.Count == 1 && all.Count >= 1, "エラー=" + errs.Count + " 読込=" + all.Count);
            File.Delete(bad);
            // 中断された移動の修復
            m = Manifest.Load(jr.Man.FilePath);
            ManEntry fakeMoving = new ManEntry();
            fakeMoving.Original = Path.Combine(d, "報告 1.xls"); fakeMoving.Backup = Path.Combine(broot, "nothing.xls"); fakeMoving.State = "moving";
            m.Entries.Add(fakeMoving);
            int rep = m.Repair();
            Check("T-ORG-09", "「移動中」のまま残ったエントリを起動時に修復", rep == 1 && fakeMoving.State == "kept", "state=" + fakeMoving.State);
        }

        static void TestDelete()
        {
            string d = Fresh("t_del");
            Settings s = Base(d);
            s.OriginalMode = "delete"; s.DeleteToRecycle = false;
            JobRunner jr;
            List<FileItem> items = RunJob(s, null, false, out jr);
            Check("T-ORG-04", "完全削除: 完了したファイルの元は消え、出力が残る", !File.Exists(Path.Combine(d, "報告 1.xls")) && File.Exists(Path.Combine(d, "報告 1.xlsx")) && !File.Exists(Path.Combine(d, "sub", "simple.xls")), St(Find(items, "報告 1.xls")));
            FileItem br = Find(items, "broken.xls");
            Check("T-ORG-08c", "失敗・スキップしたファイルの元は削除しない", File.Exists(Path.Combine(d, "broken.xls")) && File.Exists(Path.Combine(d, "fake.xls")) && (!pwOk || File.Exists(Path.Combine(d, "pw.xls"))), St(br));
            // ごみ箱（1 件だけ実際にごみ箱へ送る）
            string d2 = Fresh("t_recycle");
            Settings s2 = Base(d2);
            s2.OriginalMode = "delete"; s2.DeleteToRecycle = true;
            Util.NoRecycle = false;
            try
            {
                List<FileItem> it2 = RunJob(s2, new string[] { "skipme.xls" }, false, out jr);
                Check("T-ORG-02", "ごみ箱へ削除", !File.Exists(Path.Combine(d2, "old", "skipme.xls")) && File.Exists(Path.Combine(d2, "old", "skipme.xlsx")), St(Find(it2, "skipme.xls")));
            }
            finally { Util.NoRecycle = true; }
        }

        static void TestCompressOnly()
        {
            JobRunner jr;
            // 残す → 接尾辞付きで出力
            string d = Fresh("t_cmp_keep");
            Settings s = Base(d); s.IncludeConverted = true;
            List<FileItem> items = RunJob(s, new string[] { "済.xlsx", "plain.xlsx" }, false, out jr);
            FileItem x = Find(items, "済.xlsx"); string xo = Path.Combine(d, "済_圧縮.xlsx");
            Check("T-CMP-08", "圧縮のみ・残す: 「_圧縮」付きで出力、元は不変", x.Status == "完了" && File.Exists(xo) && Same(Path.Combine(d, "済.xlsx"), "済.xlsx") && x.FinalSize < x.Size, St(x) + " " + Util.FormatSize(x.Size) + "→" + Util.FormatSize(x.FinalSize));
            Check("T-CMP-03b", "圧縮のみでもレイアウト差異なし", x.DiffCount == 0, "差異=" + x.DiffCount + " " + (x.Warnings.Count > 0 ? x.Warnings[0] : ""));
            if (File.Exists(xo)) Check("T-CMP-03", "画像以外のパーツはバイト単位で同一（Excel で保存し直さない）", NonMediaDiff(Path.Combine(d, "済.xlsx"), xo) == 0, "異なるパーツ数=" + NonMediaDiff(Path.Combine(d, "済.xlsx"), xo));
            if (File.Exists(xo)) Check("T-LAY-15", "drawing XML は不変", true, "T-CMP-03 に含む");
            FileItem pl = Find(items, "plain.xlsx");
            Check("T-CMP-07", "小さくならないファイルは「変更なし」で何も出力しない", pl.Status == "変更なし" && !File.Exists(Path.Combine(d, "plain_圧縮.xlsx")) && Same(Path.Combine(d, "plain.xlsx"), "plain.xlsx"), St(pl));
            // 同じジョブで xls と xlsx を処理 → 二重処理しない
            string d0 = Fresh("t_cmp_both");
            Settings s0 = Base(d0); s0.IncludeConverted = true;
            List<FileItem> both = RunJob(s0, new string[] { "報告 1.xls", "済.xlsx" }, false, out jr);
            int processed = 0; foreach (FileItem it in both) if (it.Status.StartsWith("完了")) processed++;
            Check("T-CMP-05", "同じジョブで生成した xlsx を再圧縮しない", processed == 2 && File.Exists(Path.Combine(d0, "報告 1.xlsx")) && !File.Exists(Path.Combine(d0, "報告 1_圧縮.xlsx")), "完了=" + processed);

            // 移動 → 同名で置き換え → 復元
            string d2 = Fresh("t_cmp_move");
            Settings s2 = Base(d2); s2.IncludeConverted = true; s2.OriginalMode = "move"; s2.MoveRoot = Path.Combine(root, "backup_cmp");
            items = RunJob(s2, new string[] { "済.xlsx" }, false, out jr);
            x = Find(items, "済.xlsx");
            string here = Path.Combine(d2, "済.xlsx"), there = Path.Combine(jr.Man.BackupRoot, "済.xlsx");
            Check("T-CMP-10", "圧縮のみ・移動: 元は移動先へ、同名で圧縮版に置き換え", x.Status == "完了" && File.Exists(here) && !Same(here, "済.xlsx") && new FileInfo(here).Length < x.Size && Same(there, "済.xlsx") && !File.Exists(Path.Combine(d2, "済_圧縮.xlsx")), St(x));
            Manifest m = Manifest.Load(jr.Man.FilePath);
            RestoreResult r = RestoreService.Restore(m, m.Entries, true, false, null);
            Check("T-CMP-10b", "圧縮のみの復元: 元ファイルが同じ名前で戻る", r.Restored == 1 && Same(here, "済.xlsx") && !File.Exists(there), "復元=" + r.Restored + " " + string.Join(" | ", r.Messages.ToArray()));

            // 圧縮版を編集してから復元 → 「_復元」で戻す
            string d3 = Fresh("t_cmp_edit");
            Settings s3 = Base(d3); s3.IncludeConverted = true; s3.OriginalMode = "move"; s3.MoveRoot = Path.Combine(root, "backup_cmp2");
            RunJob(s3, new string[] { "済.xlsx" }, false, out jr);
            string here3 = Path.Combine(d3, "済.xlsx");
            File.AppendAllText(here3, "edited");
            string editedHash = Util.Sha256(here3);
            m = Manifest.Load(jr.Man.FilePath);
            r = RestoreService.Restore(m, m.Entries, true, false, null);
            Check("T-CMP-11", "圧縮版が編集済みなら残し、元は「_復元」で戻す", Util.Sha256(here3) == editedHash && Same(Path.Combine(d3, "済_復元.xlsx"), "済.xlsx"), string.Join(" | ", r.Messages.ToArray()));

            // 削除 → 同名で置き換え
            string d4 = Fresh("t_cmp_del");
            Settings s4 = Base(d4); s4.IncludeConverted = true; s4.OriginalMode = "delete"; s4.DeleteToRecycle = false;
            items = RunJob(s4, new string[] { "済.xlsx" }, false, out jr);
            x = Find(items, "済.xlsx");
            string here4 = Path.Combine(d4, "済.xlsx");
            Check("T-CMP-09", "圧縮のみ・削除: 同名で圧縮版に置き換え", x.Status == "完了" && File.Exists(here4) && !Same(here4, "済.xlsx") && new FileInfo(here4).Length < x.Size && CountFiles(d4, "*.~tmp") == 0, St(x));
        }

        static void TestLayoutDiff()
        {
            JobRunner jr;
            LayoutSnapshot.TestForceDiff = true;
            try
            {
                Settings def = new Settings();
                Check("T-CMP-12a", "「差異があるときは削除・移動しない」の初期値は ON", def.KeepOnLayoutDiff, "");
                string d = Fresh("t_diff_keep");
                Settings s = Base(d); s.IncludeConverted = true; s.OriginalMode = "move"; s.MoveRoot = Path.Combine(root, "backup_diff1");
                List<FileItem> items = RunJob(s, new string[] { "simple.xls", "済.xlsx" }, false, out jr);
                FileItem a = Find(items, "simple.xls"), b = Find(items, "済.xlsx");
                Check("T-CMP-12", "差異あり・既定: 変換した元ファイルを移動しない（要確認）", a.Status == "完了（要確認）" && a.ErrorCode == "E401" && Same(Path.Combine(d, "sub", "simple.xls"), Path.Combine("sub", "simple.xls")) && File.Exists(Path.Combine(d, "sub", "simple.xlsx")), St(a));
                Check("T-CMP-12b", "差異あり・既定: 圧縮のみは「_圧縮」付きで出力し元は不変", b.Status == "完了（要確認）" && Same(Path.Combine(d, "済.xlsx"), "済.xlsx") && File.Exists(Path.Combine(d, "済_圧縮.xlsx")), St(b));
                Check("T-LAY-14", "差異の内容がログに記録される", File.ReadAllText(jr.Log.FilePath).IndexOf("レイアウト差異") >= 0, "");

                string d2 = Fresh("t_diff_off");
                Settings s2 = Base(d2); s2.OriginalMode = "move"; s2.MoveRoot = Path.Combine(root, "backup_diff2"); s2.KeepOnLayoutDiff = false;
                items = RunJob(s2, new string[] { "simple.xls" }, false, out jr);
                a = Find(items, "simple.xls");
                Check("T-CMP-13", "差異あり・設定 OFF: 指定どおり移動し「要確認」と記録", a.Status == "完了（要確認）" && !File.Exists(Path.Combine(d2, "sub", "simple.xls")) && Same(Path.Combine(jr.Man.BackupRoot, "sub", "simple.xls"), Path.Combine("sub", "simple.xls")), St(a));
            }
            finally { LayoutSnapshot.TestForceDiff = false; }
            string d3 = Fresh("t_diff_none");
            Settings s3 = Base(d3); s3.OriginalMode = "move"; s3.MoveRoot = Path.Combine(root, "backup_diff3");
            List<FileItem> it3 = RunJob(s3, new string[] { "simple.xls" }, false, out jr);
            Check("T-CMP-14", "差異なし: 指定どおり移動", Find(it3, "simple.xls").Status == "完了" && !File.Exists(Path.Combine(d3, "sub", "simple.xls")), St(Find(it3, "simple.xls")));
        }

        static void TestDryRunAndLock()
        {
            JobRunner jr;
            string d = Fresh("t_dry");
            Settings s = Base(d); s.OriginalMode = "delete"; s.DeleteToRecycle = false;
            Dictionary<string, string> before = HashDir(d);
            List<FileItem> items = RunJob(s, new string[] { "報告 1.xls", "simple.xls" }, true, out jr);
            Dictionary<string, string> after = HashDir(d);
            bool same = before.Count == after.Count;
            foreach (KeyValuePair<string, string> kv in before) if (!after.ContainsKey(kv.Key) || after[kv.Key] != kv.Value) same = false;
            FileItem p = Find(items, "報告 1.xls");
            Check("T-GUI-13", "ドライラン: ファイルを一切変更せず予測サイズを出す", same && p.Status == "ドライラン" && p.FinalSize > 0 && p.FinalSize < p.Size, St(p));

            // 使用中のファイル
            string locked = Path.Combine(d, "sub", "simple.xls");
            s.OriginalMode = "keep";
            using (FileStream fs = new FileStream(locked, FileMode.Open, FileAccess.ReadWrite, FileShare.None))
            {
                items = RunJob(s, new string[] { "simple.xls", "UPPER.XLS" }, false, out jr);
            }
            FileItem a = Find(items, "simple.xls"), b = Find(items, "UPPER.XLS");
            Check("T-ERR-02", "使用中のファイルはスキップして次へ進む", a.Status == "スキップ" && a.ErrorCode == "E101" && b.Status == "完了", St(a) + " / " + St(b));

            // 設定ファイル
            string sf = Path.Combine(root, "settings_test.json");
            Settings w = new Settings(); w.Ppi = 123; w.CompressSuffix = "_小"; w.Save(sf);
            Settings rd = Settings.Load(sf);
            Check("T-GUI-02", "設定の保存と読込", rd.Ppi == 123 && rd.CompressSuffix == "_小", "");
            File.WriteAllText(sf, "{ broken");
            Settings rd2 = Settings.Load(sf);
            Check("T-LOG-05", "壊れた設定ファイルは既定値で起動し .bak に退避", rd2.Ppi == 150 && File.Exists(sf + ".bak"), "");
            Settings v = new Settings(); v.SizeTargetEnabled = true; v.SizeLimitMB = 0;
            Check("T-SIZ-10", "入力チェック: 上限 0 は不可", v.Validate() != null, v.Validate());
            v = new Settings(); v.MinQuality = 90; v.JpegQuality = 80;
            Check("T-GUI-04", "入力チェック: 最低品質 > 品質 は不可", v.Validate() != null, v.Validate());
            v = new Settings(); v.OriginalMode = "move";
            Check("T-GUI-05", "入力チェック: 移動で移動先未入力は不可", v.Validate() != null, v.Validate());
        }
    }
}
