using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using System.Text.RegularExpressions;

namespace XlsImageShrink
{
    /// <summary>専用の Excel プロセスを 1 つ管理する（利用者の Excel には接続しない）。</summary>
    public class ExcelSession : IDisposable
    {
        dynamic app;
        int pid;
        int opened;
        readonly Logger log;
        readonly int timeoutSec;
        readonly int restartEvery;
        DateTime deadline = DateTime.MaxValue;
        System.Threading.Timer watchdog;
        public bool TimedOut;

        [DllImport("user32.dll")]
        static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint processId);

        public ExcelSession(Logger log, int timeoutSec, int restartEvery)
        {
            this.log = log; this.timeoutSec = Math.Max(10, timeoutSec); this.restartEvery = Math.Max(1, restartEvery);
            watchdog = new System.Threading.Timer(Tick, null, 2000, 2000);
        }

        public static bool IsInstalled()
        {
            try { return Type.GetTypeFromProgID("Excel.Application") != null; } catch { return false; }
        }

        public int ProcessId { get { return pid; } }

        void Tick(object state)
        {
            if (DateTime.Now > deadline && pid != 0)
            {
                deadline = DateTime.MaxValue;
                TimedOut = true;
                if (log != null) log.Error("Excel が応答しないため強制終了します (PID " + pid + ")");
                KillProcess();
            }
        }

        void KillProcess()
        {
            try { Process p = Process.GetProcessById(pid); p.Kill(); p.WaitForExit(5000); } catch { }
        }

        void Ensure()
        {
            if (app != null) return;
            Type t = Type.GetTypeFromProgID("Excel.Application");
            if (t == null) throw new InvalidOperationException("Excel がインストールされていません。");
            app = Activator.CreateInstance(t);
            app.Visible = false;
            app.DisplayAlerts = false;
            app.ScreenUpdating = false;
            app.EnableEvents = false;
            try { app.AutomationSecurity = 3; } catch { }
            try { app.AskToUpdateLinks = false; } catch { }
            uint p;
            GetWindowThreadProcessId(new IntPtr((int)app.Hwnd), out p);
            pid = (int)p;
            try
            {
                app.Workbooks.Add();
                app.Calculation = -4135;        // xlCalculationManual
                app.CalculateBeforeSave = false;
            }
            catch { }
            opened = 0;
            if (log != null) log.Info("Excel を起動しました (PID " + pid + ", Version " + Convert.ToString((object)app.Version) + ")");
        }

        void Begin() { TimedOut = false; deadline = DateTime.Now.AddSeconds(timeoutSec); }
        void End() { deadline = DateTime.MaxValue; }

        public dynamic Open(string path)
        {
            if (app != null && opened >= restartEvery) { if (log != null) log.Info("Excel を再起動します"); Close(); }
            Ensure();
            opened++;
            Begin();
            try
            {
                return app.Workbooks.Open(Filename: path, UpdateLinks: 0, ReadOnly: true, Password: "\u0001nopass",
                    IgnoreReadOnlyRecommended: true, AddToMru: false);
            }
            finally { End(); }
        }

        public void SaveAs(dynamic wb, string path, int fileFormat)
        {
            Begin();
            try
            {
                try { wb.CheckCompatibility = false; } catch { }
                wb.SaveAs(Filename: path, FileFormat: fileFormat, AddToMru: false);
            }
            finally { End(); }
        }

        public void CloseWorkbook(dynamic wb)
        {
            if (wb == null) return;
            Begin();
            try { wb.Close(false); } catch { }
            finally { End(); }
            try { Marshal.FinalReleaseComObject(wb); } catch { }
            GC.Collect(); GC.WaitForPendingFinalizers();
        }

        /// <summary>タイムアウトなどで Excel が死んだ後に呼ぶ。</summary>
        public void Abandon()
        {
            KillProcess();
            app = null; pid = 0;
            GC.Collect(); GC.WaitForPendingFinalizers();
        }

        public bool IsAlive()
        {
            if (pid == 0) return false;
            try { Process p = Process.GetProcessById(pid); return !p.HasExited; } catch { return false; }
        }

        public void Close()
        {
            if (app == null) return;
            int oldPid = pid;
            try
            {
                try { foreach (dynamic wb in app.Workbooks) { try { wb.Close(false); } catch { } } } catch { }
                app.Quit();
            }
            catch { }
            try { Marshal.FinalReleaseComObject(app); } catch { }
            app = null; pid = 0;
            GC.Collect(); GC.WaitForPendingFinalizers();
            GC.Collect(); GC.WaitForPendingFinalizers();
            try
            {
                Process p = Process.GetProcessById(oldPid);
                if (!p.WaitForExit(10000)) { p.Kill(); if (log != null) log.Warn("Excel が終了しなかったため強制終了しました (PID " + oldPid + ")"); }
            }
            catch { }
        }

        public void Dispose()
        {
            Close();
            if (watchdog != null) { watchdog.Dispose(); watchdog = null; }
        }
    }

    public static class MacroDetector
    {
        static readonly Regex Xlm = new Regex(@"\b(GET\.(CELL|WORKBOOK|DOCUMENT|FORMULA|WINDOW|WORKSPACE|OBJECT|NAME|NOTE|DEF|CHART\.ITEM|BAR|TOOL|LINK\.INFO)|EVALUATE|FILES|DOCUMENTS|WINDOWS|NAMES|LINKS|CALLER|SELECTION|ACTIVE\.CELL|REFTEXT|TEXTREF|DEREF|GET\.PIVOT\.\w+)\s*\(", RegexOptions.IgnoreCase | RegexOptions.Compiled);

        /// <summary>マクロ種別を返す。空文字ならマクロなし。</summary>
        public static string Detect(dynamic wb, Settings s, Logger log, List<string> warnings)
        {
            List<string> kinds = new List<string>();
            bool vba = false;
            try { vba = (bool)wb.HasVBProject; } catch { vba = true; warnings.Add("VBA の有無を判定できないため xlsm とします"); }
            if (vba && s.EmptyVbaAsXlsx)
            {
                try
                {
                    int lines = 0;
                    foreach (dynamic c in wb.VBProject.VBComponents) lines += (int)c.CodeModule.CountOfLines;
                    if (lines == 0) vba = false;
                }
                catch { warnings.Add("VBA プロジェクトの中身を確認できないため xlsm とします（信頼設定が無効）"); }
            }
            if (vba) kinds.Add("VBA");
            try
            {
                int n = (int)wb.Excel4MacroSheets.Count + (int)wb.Excel4IntlMacroSheets.Count;
                if (n > 0) kinds.Add("Excel4.0マクロシート");
            }
            catch { }
            try
            {
                bool hit = false;
                foreach (dynamic nm in wb.Names)
                {
                    try
                    {
                        string r = Convert.ToString((object)nm.RefersTo);
                        if (r != null && Xlm.IsMatch(r)) { hit = true; break; }
                    }
                    catch { }
                }
                if (hit) kinds.Add("名前定義XLM関数");
            }
            catch { }
            return string.Join("+", kinds.ToArray());
        }
    }

    public class ShapeSnap
    {
        public string Name; public int Type; public double Top, Left, Width, Height, Rotation; public int Placement; public int Visible;
    }

    public class SheetSnap
    {
        public string Name; public int Type; public int Visible;
        public bool Detailed;
        public string Used = ""; public string Merge = ""; public string Page = ""; public string Sample = "";
        public List<double> RowH = new List<double>();
        public List<double> ColW = new List<double>();
        public List<ShapeSnap> Shapes = new List<ShapeSnap>();
    }

    public static class LayoutSnapshot
    {
        public static bool TestForceDiff = false;   // 自己テスト用

        static double D(object o)
        {
            if (o == null || o is DBNull) return -1;
            try { return Convert.ToDouble(o); } catch { return -1; }
        }
        static int I(object o)
        {
            if (o == null || o is DBNull) return -1;
            try { return Convert.ToInt32(o); } catch { return -1; }
        }
        static string S(object o)
        {
            if (o == null || o is DBNull) return "";
            if (o is double) return Math.Round((double)o, 1).ToString("R");
            if (o is float) return Math.Round((double)(float)o, 1).ToString("R");
            return Convert.ToString(o);
        }

        public static List<SheetSnap> Take(dynamic wb)
        {
            List<SheetSnap> list = new List<SheetSnap>();
            foreach (dynamic sh in wb.Sheets)
            {
                SheetSnap s = new SheetSnap();
                s.Name = Convert.ToString((object)sh.Name);
                try { s.Type = I((object)sh.Type); } catch { s.Type = 0; }
                try { s.Visible = I((object)sh.Visible); } catch { s.Visible = 0; }
                if (s.Type == -4167) { s.Detailed = true; Detail(sh, s); }
                list.Add(s);
            }
            return list;
        }

        static void Detail(dynamic sh, SheetSnap s)
        {
            try
            {
                dynamic ur = sh.UsedRange;
                int r0 = I((object)ur.Row), c0 = I((object)ur.Column);
                int rc = I((object)ur.Rows.Count), cc = I((object)ur.Columns.Count);
                s.Used = Convert.ToString((object)ur.Address);
                try { s.Merge = S((object)ur.MergeCells); } catch { s.Merge = "ERR"; }
                int rn = Math.Min(rc, 300), cn = Math.Min(cc, 100);
                for (int i = 0; i < rn; i++) s.RowH.Add(D((object)sh.Rows[r0 + i].RowHeight));
                for (int j = 0; j < cn; j++) s.ColW.Add(D((object)sh.Columns[c0 + j].ColumnWidth));
                try
                {
                    dynamic part = sh.Range[sh.Cells[r0, c0], sh.Cells[r0 + Math.Min(rc, 50) - 1, c0 + Math.Min(cc, 20) - 1]];
                    object v = (object)part.Value2;
                    StringBuilder sb = new StringBuilder();
                    Array arr = v as Array;
                    if (arr != null) { foreach (object o in arr) { sb.Append(Cell(o)); sb.Append('\u0001'); } }
                    else sb.Append(Cell(v));
                    s.Sample = sb.ToString();
                }
                catch (Exception ex) { s.Sample = "ERR:" + ex.GetType().Name; }
            }
            catch (Exception ex) { s.Used = "ERR:" + ex.GetType().Name; }

            try
            {
                foreach (dynamic sp in sh.Shapes)
                {
                    ShapeSnap p = new ShapeSnap();
                    p.Name = Convert.ToString((object)sp.Name);
                    p.Type = I((object)sp.Type);
                    p.Top = D((object)sp.Top); p.Left = D((object)sp.Left);
                    p.Width = D((object)sp.Width); p.Height = D((object)sp.Height);
                    try { p.Rotation = D((object)sp.Rotation); } catch { p.Rotation = 0; }
                    try { p.Placement = I((object)sp.Placement); } catch { p.Placement = -1; }
                    try { p.Visible = I((object)sp.Visible); } catch { p.Visible = -1; }
                    s.Shapes.Add(p);
                }
            }
            catch { }

            try
            {
                dynamic ps = sh.PageSetup;
                StringBuilder sb = new StringBuilder();
                sb.Append(S((object)ps.Orientation)).Append('|');
                sb.Append(S((object)ps.PaperSize)).Append('|');
                sb.Append(S((object)ps.Zoom)).Append('|');
                sb.Append(S((object)ps.FitToPagesWide)).Append('|');
                sb.Append(S((object)ps.FitToPagesTall)).Append('|');
                sb.Append(S((object)ps.LeftMargin)).Append('|');
                sb.Append(S((object)ps.RightMargin)).Append('|');
                sb.Append(S((object)ps.TopMargin)).Append('|');
                sb.Append(S((object)ps.BottomMargin)).Append('|');
                sb.Append(S((object)ps.PrintArea));
                s.Page = sb.ToString();
            }
            catch { s.Page = "ERR"; }
        }

        static string Cell(object o)
        {
            if (o == null) return "";
            if (o is double) return ((double)o).ToString("R");
            return Convert.ToString(o);
        }

        public static List<string> Compare(List<SheetSnap> a, List<SheetSnap> b)
        {
            List<string> d = new List<string>();
            if (TestForceDiff) d.Add("（テスト用の強制差異）");
            if (a.Count != b.Count) { d.Add("シート数: " + a.Count + " → " + b.Count); return d; }
            for (int i = 0; i < a.Count; i++)
            {
                SheetSnap x = a[i], y = b[i];
                string pre = "[" + x.Name + "] ";
                if (x.Name != y.Name) d.Add("シート名/順序: " + x.Name + " → " + y.Name);
                if (x.Type != y.Type) d.Add(pre + "シート種別: " + x.Type + " → " + y.Type);
                if (x.Visible != y.Visible) d.Add(pre + "表示状態: " + x.Visible + " → " + y.Visible);
                if (!x.Detailed) continue;
                if (x.Used != y.Used) d.Add(pre + "使用範囲: " + x.Used + " → " + y.Used);
                if (x.Merge != y.Merge) d.Add(pre + "結合セル: " + x.Merge + " → " + y.Merge);
                CompareList(d, pre + "行高", x.RowH, y.RowH, 0.1);
                CompareList(d, pre + "列幅", x.ColW, y.ColW, 0.1);
                if (x.Shapes.Count != y.Shapes.Count) d.Add(pre + "図形数: " + x.Shapes.Count + " → " + y.Shapes.Count);
                else
                {
                    for (int k = 0; k < x.Shapes.Count; k++)
                    {
                        ShapeSnap p = x.Shapes[k], q = y.Shapes[k];
                        string sp = pre + "図形「" + p.Name + "」";
                        if (p.Name != q.Name) d.Add(sp + " 名前: → " + q.Name);
                        if (p.Type != q.Type) d.Add(sp + " 種類: " + p.Type + " → " + q.Type);
                        if (Math.Abs(p.Top - q.Top) > 0.5 || Math.Abs(p.Left - q.Left) > 0.5)
                            d.Add(sp + " 位置: (" + p.Left.ToString("0.0") + "," + p.Top.ToString("0.0") + ") → (" + q.Left.ToString("0.0") + "," + q.Top.ToString("0.0") + ")");
                        if (Math.Abs(p.Width - q.Width) > 0.5 || Math.Abs(p.Height - q.Height) > 0.5)
                            d.Add(sp + " サイズ: " + p.Width.ToString("0.0") + "x" + p.Height.ToString("0.0") + " → " + q.Width.ToString("0.0") + "x" + q.Height.ToString("0.0"));
                        if (Math.Abs(p.Rotation - q.Rotation) > 0.5) d.Add(sp + " 回転: " + p.Rotation + " → " + q.Rotation);
                        if (p.Placement != q.Placement) d.Add(sp + " 配置: " + p.Placement + " → " + q.Placement);
                        if (p.Visible != q.Visible) d.Add(sp + " 表示: " + p.Visible + " → " + q.Visible);
                    }
                }
                if (x.Page != y.Page) d.Add(pre + "印刷設定: " + x.Page + " → " + y.Page);
                if (x.Sample != y.Sample) d.Add(pre + "セルの値（先頭部分）が異なります");
            }
            return d;
        }

        static void CompareList(List<string> d, string label, List<double> a, List<double> b, double tol)
        {
            if (a.Count != b.Count) { d.Add(label + " の数: " + a.Count + " → " + b.Count); return; }
            int n = 0; string first = null;
            for (int i = 0; i < a.Count; i++)
            {
                if (Math.Abs(a[i] - b[i]) > tol)
                {
                    n++;
                    if (first == null) first = "#" + (i + 1) + " " + a[i].ToString("0.00") + " → " + b[i].ToString("0.00");
                }
            }
            if (n > 0) d.Add(label + ": " + n + " か所 (" + first + ")");
        }
    }
}
