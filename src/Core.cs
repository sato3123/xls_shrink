using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Security.Cryptography;
using System.Web.Script.Serialization;

namespace XlsImageShrink
{
    public class Settings
    {
        public string TargetFolder = "";
        public bool Recursive = true;
        public string ExcludeFolders = "";
        public bool IncludeConverted = false;
        public int Ppi = 150;
        public int MaxLongSide = 2000;
        public int JpegQuality = 85;
        public bool PngToJpeg = true;
        public int MinPpi = 96;
        public int MinQuality = 50;
        public bool SizeTargetEnabled = false;
        public double SizeLimitMB = 5.0;
        public bool SkipIfUnderLimit = true;
        public string UnachievedMode = "best";      // best / convertOnly
        public string OriginalMode = "keep";        // keep / delete / move
        public bool DeleteToRecycle = true;
        public string MoveRoot = "";
        public bool KeepOnLayoutDiff = true;
        public string Conflict = "number";          // number / skip / overwrite
        public string CompressSuffix = "_圧縮";
        public bool VerifyLayout = true;
        public bool KeepTimestamp = false;
        public bool EmptyVbaAsXlsx = false;
        public int TimeoutSec = 300;
        public int ExcelRestartEvery = 50;
        public int LogKeepDays = 30;

        public static Settings Load(string path)
        {
            try
            {
                if (File.Exists(path))
                {
                    Settings s = new JavaScriptSerializer().Deserialize<Settings>(File.ReadAllText(path, Encoding.UTF8));
                    if (s != null) return s;
                }
            }
            catch
            {
                try { File.Copy(path, path + ".bak", true); } catch { }
            }
            return new Settings();
        }

        public void Save(string path)
        {
            try { Util.AtomicWrite(path, new JavaScriptSerializer().Serialize(this)); } catch { }
        }

        public Settings Clone() { return (Settings)MemberwiseClone(); }

        public string Validate()
        {
            if (Ppi < 50 || Ppi > 600) return "解像度は 50〜600 ppi で指定してください。";
            if (MinPpi < 50 || MinPpi > 600) return "最低 ppi は 50〜600 で指定してください。";
            if (MinPpi > Ppi) return "最低 ppi は解像度以下にしてください。";
            if (JpegQuality < 30 || JpegQuality > 100) return "JPEG 品質は 30〜100 で指定してください。";
            if (MinQuality < 30 || MinQuality > 100) return "最低品質は 30〜100 で指定してください。";
            if (MinQuality > JpegQuality) return "最低品質は JPEG 品質以下にしてください。";
            if (MaxLongSide != 0 && (MaxLongSide < 320 || MaxLongSide > 10000)) return "長辺上限は 0（無制限）または 320〜10000 px で指定してください。";
            if (SizeTargetEnabled && (SizeLimitMB < 0.1 || SizeLimitMB > 1000)) return "サイズ上限は 0.1〜1000 MB で指定してください。";
            if (OriginalMode == "move" && string.IsNullOrEmpty(MoveRoot)) return "移動先フォルダを指定してください。";
            if (string.IsNullOrEmpty(CompressSuffix) || CompressSuffix.Length > 20 || CompressSuffix.IndexOfAny(Path.GetInvalidFileNameChars()) >= 0)
                return "圧縮のみの接尾辞は、ファイル名に使える 1〜20 文字で指定してください。";
            return null;
        }
    }

    public static class AppPaths
    {
        public static string BaseDir;
        public const string Version = "XlsImageShrink 1.0.0";

        public static void Init(string overrideDir)
        {
            if (overrideDir != null) { BaseDir = overrideDir; }
            else
            {
                string exeDir = AppDomain.CurrentDomain.BaseDirectory;
                try
                {
                    string t = Path.Combine(exeDir, ".wtest");
                    File.WriteAllText(t, "x"); File.Delete(t);
                    BaseDir = exeDir;
                }
                catch
                {
                    BaseDir = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData), "XlsImageShrink");
                }
            }
            Directory.CreateDirectory(BaseDir);
            Directory.CreateDirectory(LogDir);
            Directory.CreateDirectory(ReportDir);
            Directory.CreateDirectory(HistoryDir);
        }
        public static string SettingsFile { get { return Path.Combine(BaseDir, "settings.json"); } }
        public static string LogDir { get { return Path.Combine(BaseDir, "logs"); } }
        public static string ReportDir { get { return Path.Combine(BaseDir, "reports"); } }
        public static string HistoryDir { get { return Path.Combine(BaseDir, "history"); } }
        public static string TempRoot { get { return Path.Combine(Path.GetTempPath(), "XlsImageShrink"); } }
    }

    public class Logger
    {
        readonly object gate = new object();
        public string FilePath;
        public event Action<string> Line;

        public Logger(string filePath) { FilePath = filePath; }

        public void Info(string msg) { Write("INFO ", msg); }
        public void Warn(string msg) { Write("WARN ", msg); }
        public void Error(string msg) { Write("ERROR", msg); }

        void Write(string level, string msg)
        {
            string line = DateTime.Now.ToString("HH:mm:ss") + " " + level + " " + msg;
            lock (gate)
            {
                try { if (FilePath != null) File.AppendAllText(FilePath, line + "\r\n", new UTF8Encoding(true)); } catch { }
            }
            Action<string> h = Line;
            if (h != null) { try { h(line); } catch { } }
        }
    }

    public enum ItemKind { Convert, CompressOnly }

    public class FileItem
    {
        public string Path;
        public long Size;
        public ItemKind Kind;
        public bool Checked = true;
        public string PreNote = "";
        public string Status = "待機";
        public string ErrorCode = "";
        public string Message = "";
        public string Macro = "";
        public string OutputPath = "";
        public long ConvSize;
        public long FinalSize;
        public double UsedPpi;
        public int UsedQuality;
        public int ImgTotal, ImgCompressed, ImgSkipped;
        public int DiffCount;
        public string OrigAction = "";
        public double Seconds;
        public List<string> Warnings = new List<string>();
        public object Tag;
        public string KindText { get { return Kind == ItemKind.Convert ? "変換+圧縮" : "圧縮のみ"; } }
    }

    public static class Util
    {
        public static bool NoRecycle = false;   // 自己テスト用: ごみ箱を使わず完全削除

        public static string Sha256(string path)
        {
            using (SHA256 sha = SHA256.Create())
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.Read))
            {
                byte[] h = sha.ComputeHash(fs);
                StringBuilder sb = new StringBuilder();
                for (int i = 0; i < h.Length; i++) sb.Append(h[i].ToString("x2"));
                return sb.ToString();
            }
        }

        public static string FormatSize(long bytes)
        {
            if (bytes >= 1024L * 1024 * 1024) return (bytes / (1024.0 * 1024 * 1024)).ToString("0.00") + "GB";
            if (bytes >= 1024 * 1024) return (bytes / (1024.0 * 1024)).ToString("0.0") + "MB";
            if (bytes >= 1024) return (bytes / 1024.0).ToString("0") + "KB";
            return bytes + "B";
        }

        public static void AtomicWrite(string path, string text)
        {
            string tmp = path + ".tmp";
            File.WriteAllText(tmp, text, new UTF8Encoding(false));
            if (File.Exists(path)) File.Delete(path);
            File.Move(tmp, path);
        }

        public static bool IsNetworkPath(string path)
        {
            try
            {
                string full = Path.GetFullPath(path);
                if (full.StartsWith(@"\\")) return true;
                DriveInfo di = new DriveInfo(Path.GetPathRoot(full));
                return di.DriveType == DriveType.Network;
            }
            catch { return false; }
        }

        public static void Recycle(string path)
        {
            if (NoRecycle) { ForceDelete(path); return; }
            Microsoft.VisualBasic.FileIO.FileSystem.DeleteFile(path,
                Microsoft.VisualBasic.FileIO.UIOption.OnlyErrorDialogs,
                Microsoft.VisualBasic.FileIO.RecycleOption.SendToRecycleBin);
        }

        public static void ForceDelete(string path)
        {
            FileAttributes a = File.GetAttributes(path);
            if ((a & FileAttributes.ReadOnly) != 0) File.SetAttributes(path, a & ~FileAttributes.ReadOnly);
            File.Delete(path);
        }

        public static string UniquePath(string dir, string baseName, string ext)
        {
            string p = Path.Combine(dir, baseName + ext);
            int n = 1;
            while (File.Exists(p)) { p = Path.Combine(dir, baseName + "(" + n + ")" + ext); n++; }
            return p;
        }

        public static bool IsLocked(string path)
        {
            try
            {
                using (new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.None)) { }
                return false;
            }
            catch (IOException) { return true; }
            catch (UnauthorizedAccessException) { return true; }
        }

        public static bool SameVolume(string a, string b)
        {
            try { return string.Equals(Path.GetPathRoot(Path.GetFullPath(a)), Path.GetPathRoot(Path.GetFullPath(b)), StringComparison.OrdinalIgnoreCase); }
            catch { return false; }
        }

        public static string RelativeTo(string root, string path)
        {
            string r = Path.GetFullPath(root).TrimEnd('\\') + "\\";
            string p = Path.GetFullPath(path);
            if (p.StartsWith(r, StringComparison.OrdinalIgnoreCase)) return p.Substring(r.Length);
            return Path.GetFileName(p);
        }

        public static void TryDeleteDir(string dir)
        {
            try { if (Directory.Exists(dir)) Directory.Delete(dir, true); } catch { }
        }

        public static string Csv(string s)
        {
            if (s == null) s = "";
            s = s.Replace("\r", " ").Replace("\n", " ");
            if (s.IndexOfAny(new char[] { ',', '"', '\r', '\n' }) >= 0) return "\"" + s.Replace("\"", "\"\"") + "\"";
            return s;
        }
    }

    public enum FileFormat { Ole, Zip, Markup, Empty, Other }

    public class ProbeResult
    {
        public FileFormat Format;
        public bool VbaStorage;
        public bool EncryptedPackage;
    }

    public static class FormatProbe
    {
        public static ProbeResult Probe(string path)
        {
            ProbeResult r = new ProbeResult();
            byte[] head = new byte[8];
            int n;
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                n = fs.Read(head, 0, 8);
            }
            if (n == 0) { r.Format = FileFormat.Empty; return r; }
            if (n >= 8 && head[0] == 0xD0 && head[1] == 0xCF && head[2] == 0x11 && head[3] == 0xE0 && head[4] == 0xA1 && head[5] == 0xB1 && head[6] == 0x1A && head[7] == 0xE1)
            {
                r.Format = FileFormat.Ole;
                r.VbaStorage = Contains(path, Encoding.Unicode.GetBytes("_VBA_PROJECT_CUR"));
                r.EncryptedPackage = Contains(path, Encoding.Unicode.GetBytes("EncryptedPackage"));
                return r;
            }
            if (n >= 4 && head[0] == 0x50 && head[1] == 0x4B && head[2] == 0x03 && head[3] == 0x04) { r.Format = FileFormat.Zip; return r; }
            int i = 0;
            if (n >= 3 && head[0] == 0xEF && head[1] == 0xBB && head[2] == 0xBF) i = 3;
            while (i < n && (head[i] == 0x20 || head[i] == 0x0D || head[i] == 0x0A || head[i] == 0x09)) i++;
            if (i < n && head[i] == (byte)'<') { r.Format = FileFormat.Markup; return r; }
            r.Format = FileFormat.Other;
            return r;
        }

        static bool Contains(string path, byte[] pat)
        {
            const int Buf = 1 << 16;
            byte[] buf = new byte[Buf + pat.Length];
            int keep = 0;
            using (FileStream fs = new FileStream(path, FileMode.Open, FileAccess.Read, FileShare.ReadWrite))
            {
                while (true)
                {
                    int n = fs.Read(buf, keep, Buf);
                    if (n <= 0) return false;
                    int len = keep + n;
                    for (int i = 0; i + pat.Length <= len; i++)
                    {
                        if (buf[i] != pat[0]) continue;
                        int j = 1;
                        while (j < pat.Length && buf[i + j] == pat[j]) j++;
                        if (j == pat.Length) return true;
                    }
                    keep = Math.Min(pat.Length - 1, len);
                    Buffer.BlockCopy(buf, len - keep, buf, 0, keep);
                }
            }
        }
    }

    public static class FileScanner
    {
        public static List<FileItem> Scan(Settings s, Logger log, Func<bool> cancel)
        {
            List<FileItem> list = new List<FileItem>();
            HashSet<string> excl = new HashSet<string>(StringComparer.OrdinalIgnoreCase);
            foreach (string e in (s.ExcludeFolders ?? "").Split(new char[] { ';', ',' }, StringSplitOptions.RemoveEmptyEntries)) excl.Add(e.Trim());
            string moveRoot = null;
            if (!string.IsNullOrEmpty(s.MoveRoot)) { try { moveRoot = Path.GetFullPath(s.MoveRoot).TrimEnd('\\'); } catch { } }
            Walk(Path.GetFullPath(s.TargetFolder), s, excl, moveRoot, list, log, cancel, true);
            return list;
        }

        // 「名前_圧縮」「名前_圧縮(1)」のような、このツールが出力したファイルか
        static bool IsOwnOutput(string stem, string suffix)
        {
            if (string.IsNullOrEmpty(suffix)) return false;
            if (stem.EndsWith(suffix, StringComparison.OrdinalIgnoreCase)) return true;
            return System.Text.RegularExpressions.Regex.IsMatch(stem, System.Text.RegularExpressions.Regex.Escape(suffix) + @"\(\d+\)$", System.Text.RegularExpressions.RegexOptions.IgnoreCase);
        }

        static void Walk(string dir, Settings s, HashSet<string> excl, string moveRoot, List<FileItem> list, Logger log, Func<bool> cancel, bool top)
        {
            if (cancel != null && cancel()) return;
            string[] files;
            try { files = Directory.GetFiles(dir); }
            catch (Exception ex) { if (log != null) log.Warn("フォルダを読めません: " + dir + " (" + ex.Message + ")"); return; }
            Array.Sort(files, StringComparer.OrdinalIgnoreCase);
            foreach (string f in files)
            {
                string name = Path.GetFileName(f);
                if (name.StartsWith("~$")) continue;
                string ext = Path.GetExtension(f).ToLowerInvariant();
                ItemKind kind;
                if (ext == ".xls") kind = ItemKind.Convert;
                else if (s.IncludeConverted && (ext == ".xlsx" || ext == ".xlsm"))
                {
                    if (IsOwnOutput(Path.GetFileNameWithoutExtension(f), s.CompressSuffix)) continue;
                    kind = ItemKind.CompressOnly;
                }
                else continue;
                try
                {
                    FileInfo fi = new FileInfo(f);
                    if ((fi.Attributes & (FileAttributes.Hidden | FileAttributes.System)) != 0) continue;
                    FileItem it = new FileItem();
                    it.Path = f; it.Size = fi.Length; it.Kind = kind;
                    try
                    {
                        ProbeResult pr = FormatProbe.Probe(f);
                        if (kind == ItemKind.Convert)
                        {
                            if (pr.Format != FileFormat.Ole) { it.PreNote = "形式不一致"; it.Checked = false; }
                            else if (pr.VbaStorage) it.PreNote = "マクロあり";
                        }
                        else
                        {
                            if (pr.Format == FileFormat.Ole && pr.EncryptedPackage) { it.PreNote = "暗号化"; it.Checked = false; }
                            else if (pr.Format != FileFormat.Zip) { it.PreNote = "形式不一致"; it.Checked = false; }
                        }
                    }
                    catch { it.PreNote = "読取不可"; }
                    list.Add(it);
                }
                catch { }
            }
            if (!s.Recursive) return;
            string[] dirs;
            try { dirs = Directory.GetDirectories(dir); } catch { return; }
            Array.Sort(dirs, StringComparer.OrdinalIgnoreCase);
            foreach (string d in dirs)
            {
                try
                {
                    DirectoryInfo di = new DirectoryInfo(d);
                    if ((di.Attributes & (FileAttributes.Hidden | FileAttributes.System | FileAttributes.ReparsePoint)) != 0) continue;
                    if (excl.Contains(di.Name)) continue;
                    if (moveRoot != null && string.Equals(di.FullName.TrimEnd('\\'), moveRoot, StringComparison.OrdinalIgnoreCase)) continue;
                }
                catch { continue; }
                Walk(d, s, excl, moveRoot, list, log, cancel, false);
            }
        }
    }
}
