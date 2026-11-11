using System;
using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Web.Script.Serialization;

namespace XlsImageShrink
{
    public class ManEntry
    {
        public string Kind = "convert";     // convert / compressOnly
        public string Original = "";
        public string Backup = "";
        public string Output = "";
        public string OriginalSha256 = "";
        public string OutputSha256 = "";
        public string Created = "";
        public string Modified = "";
        public string State = "";           // kept / deleted / moving / moved / restored / failed
        public string Note = "";
    }

    public class Manifest
    {
        public string JobId = "";
        public string Tool = AppPaths.Version;
        public string Time = "";
        public string SourceRoot = "";
        public string BackupRoot = "";
        public string Mode = "";
        public List<ManEntry> Entries = new List<ManEntry>();
        [ScriptIgnore] public string FilePath;

        public void Save()
        {
            JavaScriptSerializer ser = new JavaScriptSerializer();
            ser.MaxJsonLength = int.MaxValue;
            string json = ser.Serialize(this);
            if (FilePath != null) Util.AtomicWrite(FilePath, json);
            if (!string.IsNullOrEmpty(BackupRoot) && Directory.Exists(BackupRoot))
            {
                try { Util.AtomicWrite(Path.Combine(BackupRoot, "_manifest.json"), json); } catch { }
            }
        }

        public static Manifest Load(string path)
        {
            JavaScriptSerializer ser = new JavaScriptSerializer();
            ser.MaxJsonLength = int.MaxValue;
            Manifest m = ser.Deserialize<Manifest>(File.ReadAllText(path, Encoding.UTF8));
            if (m == null) throw new InvalidDataException("マニフェストを読めません");
            if (m.Entries == null) m.Entries = new List<ManEntry>();
            m.FilePath = path;
            return m;
        }

        public static List<Manifest> LoadAll(List<string> errors)
        {
            List<Manifest> list = new List<Manifest>();
            if (!Directory.Exists(AppPaths.HistoryDir)) return list;
            string[] files = Directory.GetFiles(AppPaths.HistoryDir, "*.manifest.json");
            Array.Sort(files); Array.Reverse(files);
            foreach (string f in files)
            {
                try { list.Add(Load(f)); }
                catch (Exception ex) { if (errors != null) errors.Add(Path.GetFileName(f) + ": " + ex.Message); }
            }
            return list;
        }

        /// <summary>異常終了で「移動中」のまま残ったエントリを、実体のある場所に合わせて直す。</summary>
        public int Repair()
        {
            int n = 0;
            foreach (ManEntry e in Entries)
            {
                if (e.State != "moving") continue;
                bool atBackup = File.Exists(e.Backup), atOrig = File.Exists(e.Original);
                if (e.Kind == "compressOnly")
                {
                    e.State = atBackup ? "moved" : "kept";
                }
                else if (atBackup && !atOrig) e.State = "moved";
                else if (atBackup && atOrig) { try { Util.ForceDelete(e.Backup); } catch { } e.State = "kept"; }
                else e.State = "kept";
                n++;
            }
            if (n > 0) Save();
            return n;
        }
    }

    public static class OriginalHandler
    {
        /// <summary>元ファイルを移動先へ退避する。失敗時は例外。元ファイルは残る。</summary>
        public static void Move(string original, string backup, string expectedHash)
        {
            Directory.CreateDirectory(Path.GetDirectoryName(backup));
            if (File.Exists(backup)) throw new IOException("移動先に同名ファイルがあります: " + backup);
            if (Util.SameVolume(original, backup))
            {
                File.Move(original, backup);
            }
            else
            {
                try
                {
                    File.Copy(original, backup, false);
                    if (Util.Sha256(backup) != expectedHash) throw new IOException("移動先のコピーが元と一致しません");
                }
                catch
                {
                    try { if (File.Exists(backup)) Util.ForceDelete(backup); } catch { }
                    throw;
                }
                Util.ForceDelete(original);
            }
        }

        public static void Delete(string original, bool recycle)
        {
            if (recycle) Util.Recycle(original); else Util.ForceDelete(original);
            if (File.Exists(original)) throw new IOException("削除できませんでした");
        }
    }

    public class RestoreResult
    {
        public int Restored, Skipped, Failed;
        public List<string> Messages = new List<string>();
    }

    public static class RestoreService
    {
        public static RestoreResult Restore(Manifest m, IEnumerable<ManEntry> entries, bool deleteGenerated, bool allowModified, Logger log)
        {
            RestoreResult res = new RestoreResult();
            foreach (ManEntry e in entries)
            {
                string name = Path.GetFileName(e.Original);
                try
                {
                    if (e.State == "restored") { res.Skipped++; res.Messages.Add(name + ": 復元済み"); continue; }
                    if (e.State != "moved") { res.Skipped++; continue; }
                    if (!File.Exists(e.Backup)) { res.Failed++; res.Messages.Add(name + ": 移動先にファイルがありません"); continue; }
                    if (!allowModified && Util.Sha256(e.Backup) != e.OriginalSha256)
                    {
                        res.Failed++; res.Messages.Add(name + ": 移動先のファイルが変更されています（復元していません）"); continue;
                    }
                    Directory.CreateDirectory(Path.GetDirectoryName(e.Original));
                    string dest = e.Original;
                    if (e.Kind == "compressOnly")
                    {
                        // 元の場所には同名の圧縮版がある
                        if (File.Exists(e.Original))
                        {
                            if (Util.Sha256(e.Original) == e.OutputSha256) Util.Recycle(e.Original);
                            else
                            {
                                dest = Util.UniquePath(Path.GetDirectoryName(e.Original), Path.GetFileNameWithoutExtension(e.Original) + "_復元", Path.GetExtension(e.Original));
                                res.Messages.Add(name + ": 圧縮版が編集されているため、元ファイルを " + Path.GetFileName(dest) + " として戻しました");
                            }
                        }
                    }
                    else
                    {
                        if (File.Exists(e.Original)) { res.Skipped++; res.Messages.Add(name + ": 元の場所に同名ファイルがあるためスキップ"); continue; }
                    }
                    MoveBack(e.Backup, dest);
                    try
                    {
                        if (!string.IsNullOrEmpty(e.Created)) File.SetCreationTimeUtc(dest, DateTime.Parse(e.Created, null, System.Globalization.DateTimeStyles.RoundtripKind));
                        if (!string.IsNullOrEmpty(e.Modified)) File.SetLastWriteTimeUtc(dest, DateTime.Parse(e.Modified, null, System.Globalization.DateTimeStyles.RoundtripKind));
                    }
                    catch { }
                    if (e.Kind != "compressOnly" && deleteGenerated && !string.IsNullOrEmpty(e.Output) && File.Exists(e.Output))
                    {
                        if (Util.Sha256(e.Output) == e.OutputSha256) Util.Recycle(e.Output);
                        else res.Messages.Add(name + ": 生成ファイルが編集されているため削除していません (" + Path.GetFileName(e.Output) + ")");
                    }
                    e.State = "restored";
                    res.Restored++;
                    if (log != null) log.Info("復元: " + dest);
                }
                catch (Exception ex)
                {
                    res.Failed++; res.Messages.Add(name + ": " + ex.Message);
                    if (log != null) log.Error("復元失敗: " + e.Original + " " + ex.Message);
                }
                try { m.Save(); } catch { }
            }
            // 全件復元できたら移動先の空フォルダを片付ける
            bool all = true;
            foreach (ManEntry e in m.Entries) if (e.State == "moved" || e.State == "moving") all = false;
            if (all && !string.IsNullOrEmpty(m.BackupRoot) && Directory.Exists(m.BackupRoot)) RemoveEmptyDirs(m.BackupRoot);
            return res;
        }

        static void MoveBack(string backup, string dest)
        {
            if (Util.SameVolume(backup, dest)) { File.Move(backup, dest); return; }
            File.Copy(backup, dest, false);
            if (Util.Sha256(backup) != Util.Sha256(dest)) { Util.ForceDelete(dest); throw new IOException("復元先のコピーが一致しません"); }
            Util.ForceDelete(backup);
        }

        static void RemoveEmptyDirs(string dir)
        {
            try
            {
                foreach (string d in Directory.GetDirectories(dir)) RemoveEmptyDirs(d);
                if (Directory.GetFiles(dir).Length == 0 && Directory.GetDirectories(dir).Length == 0) Directory.Delete(dir);
            }
            catch { }
        }
    }
}
