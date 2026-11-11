using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.IO;
using System.Text;
using System.Threading;

namespace XlsImageShrink
{
    public class JobRunner
    {
        public Settings S;
        public Logger Log;
        public string JobId;
        public Manifest Man;
        public bool DryRun;
        public string ReportPath;
        public volatile bool StopRequested;
        public ManualResetEvent PauseGate = new ManualResetEvent(true);
        public event Action<FileItem> ItemChanged;
        public event Action<int, int> Progress;

        string workRoot;
        ExcelSession xl;

        public JobRunner(Settings s)
        {
            S = s.Clone();
            JobId = DateTime.Now.ToString("yyyyMMdd_HHmmss") + "_" + Guid.NewGuid().ToString("N").Substring(0, 4);
            Log = new Logger(Path.Combine(AppPaths.LogDir, JobId + ".log"));
        }

        void Notify(FileItem it) { Action<FileItem> h = ItemChanged; if (h != null) { try { h(it); } catch { } } }

        public void Run(List<FileItem> items)
        {
            workRoot = Path.Combine(AppPaths.TempRoot, JobId);
            Directory.CreateDirectory(workRoot);
            Man = new Manifest();
            Man.JobId = JobId; Man.Time = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
            Man.SourceRoot = S.TargetFolder; Man.Mode = S.OriginalMode;
            if (S.OriginalMode == "move") Man.BackupRoot = Path.Combine(S.MoveRoot, JobId.Substring(0, 15));
            Man.FilePath = Path.Combine(AppPaths.HistoryDir, JobId + ".manifest.json");
            Log.Info("ジョブ開始 " + JobId + " 対象=" + S.TargetFolder + " 元ファイル=" + S.OriginalMode + (DryRun ? " (ドライラン)" : ""));
            xl = new ExcelSession(Log, S.TimeoutSec, S.ExcelRestartEvery);
            List<FileItem> todo = new List<FileItem>();
            foreach (FileItem it in items) if (it.Checked) todo.Add(it);
            int done = 0;
            try
            {
                foreach (FileItem it in todo)
                {
                    PauseGate.WaitOne();
                    if (StopRequested) break;
                    Stopwatch sw = Stopwatch.StartNew();
                    string wd = Path.Combine(workRoot, done.ToString());
                    try
                    {
                        Directory.CreateDirectory(wd);
                        it.Status = "処理中"; it.ErrorCode = ""; it.Message = ""; it.Warnings.Clear();
                        Notify(it);
                        Process(it, wd);
                    }
                    catch (Exception ex)
                    {
                        it.Status = "失敗"; it.ErrorCode = "E999"; it.Message = ex.Message;
                        Log.Error(it.Path + " : " + ex);
                        if (!xl.IsAlive()) xl.Abandon();
                    }
                    Util.TryDeleteDir(wd);
                    it.Seconds = sw.Elapsed.TotalSeconds;
                    Log.Info(it.Path + " → " + it.Status + (it.Message != "" ? " (" + it.Message + ")" : "")
                        + (it.FinalSize > 0 ? " " + Util.FormatSize(it.Size) + "→" + Util.FormatSize(it.FinalSize) + " 画像 " + it.ImgCompressed + "/" + it.ImgTotal + "枚" : ""));
                    Notify(it);
                    done++;
                    Action<int, int> p = Progress; if (p != null) { try { p(done, todo.Count); } catch { } }
                }
            }
            finally
            {
                try { xl.Dispose(); } catch { }
                Util.TryDeleteDir(workRoot);
                try { if (!DryRun && Man.Entries.Count > 0) Man.Save(); } catch (Exception ex) { Log.Error("履歴の保存に失敗: " + ex.Message); }
                try { WriteReport(todo); } catch (Exception ex) { Log.Error("レポートの保存に失敗: " + ex.Message); }
                Log.Info("ジョブ終了 " + done + "/" + todo.Count + " 件");
            }
        }

        void Skip(FileItem it, string code, string msg) { it.Status = "スキップ"; it.ErrorCode = code; it.Message = msg; }
        void Fail(FileItem it, string code, string msg) { it.Status = "失敗"; it.ErrorCode = code; it.Message = msg; }

        void Process(FileItem it, string wd)
        {
            if (!File.Exists(it.Path)) { Skip(it, "E106", "ファイルが見つかりません"); return; }
            if (Util.IsLocked(it.Path)) { Skip(it, "E101", "使用中"); return; }
            string ext = Path.GetExtension(it.Path).ToLowerInvariant();
            string inCopy = Path.Combine(wd, "in" + ext);
            FileInfo fi = new FileInfo(it.Path);
            long srcLen = fi.Length; DateTime srcWrite = fi.LastWriteTimeUtc, srcCreate = fi.CreationTimeUtc;
            it.Size = srcLen;
            File.Copy(it.Path, inCopy, true);
            File.SetAttributes(inCopy, FileAttributes.Normal);
            string srcHash = Util.Sha256(inCopy);
            ProbeResult pr = FormatProbe.Probe(inCopy);
            bool convert = it.Kind == ItemKind.Convert;
            if (convert)
            {
                if (pr.Format != FileFormat.Ole) { Skip(it, "E105", "実体が xls 形式ではありません"); return; }
            }
            else
            {
                if (pr.Format == FileFormat.Ole && pr.EncryptedPackage) { Skip(it, "E102", "暗号化"); return; }
                if (pr.Format != FileFormat.Zip) { Skip(it, "E105", "実体が xlsx/xlsm 形式ではありません"); return; }
                try { ImageOptimizer.CheckZip(inCopy); } catch (Exception ex) { Fail(it, "E104", "破損: " + ex.Message); return; }
            }

            // --- Excel で開く（変換、または比較用のレイアウト採取）
            string conv = inCopy;
            string outExt = ext;
            List<SheetSnap> before = null;
            if (convert || S.VerifyLayout)
            {
                dynamic wb = null;
                try { wb = xl.Open(inCopy); }
                catch (Exception ex)
                {
                    bool timeout = xl.TimedOut;
                    if (!xl.IsAlive()) xl.Abandon();
                    string m = ex.Message ?? "";
                    if (timeout) Fail(it, "E201", "Excel が応答しません（タイムアウト）");
                    else if (m.IndexOf("パスワード") >= 0 || m.ToLowerInvariant().IndexOf("password") >= 0) Skip(it, "E102", "暗号化（パスワード付き）");
                    else Fail(it, "E104", "開けません: " + m);
                    return;
                }
                try
                {
                    if (convert)
                    {
                        it.Macro = MacroDetector.Detect(wb, S, Log, it.Warnings);
                        if (pr.VbaStorage && it.Macro.IndexOf("VBA") < 0 && !S.EmptyVbaAsXlsx)
                        {
                            it.Warnings.Add("事前判定では VBA あり、Excel 判定ではなし。安全側で xlsm とします");
                            it.Macro = it.Macro == "" ? "VBA?" : it.Macro + "+VBA?";
                        }
                    }
                    if (S.VerifyLayout) before = LayoutSnapshot.Take(wb);
                    if (convert)
                    {
                        outExt = it.Macro != "" ? ".xlsm" : ".xlsx";
                        conv = Path.Combine(wd, "conv" + outExt);
                        xl.SaveAs(wb, conv, it.Macro != "" ? 52 : 51);
                    }
                }
                catch (Exception ex)
                {
                    bool timeout = xl.TimedOut;
                    if (!xl.IsAlive()) { xl.Abandon(); wb = null; }
                    Fail(it, timeout ? "E201" : "E104", timeout ? "Excel が応答しません（タイムアウト）" : "変換に失敗: " + ex.Message);
                    return;
                }
                finally { if (wb != null) xl.CloseWorkbook(wb); }
            }
            if (!convert) { try { it.Macro = ImageOptimizer.HasVbaProject(inCopy) ? "VBA" : ""; } catch { } }
            it.ConvSize = new FileInfo(conv).Length;

            // --- 画像圧縮
            SizeResult sr;
            try { sr = SizeController.Optimize(conv, wd, S, Log); }
            catch (Exception ex)
            {
                Log.Warn("画像圧縮に失敗したため圧縮なしで続行: " + ex.Message);
                it.Warnings.Add("画像圧縮に失敗（圧縮なし）: " + ex.Message);
                sr = new SizeResult(); sr.Path = conv; sr.Size = it.ConvSize; sr.Stats = new ImgStats();
            }

            // --- 検証
            List<string> diffs = new List<string>();
            for (int attempt = 0; ; attempt++)
            {
                try
                {
                    ImageOptimizer.CheckZip(sr.Path);
                    if (S.VerifyLayout && before != null)
                    {
                        dynamic wb2 = xl.Open(sr.Path);
                        try { diffs = LayoutSnapshot.Compare(before, LayoutSnapshot.Take(wb2)); }
                        finally { xl.CloseWorkbook(wb2); }
                    }
                    break;
                }
                catch (Exception ex)
                {
                    if (!xl.IsAlive()) xl.Abandon();
                    if (attempt == 0 && sr.Changed)
                    {
                        Log.Warn("圧縮後のファイルを検証できないため、圧縮なしでやり直します: " + ex.Message);
                        it.Warnings.Add("圧縮後のファイルを開けないため圧縮なしで出力");
                        SizeResult n = new SizeResult(); n.Path = conv; n.Size = it.ConvSize; n.Stats = new ImgStats(); sr = n;
                        continue;
                    }
                    Fail(it, "E402", "出力を検証できません: " + ex.Message);
                    return;
                }
            }
            it.FinalSize = sr.Size; it.UsedPpi = sr.Ppi; it.UsedQuality = sr.Quality;
            it.ImgTotal = sr.Stats.Total; it.ImgCompressed = sr.Stats.Compressed; it.ImgSkipped = sr.Stats.Skipped;
            foreach (string n in sr.Stats.Notes) it.Warnings.Add(n);
            it.DiffCount = diffs.Count;
            foreach (string d in diffs) { Log.Warn("  レイアウト差異: " + d); }
            if (diffs.Count > 0) it.Warnings.Add("レイアウト差異 " + diffs.Count + " 件: " + diffs[0]);
            if (sr.Unachieved) it.Warnings.Add("サイズ上限に届きません");

            if (!convert && !sr.Changed)
            {
                it.Status = sr.Unachieved ? "変更なし（目標未達）" : "変更なし"; it.FinalSize = it.Size; it.OrigAction = "そのまま";
                return;
            }
            if (DryRun) { it.Status = "ドライラン"; it.Message = "予測 " + Util.FormatSize(it.FinalSize); return; }

            // --- 出力先と元ファイルの扱いを決める
            bool hasDiff = diffs.Count > 0;
            string mode = S.OriginalMode;
            bool keepOrig = mode == "keep" || (hasDiff && S.KeepOnLayoutDiff);
            if (!keepOrig && mode == "delete" && S.DeleteToRecycle && Util.IsNetworkPath(it.Path))
            {
                keepOrig = true; it.Warnings.Add("ネットワーク上のファイルはごみ箱に入らないため、元ファイルを残しました");
            }
            string dir = Path.GetDirectoryName(it.Path), stem = Path.GetFileNameWithoutExtension(it.Path);
            bool replace = !convert && !keepOrig;      // 圧縮のみで削除/移動: 元と同じ名前で置き換える
            string outPath;
            if (replace) outPath = it.Path;
            else
            {
                outPath = Path.Combine(dir, (convert ? stem : stem + S.CompressSuffix) + outExt);
                if (File.Exists(outPath))
                {
                    if (S.Conflict == "skip") { Skip(it, "E108", "同名の出力ファイルがあります"); return; }
                    if (S.Conflict == "number") outPath = Util.UniquePath(dir, convert ? stem : stem + S.CompressSuffix, outExt);
                    else if (Util.IsLocked(outPath)) { Fail(it, "E101", "上書き先が使用中です"); return; }
                }
            }
            // 処理中に元ファイルが変更されていないか
            FileInfo now = new FileInfo(it.Path);
            if (!now.Exists || now.Length != srcLen || now.LastWriteTimeUtc != srcWrite) { Fail(it, "E107", "処理中に元ファイルが変更されました"); return; }

            string tmp = outPath + ".~tmp";
            try
            {
                File.Copy(sr.Path, tmp, true);
                if (new FileInfo(tmp).Length != sr.Size) throw new IOException("コピー後のサイズが一致しません");
            }
            catch (Exception ex)
            {
                try { if (File.Exists(tmp)) File.Delete(tmp); } catch { }
                Fail(it, ex is UnauthorizedAccessException ? "E502" : "E501", "出力できません: " + ex.Message);
                return;
            }
            string outHash = Util.Sha256(tmp);

            ManEntry me = new ManEntry();
            me.Kind = convert ? "convert" : "compressOnly";
            me.Original = it.Path; me.OriginalSha256 = srcHash; me.OutputSha256 = outHash;
            me.Created = srcCreate.ToString("o"); me.Modified = srcWrite.ToString("o");
            me.State = "kept";
            bool origFailed = false;

            if (replace)
            {
                try { DisposeOriginal(it, me, srcHash, mode); }
                catch (Exception ex)
                {
                    origFailed = true;
                    it.Warnings.Add("元ファイルを" + (mode == "move" ? "移動" : "削除") + "できないため残しました: " + ex.Message);
                    outPath = Util.UniquePath(dir, stem + S.CompressSuffix, outExt);
                    me.State = "kept"; me.Backup = "";
                }
                File.Move(tmp, outPath);
                me.Output = outPath;
                if (me.State == "moving") me.State = "moved";
            }
            else
            {
                if (File.Exists(outPath)) Util.ForceDelete(outPath);
                File.Move(tmp, outPath);
                me.Output = outPath;
                if (!keepOrig)
                {
                    try { DisposeOriginal(it, me, srcHash, mode); if (me.State == "moving") me.State = "moved"; }
                    catch (Exception ex)
                    {
                        origFailed = true; me.State = "kept"; me.Backup = "";
                        it.Warnings.Add("元ファイルを" + (mode == "move" ? "移動" : "削除") + "できないため残しました: " + ex.Message);
                    }
                }
            }
            if (S.KeepTimestamp) { try { File.SetLastWriteTimeUtc(outPath, srcWrite); } catch { } }
            Man.Entries.Add(me);
            try { Man.Save(); } catch { }

            it.OutputPath = outPath;
            it.OrigAction = me.State == "moved" ? "移動" : me.State == "deleted" ? "削除" : "残す";
            if (origFailed) { it.Status = "完了（要確認）"; it.ErrorCode = "E601"; }
            else if (hasDiff) { it.Status = "完了（要確認）"; it.ErrorCode = "E401"; }
            else if (sr.Unachieved) it.Status = "完了（目標未達）";
            else it.Status = "完了";
            if (it.Warnings.Count > 0) it.Message = it.Warnings[0];
        }

        void DisposeOriginal(FileItem it, ManEntry me, string srcHash, string mode)
        {
            if (mode == "move")
            {
                string backup = Path.Combine(Man.BackupRoot, Util.RelativeTo(S.TargetFolder, it.Path));
                me.Backup = backup; me.State = "moving";
                Directory.CreateDirectory(Man.BackupRoot);
                Man.Entries.Add(me);
                try { Man.Save(); } finally { Man.Entries.Remove(me); }
                OriginalHandler.Move(it.Path, backup, srcHash);
            }
            else
            {
                OriginalHandler.Delete(it.Path, S.DeleteToRecycle);
                me.State = "deleted";
            }
        }

        void WriteReport(List<FileItem> items)
        {
            ReportPath = Path.Combine(AppPaths.ReportDir, JobId + (DryRun ? "_dryrun" : "") + ".csv");
            StringBuilder sb = new StringBuilder();
            sb.AppendLine("元パス,処理,結果,エラーコード,出力パス,マクロ,元サイズ,変換直後サイズ,最終サイズ,削減率%,画像数,圧縮,スキップ,採用ppi,採用品質,レイアウト差異数,元ファイル処理,処理秒,メッセージ");
            foreach (FileItem it in items)
            {
                string rate = (it.FinalSize > 0 && it.Size > 0) ? (100.0 * (it.Size - it.FinalSize) / it.Size).ToString("0.0") : "";
                sb.AppendLine(string.Join(",", new string[] {
                    Util.Csv(it.Path), it.KindText, Util.Csv(it.Status), it.ErrorCode, Util.Csv(it.OutputPath), Util.Csv(it.Macro),
                    it.Size.ToString(), it.ConvSize.ToString(), it.FinalSize.ToString(), rate,
                    it.ImgTotal.ToString(), it.ImgCompressed.ToString(), it.ImgSkipped.ToString(),
                    it.UsedPpi > 0 ? it.UsedPpi.ToString("0") : "", it.UsedQuality > 0 ? it.UsedQuality.ToString() : "",
                    it.DiffCount.ToString(), it.OrigAction, it.Seconds.ToString("0.0"), Util.Csv(string.Join(" / ", it.Warnings.ToArray()) + (it.Warnings.Count == 0 ? it.Message : "")) }));
            }
            File.WriteAllText(ReportPath, sb.ToString(), new UTF8Encoding(true));
        }
    }
}
