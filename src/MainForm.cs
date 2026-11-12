using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Threading;
using System.Windows.Forms;

namespace XlsImageShrink
{
    public class MainForm : Form
    {
        Settings S;
        TextBox txtFolder, txtExclude, txtMoveRoot, txtSuffix, txtLog;
        CheckBox chkRecursive, chkIncludeConv, chkPngJpeg, chkSize, chkSkipUnder, chkRecycle, chkKeepDiff, chkVerify, chkTimestamp, chkEmptyVba;
        ComboBox cmbPreset, cmbConflict, cmbUnachieved;
        NumericUpDown numPpi, numLong, numQ, numMinPpi, numMinQ, numLimit;
        RadioButton rbKeep, rbDelete, rbMove;
        Button btnBrowse, btnScan, btnRun, btnPause, btnStop, btnDry, btnHistory, btnLogs, btnMoveBrowse, btnRetry;
        ListView lv;
        ProgressBar bar;
        Label lblProgress, lblSummary, lblHint;
        List<FileItem> items = new List<FileItem>();
        JobRunner runner;
        Thread worker;
        bool paused, loading, excelOk, scanning, rescanPending, runPending, runPendingDry;
        string scannedKey;
        DateTime jobStart;

        public MainForm()
        {
            Text = "XlsImageShrink - Excel 画像圧縮・形式変換";
            Font = new Font("Meiryo UI", 9f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(1180, 760);
            MinimumSize = new Size(1024, 700);
            StartPosition = FormStartPosition.CenterScreen;
            AllowDrop = true;
            DragEnter += delegate(object s, DragEventArgs e) { if (e.Data.GetDataPresent(DataFormats.FileDrop)) e.Effect = DragDropEffects.Copy; };
            DragDrop += OnDrop;
            S = Settings.Load(AppPaths.SettingsFile);
            Build();
            loading = true; ToUi(); loading = false;
            excelOk = ExcelSession.IsInstalled();
            if (!excelOk) AppendLog("ERROR E001: Excel が見つかりません。.xls の変換とレイアウト検証はできません（変換済みファイルの画像圧縮はできます）。");
            UpdateEnable();
            CleanupOld();
            FormClosing += OnClosing;
            // 前回のフォルダが残っていれば、起動時に対象ファイルを検索しておく
            Shown += delegate { if (Directory.Exists(txtFolder.Text.Trim())) RequestScan(false); else SetHint("対象フォルダを選ぶと、対象ファイルを自動で検索します。"); };
        }

        // ---------- 画面構築 ----------
        Label L(Control parent, string text, int x, int y)
        {
            Label l = new Label(); l.Text = text; l.AutoSize = true; l.Location = new Point(x, y + 3); parent.Controls.Add(l); return l;
        }
        NumericUpDown N(Control parent, int x, int y, decimal min, decimal max, int dec)
        {
            NumericUpDown n = new NumericUpDown(); n.Minimum = min; n.Maximum = max; n.DecimalPlaces = dec; n.Location = new Point(x, y); n.Width = 70;
            if (dec > 0) n.Increment = 0.5m;
            n.ValueChanged += delegate { OnSettingChanged(); };
            parent.Controls.Add(n); return n;
        }
        CheckBox C(Control parent, string text, int x, int y)
        {
            CheckBox c = new CheckBox(); c.Text = text; c.AutoSize = true; c.Location = new Point(x, y);
            c.CheckedChanged += delegate { OnSettingChanged(); UpdateEnable(); };
            parent.Controls.Add(c); return c;
        }
        Label H(Control parent, string text, int y)
        {
            Label l = new Label(); l.Text = "■ " + text; l.AutoSize = true; l.Font = new Font(Font, FontStyle.Bold); l.Location = new Point(6, y); parent.Controls.Add(l); return l;
        }

        void Build()
        {
            Panel top = new Panel(); top.Dock = DockStyle.Top; top.Height = 92;
            L(top, "対象フォルダ", 8, 8);
            txtFolder = new TextBox(); txtFolder.Location = new Point(100, 8); txtFolder.Width = 880; txtFolder.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right; top.Controls.Add(txtFolder);
            btnBrowse = new Button(); btnBrowse.Text = "参照..."; btnBrowse.Location = new Point(990, 6); btnBrowse.Width = 80; btnBrowse.Anchor = AnchorStyles.Top | AnchorStyles.Right; top.Controls.Add(btnBrowse);
            btnBrowse.Click += delegate { string p = PickFolder(txtFolder.Text); if (p != null) { txtFolder.Text = p; RequestScan(false); } };
            txtFolder.KeyDown += delegate(object s, KeyEventArgs e) { if (e.KeyCode == Keys.Enter) { e.SuppressKeyPress = true; RequestScan(false); } };
            txtFolder.Leave += delegate { RescanIfChanged(); };
            btnScan = new Button(); btnScan.Text = "再検索"; btnScan.Location = new Point(1080, 6); btnScan.Width = 90; btnScan.Anchor = AnchorStyles.Top | AnchorStyles.Right; top.Controls.Add(btnScan);
            btnScan.Click += delegate { RequestScan(false); };
            chkRecursive = C(top, "サブフォルダも対象", 100, 36);
            L(top, "除外フォルダ名（; 区切り）", 260, 36);
            txtExclude = new TextBox(); txtExclude.Location = new Point(430, 36); txtExclude.Width = 260; top.Controls.Add(txtExclude);
            txtExclude.Leave += delegate { RescanIfChanged(); };
            chkIncludeConv = C(top, "変換済みファイル（.xlsx / .xlsm）も対象にする（画像圧縮のみ）", 100, 62);
            chkRecursive.CheckedChanged += delegate { RescanIfChanged(); };
            chkIncludeConv.CheckedChanged += delegate { RescanIfChanged(); };
            lblHint = new Label(); lblHint.AutoSize = true; lblHint.Location = new Point(720, 39); lblHint.ForeColor = Color.DimGray; top.Controls.Add(lblHint);

            Panel left = new Panel(); left.Dock = DockStyle.Left; left.Width = 330; left.AutoScroll = true;
            int y = 4;
            H(left, "画像設定", y); y += 24;
            L(left, "プリセット", 16, y); cmbPreset = new ComboBox(); cmbPreset.DropDownStyle = ComboBoxStyle.DropDownList; cmbPreset.Items.AddRange(new object[] { "標準", "高圧縮", "画質優先", "カスタム" });
            cmbPreset.Location = new Point(150, y); cmbPreset.Width = 150; left.Controls.Add(cmbPreset); cmbPreset.SelectedIndexChanged += delegate { ApplyPreset(); }; y += 28;
            L(left, "解像度 (ppi)", 16, y); numPpi = N(left, 150, y, 50, 600, 0); y += 28;
            L(left, "長辺上限 (px, 0=無制限)", 16, y); numLong = N(left, 180, y, 0, 10000, 0); y += 28;
            L(left, "JPEG 品質", 16, y); numQ = N(left, 150, y, 30, 100, 0); y += 28;
            chkPngJpeg = C(left, "透過なし PNG/BMP/TIFF を JPEG 化", 16, y); y += 26;
            L(left, "最低 ppi", 16, y); numMinPpi = N(left, 80, y, 50, 600, 0); L(left, "最低品質", 160, y); numMinQ = N(left, 230, y, 30, 100, 0); y += 34;

            H(left, "サイズ目標", y); y += 24;
            chkSize = C(left, "有効   上限 (MB)", 16, y); numLimit = N(left, 180, y, 0.1m, 1000m, 1); y += 28;
            chkSkipUnder = C(left, "上限以下なら圧縮しない", 16, y); y += 26;
            L(left, "上限に届かないとき", 16, y); cmbUnachieved = new ComboBox(); cmbUnachieved.DropDownStyle = ComboBoxStyle.DropDownList; cmbUnachieved.Items.AddRange(new object[] { "最低画質で出力", "圧縮せず出力" });
            cmbUnachieved.Location = new Point(150, y); cmbUnachieved.Width = 150; left.Controls.Add(cmbUnachieved); cmbUnachieved.SelectedIndexChanged += delegate { OnSettingChanged(); }; y += 34;

            H(left, "元ファイル", y); y += 24;
            rbKeep = new RadioButton(); rbKeep.Text = "残す"; rbKeep.AutoSize = true; rbKeep.Location = new Point(16, y); left.Controls.Add(rbKeep);
            rbDelete = new RadioButton(); rbDelete.Text = "削除"; rbDelete.AutoSize = true; rbDelete.Location = new Point(96, y); left.Controls.Add(rbDelete);
            rbMove = new RadioButton(); rbMove.Text = "移動"; rbMove.AutoSize = true; rbMove.Location = new Point(176, y); left.Controls.Add(rbMove);
            EventHandler rb = delegate { OnSettingChanged(); UpdateEnable(); };
            rbKeep.CheckedChanged += rb; rbDelete.CheckedChanged += rb; rbMove.CheckedChanged += rb; y += 26;
            chkRecycle = C(left, "削除はごみ箱へ（オフ＝完全削除）", 32, y); y += 26;
            L(left, "移動先", 32, y); txtMoveRoot = new TextBox(); txtMoveRoot.Location = new Point(84, y); txtMoveRoot.Width = 160; left.Controls.Add(txtMoveRoot);
            btnMoveBrowse = new Button(); btnMoveBrowse.Text = "参照"; btnMoveBrowse.Location = new Point(250, y - 2); btnMoveBrowse.Width = 54; left.Controls.Add(btnMoveBrowse);
            btnMoveBrowse.Click += delegate { string p = PickFolder(txtMoveRoot.Text); if (p != null) txtMoveRoot.Text = p; }; y += 30;
            chkKeepDiff = C(left, "レイアウト差異があるときは\r\n元ファイルを削除・移動しない", 16, y); y += 44;
            chkKeepDiff.Click += delegate
            {
                if (!loading && !chkKeepDiff.Checked)
                {
                    if (MessageBox.Show(this, "オフにすると、変換前後でレイアウトに差異が見つかったファイルでも、元ファイルを削除・移動します。\r\nよろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes)
                        chkKeepDiff.Checked = true;
                }
            };

            H(left, "その他", y); y += 24;
            L(left, "同名の出力があるとき", 16, y); cmbConflict = new ComboBox(); cmbConflict.DropDownStyle = ComboBoxStyle.DropDownList; cmbConflict.Items.AddRange(new object[] { "連番を付ける", "スキップ", "上書き" });
            cmbConflict.Location = new Point(170, y); cmbConflict.Width = 130; left.Controls.Add(cmbConflict); cmbConflict.SelectedIndexChanged += delegate { OnSettingChanged(); }; y += 28;
            L(left, "圧縮のみの接尾辞", 16, y); txtSuffix = new TextBox(); txtSuffix.Location = new Point(170, y); txtSuffix.Width = 130; left.Controls.Add(txtSuffix); y += 28;
            chkVerify = C(left, "レイアウト検証を行う", 16, y); y += 24;
            chkTimestamp = C(left, "出力の更新日時を元ファイルに合わせる", 16, y); y += 24;
            chkEmptyVba = C(left, "コードが空の VBA は xlsx にする", 16, y); y += 30;
            Label pad = new Label(); pad.Location = new Point(0, y); pad.Size = new Size(1, 1); left.Controls.Add(pad);

            Panel bottom = new Panel(); bottom.Dock = DockStyle.Bottom; bottom.Height = 190;
            btnRun = Btn(bottom, "実行", 8, 110); btnPause = Btn(bottom, "一時停止", 124, 90); btnStop = Btn(bottom, "中止", 220, 80);
            btnDry = Btn(bottom, "ドライラン", 306, 90); btnRetry = Btn(bottom, "失敗分を再実行", 402, 120);
            btnHistory = Btn(bottom, "履歴・復元", 0, 100); btnLogs = Btn(bottom, "ログフォルダ", 0, 100);
            btnLogs.Left = bottom.Width - 108; btnLogs.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnHistory.Left = bottom.Width - 214; btnHistory.Anchor = AnchorStyles.Top | AnchorStyles.Right;
            btnRun.Click += delegate { StartJob(false, false); };
            btnDry.Click += delegate { StartJob(true, false); };
            btnRetry.Click += delegate { StartJob(false, true); };
            btnPause.Click += delegate { TogglePause(); };
            btnStop.Click += delegate { if (runner != null) { runner.StopRequested = true; runner.PauseGate.Set(); AppendLog("中止を要求しました。処理中のファイルが終わりしだい停止します。"); } };
            btnHistory.Click += delegate { using (HistoryForm f = new HistoryForm()) f.ShowDialog(this); };
            btnLogs.Click += delegate { try { Process.Start(AppPaths.LogDir); } catch { } };
            bar = new ProgressBar(); bar.Location = new Point(8, 40); bar.Size = new Size(420, 18); bottom.Controls.Add(bar);
            lblProgress = L(bottom, "", 440, 38);
            lblSummary = L(bottom, "", 700, 38);
            txtLog = new TextBox(); txtLog.Multiline = true; txtLog.ReadOnly = true; txtLog.ScrollBars = ScrollBars.Vertical; txtLog.Location = new Point(8, 64);
            txtLog.Size = new Size(bottom.Width - 16, 120); txtLog.Anchor = AnchorStyles.Left | AnchorStyles.Top | AnchorStyles.Right | AnchorStyles.Bottom; txtLog.BackColor = SystemColors.Window; bottom.Controls.Add(txtLog);

            lv = new ListView(); lv.Dock = DockStyle.Fill; lv.View = View.Details; lv.CheckBoxes = true; lv.FullRowSelect = true; lv.GridLines = true; lv.HideSelection = false;
            lv.Columns.Add("パス", 340); lv.Columns.Add("サイズ", 70, HorizontalAlignment.Right); lv.Columns.Add("処理", 80); lv.Columns.Add("マクロ", 90);
            lv.Columns.Add("出力", 60); lv.Columns.Add("結果サイズ", 80, HorizontalAlignment.Right); lv.Columns.Add("状態", 120); lv.Columns.Add("メッセージ", 300);
            lv.ItemChecked += delegate(object s, ItemCheckedEventArgs e) { FileItem it = e.Item.Tag as FileItem; if (it != null && worker == null) it.Checked = e.Item.Checked; };
            ContextMenuStrip cm = new ContextMenuStrip();
            cm.Items.Add("フォルダを開く", null, delegate { FileItem it = Selected(); if (it != null) { try { Process.Start("explorer.exe", "/select,\"" + (File.Exists(it.OutputPath) ? it.OutputPath : it.Path) + "\""); } catch { } } });
            cm.Items.Add("すべて選択", null, delegate { foreach (ListViewItem li in lv.Items) li.Checked = true; });
            cm.Items.Add("すべて解除", null, delegate { foreach (ListViewItem li in lv.Items) li.Checked = false; });
            lv.ContextMenuStrip = cm;

            Controls.Add(lv); Controls.Add(left); Controls.Add(bottom); Controls.Add(top);
        }

        Button Btn(Control parent, string text, int x, int w)
        {
            Button b = new Button(); b.Text = text; b.Location = new Point(x, 6); b.Size = new Size(w, 28); parent.Controls.Add(b); return b;
        }

        FileItem Selected() { return lv.SelectedItems.Count > 0 ? lv.SelectedItems[0].Tag as FileItem : null; }

        string PickFolder(string start)
        {
            using (FolderBrowserDialog d = new FolderBrowserDialog())
            {
                if (Directory.Exists(start)) d.SelectedPath = start;
                return d.ShowDialog(this) == DialogResult.OK ? d.SelectedPath : null;
            }
        }

        void OnDrop(object sender, DragEventArgs e)
        {
            string[] paths = e.Data.GetData(DataFormats.FileDrop) as string[];
            if (paths == null || paths.Length == 0 || worker != null) return;
            string p = paths[0];
            if (File.Exists(p)) p = Path.GetDirectoryName(p);
            if (Directory.Exists(p)) { txtFolder.Text = p; RequestScan(false); }
        }

        // ---------- 設定 ----------
        void ToUi()
        {
            txtFolder.Text = S.TargetFolder; chkRecursive.Checked = S.Recursive; txtExclude.Text = S.ExcludeFolders; chkIncludeConv.Checked = S.IncludeConverted;
            numPpi.Value = Clamp(S.Ppi, numPpi); numLong.Value = Clamp(S.MaxLongSide, numLong); numQ.Value = Clamp(S.JpegQuality, numQ);
            chkPngJpeg.Checked = S.PngToJpeg; numMinPpi.Value = Clamp(S.MinPpi, numMinPpi); numMinQ.Value = Clamp(S.MinQuality, numMinQ);
            chkSize.Checked = S.SizeTargetEnabled; numLimit.Value = Clamp((decimal)S.SizeLimitMB, numLimit); chkSkipUnder.Checked = S.SkipIfUnderLimit;
            cmbUnachieved.SelectedIndex = S.UnachievedMode == "convertOnly" ? 1 : 0;
            rbKeep.Checked = S.OriginalMode == "keep"; rbDelete.Checked = S.OriginalMode == "delete"; rbMove.Checked = S.OriginalMode == "move";
            if (!rbKeep.Checked && !rbDelete.Checked && !rbMove.Checked) rbKeep.Checked = true;
            chkRecycle.Checked = S.DeleteToRecycle; txtMoveRoot.Text = S.MoveRoot; chkKeepDiff.Checked = S.KeepOnLayoutDiff;
            cmbConflict.SelectedIndex = S.Conflict == "skip" ? 1 : S.Conflict == "overwrite" ? 2 : 0;
            txtSuffix.Text = S.CompressSuffix; chkVerify.Checked = S.VerifyLayout; chkTimestamp.Checked = S.KeepTimestamp; chkEmptyVba.Checked = S.EmptyVbaAsXlsx;
            cmbPreset.SelectedIndex = PresetIndex();
        }
        static decimal Clamp(decimal v, NumericUpDown n) { return Math.Max(n.Minimum, Math.Min(n.Maximum, v)); }

        void FromUi()
        {
            S.TargetFolder = txtFolder.Text.Trim(); S.Recursive = chkRecursive.Checked; S.ExcludeFolders = txtExclude.Text.Trim(); S.IncludeConverted = chkIncludeConv.Checked;
            S.Ppi = (int)numPpi.Value; S.MaxLongSide = (int)numLong.Value; S.JpegQuality = (int)numQ.Value; S.PngToJpeg = chkPngJpeg.Checked;
            S.MinPpi = (int)numMinPpi.Value; S.MinQuality = (int)numMinQ.Value;
            S.SizeTargetEnabled = chkSize.Checked; S.SizeLimitMB = (double)numLimit.Value; S.SkipIfUnderLimit = chkSkipUnder.Checked;
            S.UnachievedMode = cmbUnachieved.SelectedIndex == 1 ? "convertOnly" : "best";
            S.OriginalMode = rbDelete.Checked ? "delete" : rbMove.Checked ? "move" : "keep";
            S.DeleteToRecycle = chkRecycle.Checked; S.MoveRoot = txtMoveRoot.Text.Trim(); S.KeepOnLayoutDiff = chkKeepDiff.Checked;
            S.Conflict = cmbConflict.SelectedIndex == 1 ? "skip" : cmbConflict.SelectedIndex == 2 ? "overwrite" : "number";
            S.CompressSuffix = txtSuffix.Text.Trim(); S.VerifyLayout = chkVerify.Checked; S.KeepTimestamp = chkTimestamp.Checked; S.EmptyVbaAsXlsx = chkEmptyVba.Checked;
        }

        int PresetIndex()
        {
            int p = (int)numPpi.Value, q = (int)numQ.Value;
            if (p == 150 && q == 85) return 0;
            if (p == 96 && q == 65) return 1;
            if (p == 220 && q == 90) return 2;
            return 3;
        }
        void ApplyPreset()
        {
            if (loading) return;
            loading = true;
            if (cmbPreset.SelectedIndex == 0) { numPpi.Value = 150; numQ.Value = 85; }
            else if (cmbPreset.SelectedIndex == 1) { numPpi.Value = 96; numQ.Value = 65; if (numMinPpi.Value > 96) numMinPpi.Value = 72; if (numMinQ.Value > 65) numMinQ.Value = 50; }
            else if (cmbPreset.SelectedIndex == 2) { numPpi.Value = 220; numQ.Value = 90; }
            loading = false;
        }
        void OnSettingChanged()
        {
            if (loading || cmbPreset == null || numQ == null) return;
            loading = true; cmbPreset.SelectedIndex = PresetIndex(); loading = false;
        }

        void UpdateEnable()
        {
            if (btnRun == null) return;
            bool busy = worker != null;
            bool folderOk = Directory.Exists(txtFolder.Text.Trim());
            foreach (Control c in new Control[] { txtFolder, btnBrowse, btnScan, chkRecursive, txtExclude, chkIncludeConv, cmbPreset, numPpi, numLong, numQ, chkPngJpeg, numMinPpi, numMinQ,
                chkSize, rbKeep, rbDelete, rbMove, cmbConflict, txtSuffix, chkVerify, chkTimestamp, chkEmptyVba, btnHistory })
                c.Enabled = !busy;
            numLimit.Enabled = !busy && chkSize.Checked; chkSkipUnder.Enabled = !busy && chkSize.Checked; cmbUnachieved.Enabled = !busy && chkSize.Checked;
            chkRecycle.Enabled = !busy && rbDelete.Checked;
            txtMoveRoot.Enabled = !busy && rbMove.Checked; btnMoveBrowse.Enabled = !busy && rbMove.Checked;
            chkKeepDiff.Enabled = !busy && !rbKeep.Checked && chkVerify.Checked;
            // 検索中でも押せる（検索が終わってから確認して実行する）
            btnRun.Enabled = !busy && folderOk; btnDry.Enabled = !busy && folderOk;
            btnScan.Enabled = !busy && !scanning;
            bool anyFailed = false;
            foreach (FileItem it in items) if (it.Status == "失敗" || (it.Status == "スキップ" && it.ErrorCode == "E101")) anyFailed = true;
            btnRetry.Enabled = !busy && !scanning && anyFailed;
            btnPause.Enabled = busy; btnStop.Enabled = busy;
        }

        // ---------- スキャン（フォルダを選んだら自動で行う） ----------
        void SetHint(string t) { if (lblHint != null) lblHint.Text = t; }

        string ScanKey()
        {
            return string.Join("|", new string[] { S.TargetFolder, S.Recursive.ToString(), S.ExcludeFolders ?? "", S.IncludeConverted.ToString(), S.CompressSuffix ?? "",
                S.OriginalMode == "move" ? (S.MoveRoot ?? "") : "" });
        }

        // 検索条件が前回の検索から変わっていれば検索し直す
        void RescanIfChanged()
        {
            if (loading || worker != null || txtFolder == null) return;
            FromUi();
            if (!Directory.Exists(S.TargetFolder)) { if (scannedKey != null || items.Count > 0) RequestScan(false); else UpdateEnable(); return; }
            if (ScanKey() != scannedKey) RequestScan(false);
        }

        /// <summary>対象フォルダを検索して一覧に表示する。sync=true は検索が終わるまで戻らない。</summary>
        internal void RequestScan(bool sync)
        {
            if (worker != null) return;
            if (scanning) { rescanPending = true; return; }
            FromUi();
            if (S.TargetFolder == "" || !Directory.Exists(S.TargetFolder))
            {
                ShowItems(new List<FileItem>(), null, null);
                SetHint(S.TargetFolder == "" ? "対象フォルダを選ぶと、対象ファイルを自動で検索します。" : "フォルダが見つかりません: " + S.TargetFolder);
                return;
            }
            if (S.OriginalMode == "move" && S.MoveRoot != "")
            {
                try
                {
                    string a = Path.GetFullPath(S.TargetFolder).TrimEnd('\\') + "\\", b = Path.GetFullPath(S.MoveRoot).TrimEnd('\\') + "\\";
                    if (b.StartsWith(a, StringComparison.OrdinalIgnoreCase)) AppendLog("WARN 移動先が対象フォルダの内側にあります。移動先フォルダは検索対象から除外します。");
                }
                catch { }
            }
            S.Save(AppPaths.SettingsFile);
            Settings snap = S.Clone();
            string key = ScanKey();
            scanning = true;
            SetHint("検索中…  " + snap.TargetFolder);
            lblSummary.Text = "";
            UpdateEnable();
            if (sync)
            {
                List<FileItem> found = null; string error = null;
                try { found = FileScanner.Scan(snap, null, null); } catch (Exception ex) { error = ex.Message; }
                ScanDone(found, key, error, snap);
                return;
            }
            Thread t = new Thread(delegate()
            {
                List<FileItem> found = null; string error = null;
                try { found = FileScanner.Scan(snap, null, null); } catch (Exception ex) { error = ex.Message; }
                UiPost(delegate { ScanDone(found, key, error, snap); });
            });
            t.IsBackground = true;
            t.Start();
        }

        void ScanDone(List<FileItem> found, string key, string error, Settings snap)
        {
            scanning = false;
            if (rescanPending) { rescanPending = false; UpdateEnable(); RequestScan(false); return; }
            if (error != null)
            {
                ShowItems(new List<FileItem>(), null, snap);
                SetHint("検索に失敗しました: " + error);
                AppendLog("ERROR 検索に失敗しました: " + error);
                return;
            }
            ShowItems(found, key, snap);
            if (runPending)
            {
                runPending = false;
                int n = 0; foreach (FileItem it in items) if (it.Checked) n++;
                if (n == 0) { MessageBox.Show(this, "処理するファイルがありません。", Text); return; }
                if (MessageBox.Show(this, "対象ファイルを検索しました。" + n + " 件を処理します。実行しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) == DialogResult.Yes)
                    StartJob(runPendingDry, false);
            }
        }

        void ShowItems(List<FileItem> found, string key, Settings snap)
        {
            items = found;
            scannedKey = key;
            // 出力の競合を事前に表示
            foreach (FileItem it in items)
            {
                if (it.Kind != ItemKind.Convert || it.PreNote == "形式不一致") continue;
                string stem = Path.Combine(Path.GetDirectoryName(it.Path), Path.GetFileNameWithoutExtension(it.Path));
                if (File.Exists(stem + ".xlsx") || File.Exists(stem + ".xlsm")) it.Message = "同名の出力あり（" + cmbConflict.Text + "）";
            }
            lv.BeginUpdate(); lv.Items.Clear();
            long total = 0; int nConv = 0, nCmp = 0, nOn = 0;
            foreach (FileItem it in items)
            {
                ListViewItem li = new ListViewItem(snap != null ? Util.RelativeTo(snap.TargetFolder, it.Path) : it.Path);
                for (int i = 0; i < 7; i++) li.SubItems.Add("");
                if (it.PreNote != "" && it.PreNote != "マクロあり") it.Status = "対象外（" + it.PreNote + "）";
                if (it.Kind == ItemKind.Convert && !excelOk && it.Checked) { it.Checked = false; it.Status = "対象外（Excel なし）"; }
                li.Tag = it; it.Tag = li; li.Checked = it.Checked;
                lv.Items.Add(li); Refresh(it); total += it.Size;
                if (it.Kind == ItemKind.Convert) nConv++; else nCmp++;
                if (it.Checked) nOn++;
            }
            lv.EndUpdate();
            if (snap != null)
            {
                lblSummary.Text = items.Count + " 件 / 合計 " + Util.FormatSize(total);
                string detail = "xls " + nConv + " 件" + (snap.IncludeConverted ? "、変換済み " + nCmp + " 件" : "");
                SetHint(items.Count == 0 ? "対象ファイルが見つかりません（" + (snap.Recursive ? "サブフォルダを含む" : "サブフォルダを含まない") + "）"
                    : items.Count + " 件見つかりました（" + detail + "）。処理対象 " + nOn + " 件。「実行」で開始します。");
                AppendLog("検索完了: " + snap.TargetFolder + "  " + items.Count + " 件 / " + Util.FormatSize(total) + "（" + detail + "）");
            }
            else lblSummary.Text = "";
            bar.Value = 0; lblProgress.Text = "";
            UpdateEnable();
        }

        // ----- テスト用 -----
        internal void TestSetFolder(string p) { txtFolder.Text = p; RequestScan(true); }
        internal int TestListCount() { return lv.Items.Count; }
        internal int TestCheckedCount() { int n = 0; foreach (ListViewItem li in lv.Items) if (li.Checked) n++; return n; }
        internal bool TestRunEnabled() { return btnRun.Enabled; }
        internal string TestHint() { return lblHint.Text; }
        internal void TestSetIncludeConverted(bool v) { loading = true; chkIncludeConv.Checked = v; loading = false; }
        internal void TestSetRecursive(bool v) { loading = true; chkRecursive.Checked = v; loading = false; }

        void Refresh(FileItem it)
        {
            ListViewItem li = it.Tag as ListViewItem; if (li == null) return;
            li.SubItems[1].Text = Util.FormatSize(it.Size);
            li.SubItems[2].Text = it.KindText;
            li.SubItems[3].Text = it.Macro != "" ? it.Macro : (it.PreNote == "マクロあり" ? "あり" : "");
            li.SubItems[4].Text = it.OutputPath != "" ? Path.GetExtension(it.OutputPath).TrimStart('.') : "";
            li.SubItems[5].Text = it.FinalSize > 0 ? Util.FormatSize(it.FinalSize) : "";
            li.SubItems[6].Text = it.Status;
            li.SubItems[7].Text = it.Message;
            li.ForeColor = it.Status.StartsWith("失敗") ? Color.Firebrick : it.Status.IndexOf("要確認") >= 0 || it.Status.IndexOf("未達") >= 0 ? Color.DarkOrange : it.Status.StartsWith("完了") ? Color.DarkGreen : SystemColors.WindowText;
        }

        // ---------- 実行 ----------
        void StartJob(bool dry, bool retryOnly)
        {
            if (worker != null) return;
            if (scanning) { runPending = true; runPendingDry = dry; SetHint("検索が終わりしだい、確認して実行します…"); return; }
            FromUi();
            string err = S.Validate();
            if (err != null) { MessageBox.Show(this, err, Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            if (!retryOnly && ScanKey() != scannedKey)
            {
                // フォルダや検索条件が変わっていたら検索し直してから確認する
                RequestScan(true);
                int n = 0; foreach (FileItem it in items) if (it.Checked) n++;
                if (n == 0) { MessageBox.Show(this, "処理するファイルがありません。", Text); return; }
                if (MessageBox.Show(this, "対象ファイルを検索しました。" + n + " 件を処理します。実行しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            }
            List<FileItem> todo = new List<FileItem>();
            foreach (FileItem it in items)
            {
                bool want = retryOnly ? (it.Status == "失敗" || (it.Status == "スキップ" && it.ErrorCode == "E101")) : it.Checked;
                it.Checked = want;
                if (want) todo.Add(it);
            }
            if (todo.Count == 0) { MessageBox.Show(this, "処理するファイルがありません。一覧でチェックを付けてください。", Text); return; }
            bool noExcelVerify = false;
            if (!excelOk)
            {
                int conv = 0;
                foreach (FileItem it in todo) if (it.Kind == ItemKind.Convert) { it.Checked = false; conv++; }
                if (conv > 0) todo.RemoveAll(delegate(FileItem it) { return it.Kind == ItemKind.Convert; });
                if (todo.Count == 0) { MessageBox.Show(this, "Excel が見つからないため、.xls の変換はできません。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
                if (S.VerifyLayout) noExcelVerify = true;
            }
            if (!dry && S.OriginalMode == "delete" && !S.DeleteToRecycle)
            {
                if (MessageBox.Show(this, "元ファイル " + todo.Count + " 件を完全に削除します（ごみ箱に入らず、元に戻せません）。\r\n実行しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            }
            if (!dry && S.Conflict == "overwrite")
            {
                if (MessageBox.Show(this, "同名の出力ファイルがある場合は上書きします。実行しますか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) return;
            }
            try
            {
                long max = 0; foreach (FileItem it in todo) if (it.Size > max) max = it.Size;
                DriveInfo di = new DriveInfo(Path.GetPathRoot(Path.GetTempPath()));
                if (di.AvailableFreeSpace < max * 3) { MessageBox.Show(this, "作業用の空き容量が足りません（最大ファイルの 3 倍が必要です）。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning); return; }
            }
            catch { }
            S.Save(AppPaths.SettingsFile);
            runner = new JobRunner(S); runner.DryRun = dry;
            if (noExcelVerify) { runner.S.VerifyLayout = false; AppendLog("WARN Excel がないため、レイアウト検証をせずに実行します。"); }
            runner.Log.Line += delegate(string line) { UiPost(delegate { AppendLog(line); }); };
            runner.ItemChanged += delegate(FileItem it) { UiPost(delegate { Refresh(it); }); };
            runner.Progress += delegate(int done, int total) { UiPost(delegate { ShowProgress(done, total); }); };
            bar.Maximum = todo.Count; bar.Value = 0; jobStart = DateTime.Now; paused = false; btnPause.Text = "一時停止";
            JobRunner r = runner; List<FileItem> all = items;
            worker = new Thread(delegate()
            {
                try { r.Run(all); }
                catch (Exception ex) { r.Log.Error("ジョブが異常終了: " + ex); }
                UiPost(delegate { JobFinished(); });
            });
            worker.SetApartmentState(ApartmentState.STA);
            worker.IsBackground = true;
            worker.Start();
            UpdateEnable();
        }

        void UiPost(MethodInvoker a) { try { if (!IsDisposed && IsHandleCreated) BeginInvoke(a); } catch { } }

        void ShowProgress(int done, int total)
        {
            bar.Value = Math.Min(bar.Maximum, done);
            double sec = (DateTime.Now - jobStart).TotalSeconds;
            string rest = done > 0 && done < total ? "  残り約 " + Math.Ceiling(sec / done * (total - done) / 60.0) + " 分" : "";
            long a = 0, b = 0;
            foreach (FileItem it in items) if (it.FinalSize > 0) { a += it.Size; b += it.FinalSize; }
            lblProgress.Text = done + " / " + total + rest;
            lblSummary.Text = "削減 " + Util.FormatSize(a) + " → " + Util.FormatSize(b);
        }

        void JobFinished()
        {
            string rep = runner != null ? runner.ReportPath : null;
            worker = null; runner = null;
            foreach (FileItem it in items) { ListViewItem li = it.Tag as ListViewItem; if (li != null) li.Checked = it.Checked; }
            UpdateEnable();
            if (rep != null) AppendLog("レポート: " + rep);
        }

        void TogglePause()
        {
            if (runner == null) return;
            paused = !paused;
            if (paused) { runner.PauseGate.Reset(); btnPause.Text = "再開"; AppendLog("一時停止します（処理中のファイルが終わりしだい）。"); }
            else { runner.PauseGate.Set(); btnPause.Text = "一時停止"; AppendLog("再開します。"); }
        }

        void AppendLog(string line)
        {
            if (txtLog.TextLength > 200000) txtLog.Text = txtLog.Text.Substring(100000);
            txtLog.AppendText(line + "\r\n");
        }

        void OnClosing(object sender, FormClosingEventArgs e)
        {
            if (worker != null)
            {
                if (MessageBox.Show(this, "処理中です。中止して終了しますか？\r\n（処理中のファイルが終わるまで待ちます。元ファイルは壊れません）", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Warning) != DialogResult.Yes) { e.Cancel = true; return; }
                runner.StopRequested = true; runner.PauseGate.Set();
                Thread w = worker;
                Cursor = Cursors.WaitCursor;
                while (w.IsAlive) { Application.DoEvents(); Thread.Sleep(50); }
            }
            try { FromUi(); S.Save(AppPaths.SettingsFile); } catch { }
        }

        // 古いログ・残った作業フォルダの掃除、移動中のまま残った履歴の修復
        void CleanupOld()
        {
            try
            {
                foreach (string f in Directory.GetFiles(AppPaths.LogDir, "*.log"))
                    if (File.GetLastWriteTime(f) < DateTime.Now.AddDays(-Math.Max(1, S.LogKeepDays))) File.Delete(f);
            }
            catch { }
            try
            {
                if (Directory.Exists(AppPaths.TempRoot))
                    foreach (string d in Directory.GetDirectories(AppPaths.TempRoot))
                        if (Directory.GetLastWriteTime(d) < DateTime.Now.AddHours(-1)) Util.TryDeleteDir(d);
            }
            catch { }
            try
            {
                foreach (Manifest m in Manifest.LoadAll(null)) { int n = m.Repair(); if (n > 0) AppendLog("履歴 " + m.JobId + " の中断された移動 " + n + " 件を修復しました。"); }
            }
            catch { }
        }
    }
}
