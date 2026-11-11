using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Windows.Forms;

namespace XlsImageShrink
{
    public class HistoryForm : Form
    {
        ListView lvJobs, lvFiles;
        CheckBox chkDeleteGenerated;
        Button btnRestoreSel, btnRestoreAll, btnOpen, btnClose;
        List<Manifest> jobs = new List<Manifest>();

        public HistoryForm()
        {
            Text = "履歴・復元";
            Font = new Font("Meiryo UI", 9f);
            AutoScaleMode = AutoScaleMode.Dpi;
            ClientSize = new Size(980, 600);
            StartPosition = FormStartPosition.CenterParent;

            lvJobs = new ListView(); lvJobs.View = View.Details; lvJobs.FullRowSelect = true; lvJobs.HideSelection = false; lvJobs.MultiSelect = false;
            lvJobs.Dock = DockStyle.Top; lvJobs.Height = 200;
            lvJobs.Columns.Add("日時", 150); lvJobs.Columns.Add("対象フォルダ", 380); lvJobs.Columns.Add("件数", 60, HorizontalAlignment.Right);
            lvJobs.Columns.Add("元ファイル", 80); lvJobs.Columns.Add("復元状態", 160);
            lvJobs.SelectedIndexChanged += delegate { ShowFiles(); };

            lvFiles = new ListView(); lvFiles.View = View.Details; lvFiles.FullRowSelect = true; lvFiles.HideSelection = false; lvFiles.Dock = DockStyle.Fill;
            lvFiles.Columns.Add("元のパス", 330); lvFiles.Columns.Add("種別", 70); lvFiles.Columns.Add("移動先", 260); lvFiles.Columns.Add("出力", 180); lvFiles.Columns.Add("状態", 90);

            Panel p = new Panel(); p.Dock = DockStyle.Bottom; p.Height = 44;
            chkDeleteGenerated = new CheckBox(); chkDeleteGenerated.Text = "生成した xlsx/xlsm を削除する（ごみ箱へ）"; chkDeleteGenerated.Checked = true; chkDeleteGenerated.AutoSize = true; chkDeleteGenerated.Location = new Point(8, 12); p.Controls.Add(chkDeleteGenerated);
            btnRestoreSel = B(p, "選択を復元", 330, 110); btnRestoreAll = B(p, "ジョブ全体を復元", 446, 130); btnOpen = B(p, "マニフェストを開く...", 582, 150); btnClose = B(p, "閉じる", 860, 100);
            btnRestoreSel.Click += delegate { DoRestore(false); };
            btnRestoreAll.Click += delegate { DoRestore(true); };
            btnOpen.Click += delegate { OpenManifest(); };
            btnClose.Click += delegate { Close(); };

            Controls.Add(lvFiles); Controls.Add(p); Controls.Add(lvJobs);
            LoadJobs();
        }

        Button B(Control parent, string text, int x, int w)
        {
            Button b = new Button(); b.Text = text; b.Location = new Point(x, 8); b.Size = new Size(w, 28); parent.Controls.Add(b); return b;
        }

        void LoadJobs()
        {
            List<string> errors = new List<string>();
            jobs = Manifest.LoadAll(errors);
            FillJobs();
            if (errors.Count > 0)
                MessageBox.Show(this, "読み込めない履歴があります:\r\n" + string.Join("\r\n", errors.ToArray()) + "\r\n\r\n移動先フォルダの _manifest.json を「マニフェストを開く」で読み込めます。", Text, MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        void FillJobs()
        {
            lvJobs.Items.Clear();
            foreach (Manifest m in jobs)
            {
                int moved = 0, restored = 0;
                foreach (ManEntry e in m.Entries) { if (e.State == "moved") moved++; if (e.State == "restored") restored++; }
                string st = m.Mode != "move" ? "（復元対象なし）" : moved == 0 && restored > 0 ? "復元済み" : restored > 0 ? "一部復元 (" + restored + "/" + (moved + restored) + ")" : moved > 0 ? "未復元 (" + moved + " 件)" : "（移動なし）";
                ListViewItem li = new ListViewItem(new string[] { m.Time, m.SourceRoot, m.Entries.Count.ToString(), m.Mode == "move" ? "移動" : m.Mode == "delete" ? "削除" : "残す", st });
                li.Tag = m; lvJobs.Items.Add(li);
            }
            ShowFiles();
        }

        Manifest Current() { return lvJobs.SelectedItems.Count > 0 ? lvJobs.SelectedItems[0].Tag as Manifest : null; }

        void ShowFiles()
        {
            lvFiles.Items.Clear();
            Manifest m = Current();
            bool can = false;
            if (m != null)
            {
                foreach (ManEntry e in m.Entries)
                {
                    string st = e.State == "moved" ? "移動済み" : e.State == "restored" ? "復元済み" : e.State == "deleted" ? "削除済み" : e.State == "kept" ? "残した" : e.State;
                    ListViewItem li = new ListViewItem(new string[] { e.Original, e.Kind == "compressOnly" ? "圧縮のみ" : "変換", e.Backup, Path.GetFileName(e.Output), st });
                    li.Tag = e; lvFiles.Items.Add(li);
                    if (e.State == "moved") can = true;
                }
            }
            btnRestoreSel.Enabled = can; btnRestoreAll.Enabled = can;
        }

        void DoRestore(bool all)
        {
            Manifest m = Current(); if (m == null) return;
            List<ManEntry> target = new List<ManEntry>();
            if (all) target.AddRange(m.Entries);
            else foreach (ListViewItem li in lvFiles.SelectedItems) target.Add((ManEntry)li.Tag);
            if (target.Count == 0) { MessageBox.Show(this, "復元するファイルを選択してください。", Text); return; }
            if (MessageBox.Show(this, target.Count + " 件を元の場所に戻します。よろしいですか？", "確認", MessageBoxButtons.YesNo, MessageBoxIcon.Question) != DialogResult.Yes) return;
            Cursor = Cursors.WaitCursor;
            RestoreResult r = RestoreService.Restore(m, target, chkDeleteGenerated.Checked, false, null);
            Cursor = Cursors.Default;
            string msg = "復元 " + r.Restored + " 件 / スキップ " + r.Skipped + " 件 / 失敗 " + r.Failed + " 件";
            if (r.Messages.Count > 0) msg += "\r\n\r\n" + string.Join("\r\n", r.Messages.GetRange(0, Math.Min(20, r.Messages.Count)).ToArray());
            MessageBox.Show(this, msg, Text, MessageBoxButtons.OK, r.Failed > 0 ? MessageBoxIcon.Warning : MessageBoxIcon.Information);
            int idx = lvJobs.SelectedIndices.Count > 0 ? lvJobs.SelectedIndices[0] : -1;
            FillJobs();
            if (idx >= 0 && idx < lvJobs.Items.Count) lvJobs.Items[idx].Selected = true;
        }

        void OpenManifest()
        {
            using (OpenFileDialog d = new OpenFileDialog())
            {
                d.Filter = "マニフェスト (*.json)|*.json";
                if (d.ShowDialog(this) != DialogResult.OK) return;
                try
                {
                    Manifest m = Manifest.Load(d.FileName);
                    jobs.Insert(0, m); FillJobs(); lvJobs.Items[0].Selected = true;
                }
                catch (Exception ex) { MessageBox.Show(this, "読み込めません: " + ex.Message, Text, MessageBoxButtons.OK, MessageBoxIcon.Error); }
            }
        }
    }
}
