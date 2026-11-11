using System;
using System.IO;
using System.Windows.Forms;

namespace XlsImageShrink
{
    static class Program
    {
        [STAThread]
        static int Main(string[] args)
        {
            if (args.Length >= 2 && args[0] == "--citest")
            {
                try { return CiTest.Run(args[1]); }
                catch (Exception ex)
                {
                    try { File.AppendAllText(Path.Combine(args[1], "citest_report.txt"), "\r\nFATAL " + ex + "\r\n"); } catch { }
                    return 2;
                }
            }
            if (args.Length >= 2 && args[0] == "--selftest")
            {
                try { return SelfTest.Run(args[1]); }
                catch (Exception ex)
                {
                    try { File.AppendAllText(Path.Combine(args[1], "selftest_report.txt"), "\r\nFATAL " + ex + "\r\n"); } catch { }
                    return 2;
                }
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            AppPaths.Init(null);
            Application.ThreadException += delegate(object s, System.Threading.ThreadExceptionEventArgs e)
            {
                MessageBox.Show("エラーが発生しました:\r\n" + e.Exception.Message, "XlsImageShrink", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            Application.Run(new MainForm());
            return 0;
        }
    }
}
