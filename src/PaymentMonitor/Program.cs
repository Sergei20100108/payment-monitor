using System;
using System.IO;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;

namespace PaymentMonitor
{
    internal static class Program
    {
        private const int AttachParentProcess = -1;

        [DllImport("kernel32.dll", SetLastError = true)]
        private static extern bool AttachConsole(int dwProcessId);

        [STAThread]
        private static int Main(string[] args)
        {
            // Режим консоли: PaymentMonitor.exe --import ... --list ... (см. README)
            if (args != null && args.Length > 0)
            {
                AttachConsole(AttachParentProcess);
                return Cli.Run(args);
            }

            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            // чтобы падение не было «молчаливым» — пишем трассировку в logs\crash.txt
            Application.ThreadException += delegate (object s, System.Threading.ThreadExceptionEventArgs e)
            {
                LogCrash(e.Exception);
                MessageBox.Show("Непредвиденная ошибка:\r\n\r\n" + e.Exception.Message +
                    "\r\n\r\nПодробности: logs\\crash.txt", "PaymentMonitor",
                    MessageBoxButtons.OK, MessageBoxIcon.Error);
            };
            AppDomain.CurrentDomain.UnhandledException += delegate (object s, UnhandledExceptionEventArgs e)
            {
                LogCrash(e.ExceptionObject as Exception);
            };

            try
            {
                Application.Run(new MainForm());
            }
            catch (Exception ex)
            {
                LogCrash(ex);
                MessageBox.Show("Не удалось запустить приложение:\r\n\r\n" + ex.Message,
                    "PaymentMonitor", MessageBoxButtons.OK, MessageBoxIcon.Error);
                return 4;
            }
            return 0;
        }

        private static void LogCrash(Exception ex)
        {
            try
            {
                string dir = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "logs");
                Directory.CreateDirectory(dir);
                File.AppendAllText(Path.Combine(dir, "crash.txt"),
                    "==== " + DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss") + Environment.NewLine +
                    (ex == null ? "(нет данных)" : ex.ToString()) + Environment.NewLine + Environment.NewLine,
                    Encoding.UTF8);
            }
            catch { }
        }
    }
}
