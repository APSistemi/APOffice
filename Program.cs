using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace APOffice
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);

            Application.ThreadException += (sender, e) =>
            {
                try
                {
                    new clsFuncs().ErrorLog("ThreadException", e.Exception.Message + "\r\n" + e.Exception.StackTrace);
                }
                catch { }
                MessageBox.Show("Si è verificato un errore:\r\n" + e.Exception.Message, "APOffice", MessageBoxButtons.OK, MessageBoxIcon.Error);
            };

            AppDomain.CurrentDomain.UnhandledException += (sender, e) =>
            {
                try
                {
                    Exception ex = e.ExceptionObject as Exception;
                    if (ex != null)
                    {
                        new clsFuncs().ErrorLog("AppDomain_UnhandledException", ex.Message + "\r\n" + ex.StackTrace);
                    }
                }
                catch { }
            };

            // Avvio con Splash Screen professionale
            using (frmSplash splash = new frmSplash())
            {
                if (splash.ShowDialog() == DialogResult.OK)
                {
                    Application.Run(new frmAMain());
                }
            }
        }
    }
}
