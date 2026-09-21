using System;
using System.Drawing;
using System.Windows.Forms;

namespace APOffice
{
    public partial class frmWait : Form
    {
        private static System.Threading.Thread _waitThread;
        private static frmWait _instance;
        private static readonly object _lock = new object();

        public frmWait()
        {
            InitializeComponent();
        }

        public void SetTitle(string title)
        {
            try
            {
                if (this.InvokeRequired)
                {
                    this.BeginInvoke(new Action(() => SetTitle(title)));
                    return;
                }
                if (!string.IsNullOrEmpty(title))
                    lblTitle.Text = title.ToUpper();
                else
                    lblTitle.Text = "ATTENDERE...";
            }
            catch { }
        }

        public static void UpdateText(string title)
        {
            lock (_lock)
            {
                var inst = _instance;
                if (inst != null && !inst.IsDisposed)
                {
                    inst.SetTitle(title);
                }
            }
        }

        public static void ShowWait(string title = "")
        {
            lock (_lock)
            {
                if (_waitThread != null && _waitThread.IsAlive)
                {
                    var inst = _instance;
                    if (inst != null && !inst.IsDisposed)
                    {
                        inst.SetTitle(title);
                    }
                    return;
                }

                _waitThread = new System.Threading.Thread(() =>
                {
                    try
                    {
                        Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
                        Application.ThreadException += (s, e) => { };
                        _instance = new frmWait();
                        _instance.SetTitle(title);
                        Application.Run(_instance);
                    }
                    catch { }
                    finally
                    {
                        _instance = null;
                    }
                });
                _waitThread.IsBackground = true;
                _waitThread.SetApartmentState(System.Threading.ApartmentState.STA);
                _waitThread.Start();
            }
        }

        public static void CloseWait()
        {
            lock (_lock)
            {
                var inst = _instance;
                _instance = null;

                if (inst != null)
                {
                    try
                    {
                        if (inst.IsHandleCreated && !inst.IsDisposed)
                        {
                            inst.BeginInvoke(new Action(() =>
                            {
                                try
                                {
                                    inst.Close();
                                    inst.Dispose();
                                }
                                catch { }
                            }));
                        }
                    }
                    catch { }
                }

                if (_waitThread != null && _waitThread.IsAlive)
                {
                    try
                    {
                        _waitThread.Join(200);
                    }
                    catch { }
                }
                _waitThread = null;
            }
        }
    }
}
