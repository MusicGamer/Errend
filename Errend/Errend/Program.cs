using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Threading;
using System.Runtime.InteropServices;
using System.Globalization;
using System.Diagnostics;


namespace Errend
{
    static class Program
    {
        [DllImport("user32.dll")]
        static extern bool SetForegroundWindow(IntPtr hWnd);

        [DllImport("user32.dll")]
        static extern bool ShowWindow(IntPtr hWnd, int nCmdShow);

        [DllImport("user32.dll")]
        public static extern bool IsIconic(IntPtr handle);
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        private static Mutex mutex = null;
        private static IntPtr handle;
        [STAThread]
        static void Main()
        {           
            const string appName = "Errend";
            bool createdNew;
            mutex = new Mutex(true, appName, out createdNew);
            if (!createdNew)
            {
                Process[] p = Process.GetProcessesByName("Errend");
                handle = (p[0].Id != Process.GetCurrentProcess().Id ? p[0] : p[1]).MainWindowHandle;
                SetForegroundWindow(handle);
                if (IsIconic(handle))
                {
                    ShowWindow((p[0].Id != Process.GetCurrentProcess().Id ? p[0] : p[1]).MainWindowHandle, 9);
                }
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());
        }
    }
}
