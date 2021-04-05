using System;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PPTXcreator
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            // Configure error handling
            Application.SetUnhandledExceptionMode(UnhandledExceptionMode.CatchException);
            Application.ThreadException += new System.Threading.ThreadExceptionEventHandler(CrashHandlerUI);
            AppDomain.CurrentDomain.UnhandledException += new UnhandledExceptionEventHandler(CrashHandlerDomain);

            Settings.Load();
            Settings.Save();

            if (Settings.Instance.EnableUpdateChecker) Task.Run(() => UpdateChecker.CheckReleases());

            // start UI
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Window());

            Settings.Save();
        }

        /// <summary>
        /// Handle unhandled exceptions in the UI thread
        /// </summary>
        private static void CrashHandlerUI(object sender, System.Threading.ThreadExceptionEventArgs args)
        {
            Dialogs.CrashNotification(args.Exception.Message + "\n\nStack Trace: " + args.Exception.StackTrace);
            Application.Exit();
        }

        /// <summary>
        /// Handle unhandled exceptions which are not in the UI thread
        /// </summary>
        private static void CrashHandlerDomain(object sender, UnhandledExceptionEventArgs args)
        {
            Exception exception = (Exception)args.ExceptionObject;
            Dialogs.CrashNotification(exception.Message + "\n\nStack Trace: " + exception.StackTrace);
            Application.Exit();
        }
    }
}
