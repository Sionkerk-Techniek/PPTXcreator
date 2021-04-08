using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

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

        /// <summary>
        /// Helper function for loading files and handling exceptions
        /// </summary>
        /// <param name="path">Path to the file that has to be read</param>
        /// <param name="filecontents">The contents of the file, or an empty string if reading failed</param>
        /// <returns>Whether the function succeeded</returns>
        public static bool TryGetFileContents(string path, out string filecontents)
        {
            try
            {
                filecontents = File.ReadAllText(path);
                return true;
            }
            catch (Exception ex) when (ex is DirectoryNotFoundException or FileNotFoundException)
            {
                Dialogs.GenericWarning($"{path} kon niet worden geopgend omdat het bestand niet gevonden is.");
                filecontents = "";
                return false;
            }
            catch (Exception ex) when (ex is IOException
                or UnauthorizedAccessException
                or NotSupportedException
                or System.Security.SecurityException)
            {
                Dialogs.GenericWarning($"{path} kon niet worden geopend.\n\n" +
                    $"De volgende foutmelding werd gegeven: {ex.Message}");
                filecontents = "";
                return false;
            }
        }
    }
}
