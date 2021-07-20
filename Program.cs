using System;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace PPTXcreator
{
    static class Program
    {
        public static Window MainWindow { get; private set; }

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
            Service.UpdateJsonCache();

            if (Settings.Instance.EnableUpdateChecker) Task.Run(() => UpdateChecker.CheckReleases());

            // start UI
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            MainWindow = new Window();
            Application.Run(MainWindow);

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
            filecontents = "";

            try
            {
                filecontents = File.ReadAllText(path);
                return true;
            }
            catch (Exception ex) when (ex is DirectoryNotFoundException || ex is FileNotFoundException)
            {
                Dialogs.GenericWarning($"'{path}' kon niet worden geopend omdat het bestand niet gevonden is.");
                return false;
            }
            catch (Exception ex) when (ex is IOException
                || ex is UnauthorizedAccessException
                || ex is NotSupportedException
                || ex is System.Security.SecurityException)
            {
                Dialogs.GenericWarning($"'{path}' kon niet worden geopend.\n\n" +
                    $"De volgende foutmelding werd gegeven: {ex.Message}");
                return false;
            }
        }

        /// <summary>
        /// Helper function for opening a filestream and handling exceptions
        /// </summary>
        /// <param name="path">Path to the file that has to be read</param>
        /// <param name="stream">The filestream, or null if reading failed</param>
        /// <returns>Whether the function succeeded</returns>
        public static bool TryGetFileStream(string path, out FileStream stream)
        {
            stream = null;

            try
            {
                stream = File.OpenRead(path);
                return true;
            }
            catch (IOException ex) when (ex is DirectoryNotFoundException || ex is FileNotFoundException)
            {
                Dialogs.GenericWarning($"'{path}' kon niet worden geopend omdat het bestand niet gevonden is.");
                return false;
            }
            catch (Exception ex) when (ex is IOException
                || ex is UnauthorizedAccessException
                || ex is NotSupportedException)
            {
                Dialogs.GenericWarning($"'{path}' kon niet worden geopend.\n\n" +
                    $"De volgende foutmelding werd gegeven: {ex.Message}");
                return false;
            }
        }
    }
}
