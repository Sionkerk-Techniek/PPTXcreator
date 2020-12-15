using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;


namespace PPTXcreator
{
    static class Settings
    {
        // Where the settings file is located
        private const string SettingsPath = "./settings.cfg";

        // Settings that can be defined in settings.cfg
        public static string OutputFolder = "./presentaties";                       // Folder the output will be saved to
        public static string InfoJSON;                                              // File which holds all known information for future services
        public static string LastFutureService;                                     // Holds the last used datetime for 'the next service'
        public static string PPTXTemplatePre = "./template_voordienst.pptx";        // Powerpoint templates
        public static string PPTXTemplateDuring = "./template_tijdensdienst.pptx";  // TODO: use only one collection of all unique slides to build pptxs from
        public static string PPTXTemplateAfter = "./template_nadienst.pptx";
        public static string Collecte2 = "Algemeen kerkenwerk";                     // Text used as 2e collecte
        public static bool CheckForUpdates = true;                                  // If true, the program will check for new releases on GitHub


        /// <summary>
        /// Load settings from the file located at SettingsPath
        /// </summary>
        public static void Load()
        {
            // If the settings file cannot be found, notify the user and abort loading
            if (!FileAvailable()) return;

            // Loop over every line in the settings file
            foreach (string line in File.ReadAllLines(SettingsPath))
            {
                // Skip empty lines
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Split every line in a settingID and a value (separated by '=')
                ParseSetting(line, out string settingID, out string value);

                // Change the respective settingID to the new value
                switch (settingID)
                {
                    case "OUTPUT_FOLDER":
                        OutputFolder = value;
                        break;
                    case "KERKBODE_JSON":
                        InfoJSON = value;
                        break;
                    case "PPTX_TEMPLATE_VOOR":
                        PPTXTemplatePre = value;
                        break;
                    case "PPTX_TEMPLATE_TIJDENS":
                        PPTXTemplateDuring = value;
                        break;
                    case "PPTX_TEMPLATE_NA":
                        PPTXTemplateAfter = value;
                        break;
                    case "VOLGENDE_DIENST":
                        LastFutureService = value;
                        break;
                    case "CHECK_UPDATES_ON_STARTUP":
                        // Try to parse the value as a boolean, and set to the default 'true' if conversion failed
                        bool parseSuccess = bool.TryParse(value, out CheckForUpdates);
                        if (!parseSuccess) CheckForUpdates = true;
                        break;
                    default:
                        Console.WriteLine($"Unknown settingID: {settingID}");
                        break;
                }
            }
        }


        /// <summary>
        /// Change the settingID 'settingID', and write to the file with <see cref="WriteSetting"/>
        /// </summary>
        /// <param name="settingID">Any setting identifier. Valid IDs are: OUTPUT_FOLDER, KERKBODE_JSON,
        /// PPTX_TEMPLATE_VOOR, PPTX_TEMPLATE_TIJDENS, PPTX_TEMPLATE_NA, VOLGENDE_DIENST, CHECK_UPDATES_ON_STARTUP</param>
        /// <param name="value">The new value of the setting to be saved</param>
        public static void ChangeSetting(string settingID, string value)
        {
            if (!FileAvailable()) return;

            // Loop over the lines in the settings file and keep track of the line number
            string[] settingLines = File.ReadAllLines(SettingsPath);
            for (int lineNumber = 0; lineNumber < settingLines.Length; lineNumber++)
            {
                string line = settingLines[lineNumber];
                if (string.IsNullOrWhiteSpace(line)) continue;

                // Split every line in a settingID and a value (separated by '='). Discard the value.
                ParseSetting(line, out string fileSettingID, out _);

                // If the settingID is found, replace the line with a new value and end this function
                if (fileSettingID == settingID)
                {
                    WriteSetting(settingLines, lineNumber, $"{settingID} = {value}");
                    return;
                }
            }

            // If the setting hasn't been found, append to the file by resizing settingLines
            Array.Resize(ref settingLines, settingLines.Length + 1);
            WriteSetting(settingLines, settingLines.Length - 1, $"{settingID} = {value}");
        }


        /// <summary>
        /// Parses a string by splitting at the first equals sign and stripping leading/trailing whitespace
        /// </summary>
        /// <param name="settingLine">The inputline to be parsed</param>
        /// <param name="settingID">The output settingID</param>
        /// <param name="value">The output value of the settingID</param>
        private static void ParseSetting(string settingLine, out string settingID, out string value)
        {
            string[] lineElements = settingLine.Split("=".ToCharArray(), 2, StringSplitOptions.None);
            settingID = lineElements[0].Trim().ToUpper();
            value = lineElements[1].Trim();
        }


        /// <summary>
        /// Changes a single element in an array of settings and writes the array to the file at SettingsPath
        /// </summary>
        /// <param name="settingsArray">The settings to be written</param>
        /// <param name="lineNumber">The index of the element that should be changed</param>
        /// <param name="setting">The setting that should go in settingsArray[lineNumber]</param>
        private static void WriteSetting(string[] settingsArray, int lineNumber, string setting)
        {
            settingsArray[lineNumber] = setting;
            File.WriteAllLines(SettingsPath, settingsArray); // TODO: check for write permission maybe
        }


        /// <summary>
        /// Check whether the settings file exists at SettingsPath.
        /// If false, call <see cref="FileCreate"/>.
        /// </summary>
        /// <returns>Whether or not the settings file is available</returns>
        private static bool FileAvailable()
        {
            if (!File.Exists(SettingsPath))
            {
                // If a settings file is created in FileCreate(), return true. If not, return false
                if (FileCreate()) return true;
                else return false;
            }
            else
            {
                return true;
            }
        }


        /// <summary>
        /// Ask the user if a new settings file at should be created at SettingsPath.
        /// </summary>
        /// <returns>True if the file was created, false otherwise</returns>
        private static bool FileCreate()
        {
            // Ask the user if a new settings file should be created
            DialogResult dialogResult = MessageBox.Show(
                $"Het configuratiebestand is niet gevonden op locatie '{SettingsPath}'. Wil je een nieuw bestand maken?",
                "Er is een fout opgetreden",
                MessageBoxButtons.YesNoCancel,
                MessageBoxIcon.Warning
            );

            // Create a file if the user answered 'yes'
            if (dialogResult == DialogResult.Yes)
            {
                File.Create(SettingsPath).Close();
                return true;
            }
            else
            {
                // If the user doesn't want to create a new settings file, warn them the settings won't be saved and stop
                MessageBox.Show(
                    "Instellingen kunnen niet worden opgeslagen en gaan verloren als dit programma wordt afgesloten.",
                    "",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Information
                );
                return false;
            }
        }
    }
}
