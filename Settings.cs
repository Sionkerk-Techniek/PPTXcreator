using System;
using System.Collections.Generic;
using System.Text;
using System.Linq; // for ToList()
using System.IO;
using System.Windows.Forms;

namespace PPTXcreator
{
    static class Settings
    {
        // Where the settings file is located
        private const string SettingsPath = "./settings.cfg";

        private static Dictionary<string, string> SettingsDictionary = new Dictionary<string, string>
        {
            { "Template voor dienst", "../../../PPTXcreatorfiles/template_voor-dienst-v2.pptx" },
            { "Template tijdens dienst", "../../../PPTXcreatorfiles/template_tijdens-dienst-v2.pptx" },
            { "Template na dienst", "../../../PPTXcreatorfiles/template_na-dienst-v2.pptx" },
            { "QR afbeelding", "../../../PPTXcreatorfiles/QR.png" },
            { "Kerkdiensten Xml", "../../../PPTXcreatorfiles/services.xml" },
            { "Organisten Xml", "../../../PPTXcreatorfiles/organisten.xml" },
            { "Output folder", "./output" },
            { "Volgende dienst", "" },
            { "Check voor updates", "True" },
            { "Automatisch aanvullen", "True" }
        };

        // These properties effectively translate the keys to english
        // and provide easier access to the values
        public static string TemplatePathBefore
        {
            get => SettingsDictionary["Template voor dienst"];
            set => SettingsDictionary["Template voor dienst"] = value;
        }
        public static string TemplatePathDuring
        {
            get => SettingsDictionary["Template tijdens dienst"];
            set => SettingsDictionary["Template tijdens dienst"] = value;
        }
        public static string TemplatePathAfter
        {
            get => SettingsDictionary["Template na dienst"];
            set => SettingsDictionary["Template tijdens dienst"] = value;
        }
        public static string ImagePath
        {
            get => SettingsDictionary["QR afbeelding"];
            set => SettingsDictionary["QR afbeelding"] = value;
        }
        public static string ServicesXml
        {
            get => SettingsDictionary["Kerkdiensten Xml"];
            set => SettingsDictionary["Template tijdens dienst"] = value;
        }
        public static string OrganistXml
        {
            get => SettingsDictionary["Organisten Xml"];
            set => SettingsDictionary["Organisten Xml"] = value;
        }
        public static string OutputFolderPath
        {
            get => SettingsDictionary["Output folder"];
            set => SettingsDictionary["Output folder"] = value;
        }
        public static string LastFutureService
        {
            get => SettingsDictionary["Volgende dienst"];
            set => SettingsDictionary["Volgende dienst"] = value;
        }
        public static bool CheckForUpdates
        {
            get => bool.Parse(SettingsDictionary["Check voor updates"]);
            set => SettingsDictionary["Check voor updates"] = value.ToString();
        }
        public static bool AutoPopulate
        {
            get => bool.Parse(SettingsDictionary["Automatisch aanvullen"]);
            set => SettingsDictionary["Automatisch aanvullen"] = value.ToString();
        }

        /// <summary>
        /// Load settings from the file located at <see cref="SettingsPath"/>
        /// </summary>
        public static void Load()
        {
            // If the settings file cannot be found, notify the user and abort loading
            if (!FileAvailable()) return;

            // Loop over every line in the settings file
            foreach (string line in File.ReadAllLines(SettingsPath))
            {
                // Skip lines without key-value pair
                if (!line.Contains("=")) continue;
                
                // Split every line in a settingID and a value (separated by '=')
                (string settingID, string value) = ParseSetting(line);

                // Change the respective settingID to the new value
                if (SettingsDictionary.ContainsKey(settingID) && !string.IsNullOrWhiteSpace(value))
                {
                    SettingsDictionary[settingID] = value;
                }
            }
        }

        /// <summary>
        /// Save all settings to the file located at <see cref="SettingsPath"/>
        /// </summary>
        public static void Save()
        {
            // Check if the file exists before writing to it
            if (!File.Exists(SettingsPath)) return;

            // Copy the dictionary and read the settings file
            Dictionary<string, string> settingsDictionary = SettingsDictionary;
            List<string> settingsLines = File.ReadAllLines(SettingsPath).ToList();

            // Overwrite every setting in settingsLines with the value in the dictionary
            for (int i = 0; i < settingsLines.Count; i++)
            {
                string line = settingsLines[i];
                if (line.StartsWith("#") || !line.Contains("=")) continue;
                
                (string settingID, _) = ParseSetting(line);
                if (settingsDictionary.ContainsKey(settingID))
                {
                    settingsLines[i] = $"{settingID} = {settingsDictionary[settingID]}";
                    settingsDictionary.Remove(settingID);
                }
            }

            // Append all other settings which weren't in the file before
            foreach (KeyValuePair<string, string> keyValuePair in settingsDictionary)
            {
                settingsLines.Add($"{keyValuePair.Key} = {keyValuePair.Value}");
            }
            
            // Write the settings to file
            File.WriteAllLines(SettingsPath, settingsLines);
        }

        /// <summary>
        /// Parses a string by splitting at the first equals sign and stripping leading/trailing whitespace
        /// </summary>
        /// <param name="settingLine">The inputline to be parsed</param>
        /// <returns>The settingID and the value</returns>
        private static (string, string) ParseSetting(string settingLine)
        {
            string[] lineElements = settingLine.Split("=".ToCharArray(), 2, StringSplitOptions.None);
            string settingID = lineElements[0].Trim();
            string value = lineElements[1].Trim();
            return (settingID, value);
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
