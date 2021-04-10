using System;
using System.IO;
using System.Reflection;
using System.Text.Json;
using System.Text.Json.Serialization;
using System.Security;

namespace PPTXcreator
{
    class Settings
    {
        // Where the settings file is located
        private const string SettingsPath = @"settings.json";

        private static Settings instance;
        public static Settings Instance
        {
            get
            {
                if (instance == null) instance = new Settings();
                return instance;
            }
            set => instance = value;
        }

        private string pathTemplateBefore = "template_voor_dienst.pptx";
        [JsonPropertyName("Bestandspad template voor dienst")]
        public string PathTemplateBefore
        {
            get => pathTemplateBefore;
            set => pathTemplateBefore = GetPath(value);
        }

        private string pathTemplateDuring = "template_tijdens_dienst.pptx";
        [JsonPropertyName("Bestandspad template tijdens dienst")]
        public string PathTemplateDuring
        {
            get => pathTemplateDuring;
            set => pathTemplateDuring = GetPath(value);
        }

        private string pathTemplateAfter = "template_na_dienst.pptx";
        [JsonPropertyName("Bestandspad template na dienst")]
        public string PathTemplateAfter
        {
            get => pathTemplateAfter;
            set => pathTemplateAfter = GetPath(value);
        }

        private string pathServicesJson = "diensten.json";
        [JsonPropertyName("Bestandspad kerkdiensten json")]
        public string PathServicesJson
        {
            get => pathServicesJson;
            set => pathServicesJson = GetPath(value);
        }

        private string pathOutputFolder = "./output";
        [JsonPropertyName("Outputfolder")]
        public string PathOutputFolder
        {
            get => pathOutputFolder;
            set => pathOutputFolder = GetPath(value);
        }

        private string pathQRImage;
        [JsonIgnore]
        public string PathQRImage
        {
            get => pathQRImage;
            set => pathQRImage = GetPath(value);
        }

        [JsonPropertyName("QR placeholder beschrijving")]
        public string ImageDescription { get; set; } = "QR-code";

        [JsonPropertyName("Volgende dienst")]
        public DateTime NextService { get; set; }

        [JsonPropertyName("Check voor updates")]
        public bool EnableUpdateChecker { get; set; } = true;

        [JsonPropertyName("Automatisch invullen van velden")]
        public bool EnableAutoPopulate { get; set; } = true;

        [JsonPropertyName("QR-afbeeldingen bewerken")]
        public bool EnableEditQR { get; set; } = true;

        [JsonPropertyName("QR opslaan in de outputfolder")]
        public bool EnableExportQR { get; set; } = true;

        public KeywordSettings Keywords { get; set; } = new KeywordSettings();

        /// <summary>
        /// Load settings from the file located at <see cref="SettingsPath"/>
        /// </summary>
        public static void Load()
        {
            if (!Program.TryGetFileContents(SettingsPath, out string filecontents)) return;

            try
            {
                JsonSerializerOptions options = new JsonSerializerOptions()
                {
                    AllowTrailingCommas = true,
                    ReadCommentHandling = JsonCommentHandling.Skip
                };
                Instance = JsonSerializer.Deserialize<Settings>(filecontents, options);
            }
            catch (JsonException ex)
            {
                Dialogs.GenericWarning("Instellingen konden niet worden geladen vanwege ongeldige waarden in " +
                    $"{SettingsPath}. Standaardwaarden worden gebruikt.\n\nVolledige error: {ex.Message}");
            }
        }

        /// <summary>
        /// Save all settings to the file located at <see cref="SettingsPath"/>
        /// </summary>
        public static void Save()
        {
            JsonSerializerOptions options = new JsonSerializerOptions() { WriteIndented = true };
            string json = JsonSerializer.Serialize(Instance, options);
            try
            {
                File.WriteAllText(SettingsPath, json);
            }
            catch (Exception ex) when (
                ex is IOException
                || ex is UnauthorizedAccessException
                || ex is SecurityException
            )
            {
                Dialogs.GenericWarning("Instellingen konden niet worden opgeslagen. " +
                    $"De volgende foutmelding werd gegeven: {ex.Message}");
            }
        }

        /// <summary>
        /// Creates a relative path from the executing assembly to the file at
        /// <paramref name="path"/> if the file is in the same directory as or
        /// in a subdirectory of the assembly, or an absolute path otherwise
        /// </summary>
        public static string GetPath(string path)
        {
            Uri assemblyPath = new Uri(AppContext.BaseDirectory);
            Uri targetPath = new Uri(Path.GetFullPath(path));
            string relativePath = assemblyPath.MakeRelativeUri(targetPath).ToString();
            relativePath = relativePath.Replace('/', '\\'); // URIs use forward slashes
            
            if (relativePath.StartsWith("..")) return path;
            else return relativePath;
        }
    }
}
