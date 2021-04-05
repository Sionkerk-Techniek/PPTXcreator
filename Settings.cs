using System;
using System.Collections.Generic;
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

        private string pathOrganistsJson = "organisten.json";
        [JsonPropertyName("Bestandspad organisten json")]
        public string PathOrganistsJson
        {
            get => pathOrganistsJson;
            set => pathOrganistsJson = GetPath(value);
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

        public KeywordSettings Keywords { get; set; }

        /// <summary>
        /// Load settings from the file located at <see cref="SettingsPath"/>
        /// </summary>
        public static void Load()
        {
            // If the settings file cannot be found, notify the user and abort loading
            try
            {
                using (var file = File.OpenText(SettingsPath))
                {
                    JsonSerializerOptions options = new JsonSerializerOptions
                    {
                        AllowTrailingCommas = true,
                        ReadCommentHandling = JsonCommentHandling.Skip
                    };
                    Instance = JsonSerializer.Deserialize<Settings>(file.ReadToEnd(), options);
                }
            }
            catch (FileNotFoundException)
            {
                Dialogs.GenericWarning($"De instellingen konden niet worden geladen " +
                    $"omdat {SettingsPath} niet is gevonden. Standaardwaarden worden gebruikt.");
            }
            catch (Exception ex) when (ex is IOException
                || ex is UnauthorizedAccessException
                || ex is NotSupportedException
            )
            {
                Dialogs.GenericWarning("Instellingen konden niet worden geladen. Standaardwaarden " +
                    "worden gebruikt.\n\n De volgende foutmelding werd gegeven: " + ex.Message);
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
            JsonSerializerOptions options = new JsonSerializerOptions { WriteIndented = true};
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
            // TÖDO: replace with Path.GetRelativePath, which is only available in .NET 5.0+

            Uri uri;
            try
            {
                uri = new Uri(path);
            }
            catch (UriFormatException) // Probably already a relative path
            {
                return path;
            }

            Uri assembly = new Uri(Assembly.GetExecutingAssembly().Location);
            if (assembly.IsBaseOf(uri))
            {
                uri = assembly.MakeRelativeUri(uri);
            }

            string uriString = uri.ToString();
            if (uriString.StartsWith("file:///"))
            {
                uriString = uriString.Substring(8);
            }

            return uriString;
        }
    }
}
