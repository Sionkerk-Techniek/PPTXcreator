using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Globalization;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PPTXcreator
{
    public class Service
    {
        public DateTime Datetime { get; set; }

        [JsonPropertyName("Voorganger")]
        public PastorInfo Pastor { get; set; } = new PastorInfo();

        [JsonPropertyName("Collecten")]
        public CollectionsInfo Collections { get; set; } = new CollectionsInfo();

        public string Organist { get; set; } = "organist";

        /// <summary>
        /// Data class containing information about a service
        /// </summary>
        public static (Service current, Service next) GetCurrentAndNext(DateTime datetime)
        {
            // Initialize some values
            JsonElement[] services;
            IEnumerable<JsonElement> organists;
            Service current = new Service();
            Service next = new Service();

            // Load the file contents
            if (!Program.TryGetFileContents(Settings.Instance.PathServicesJson, out string servicesFile)
                || !Program.TryGetFileContents(Settings.Instance.PathOrganistsJson, out string organistsFile))
            {
                return (current, next);
            }

            // Parse the file contents
            try
            {
                JsonDocumentOptions options = new JsonDocumentOptions {
                    AllowTrailingCommas = true, CommentHandling = JsonCommentHandling.Skip };
                services = JsonDocument.Parse(servicesFile, options).RootElement.EnumerateArray().ToArray();
                organists = JsonDocument.Parse(organistsFile, options).RootElement.EnumerateArray();
            }
            catch (Exception ex) when (ex is JsonException || ex is InvalidOperationException)
            {
                Dialogs.GenericWarning($"{Settings.Instance.PathServicesJson} of {Settings.Instance.PathOrganistsJson}" +
                    $" hebben niet de juiste structuur.\n\nDe volgende foutmelding werd gegeven: {ex.Message}");
                return (current, next);
            }

            // Iterate over all services in servicesjson and find the one which matches datetime
            for (int i = 0; i < services.Length; i++)
            {
                if (services[i].GetProperty("Datetime").GetDateTime() == datetime)
                {
                    current = JsonSerializer.Deserialize<Service>(services[i].GetRawText());

                    // Try to get the next service if possible
                    try
                    {
                        next = JsonSerializer.Deserialize<Service>(services[i + 1].GetRawText());
                    }
                    catch (IndexOutOfRangeException)
                    {
                        next = new Service();
                    }

                    break;
                }
            }

            // Get the organist
            current.Organist = (from organist in organists
                                where organist.GetProperty("Datetime").GetDateTime() == datetime
                                select organist.GetProperty("Organist").GetString()).FirstOrDefault();
            if (current.Organist == null) current.Organist = "organist";

            return (current, next);
        }
    }
}
