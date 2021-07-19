using System;
using System.Collections.Generic;
using System.Linq;
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
        /// Get an array of service elements from the JSON file at <see cref="Settings.PathServicesJson"/>
        /// </summary>
        /// <returns>An array of service elements, or null if the file could not be read</returns>
        public static JsonElement[] GetJsonElements()
        {
            JsonElement[] services = null;

            // Load the file contents
            // TODO: caching - current method will give performance issues if the files become large or are loaded very often
            if (!Program.TryGetFileContents(Settings.Instance.PathServicesJson, out string servicesFile))
            {
                return services;
            }

            // Parse the file contents
            try
            {
                JsonDocumentOptions options = new JsonDocumentOptions()
                {
                    AllowTrailingCommas = true,
                    CommentHandling = JsonCommentHandling.Skip
                };
                services = JsonDocument.Parse(servicesFile, options).RootElement.EnumerateArray().ToArray();
            }
            catch (Exception ex) when (ex is JsonException || ex is InvalidOperationException)
            {
                Dialogs.GenericWarning($"'{Settings.Instance.PathServicesJson}'" +
                    $" heeft niet de juiste structuur.\n\nDe volgende foutmelding werd gegeven: {ex.Message}");
            }

            return services;
        }

        /// <summary>
        /// Build service objects for the current and next service from JSON data
        /// </summary>
        public static (Service current, Service next) GetCurrentAndNext(DateTime datetime)
        {
            // Initialize some values
            Service current = new Service();
            Service next = new Service();
            JsonElement[] services = GetJsonElements();

            if (services == null) return (current, next);

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

            return (current, next);
        }

        public static DateTime GetPrevious(DateTime current)
        {
            JsonElement[] services = GetJsonElements();
            List<DateTime> dateTimes = (from JsonElement service in services
                                        let dateTime = service.GetProperty("Datetime").GetDateTime()
                                        orderby dateTime
                                        where dateTime < current
                                        select dateTime).ToList();

            if (dateTimes.Count == 0) return DateTime.MinValue;
            else return dateTimes.Last();
        }

        public static DateTime GetNext(DateTime current)
        {
            JsonElement[] services = GetJsonElements();
            List<DateTime> dateTimes = (from JsonElement service in services
                                        let dateTime = service.GetProperty("Datetime").GetDateTime()
                                        orderby dateTime
                                        where dateTime > current
                                        select dateTime).ToList();

            if (dateTimes.Count == 0) return DateTime.MinValue;
            else return dateTimes.First();
        }
    }
}
