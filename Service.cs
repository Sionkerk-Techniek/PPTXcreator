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

        [JsonIgnore] // static objects are already ignored but this makes it more clear
        public static JsonElement[] JsonCache { get; private set; }

        /// <summary>
        /// Get an array of services from the JSON file at <see cref="Settings.PathServicesJson"/>
        /// and save it to <see cref="JsonCache"/>
        /// </summary>
        public static void UpdateJsonCache()
        {
            // Load the file contents
            if (!Program.TryGetFileContents(Settings.Instance.PathServicesJson, out string fileContent)) return;

            // Parse the file contents
            try
            {
                JsonDocumentOptions options = new JsonDocumentOptions()
                {
                    AllowTrailingCommas = true,
                    CommentHandling = JsonCommentHandling.Skip
                };

                // Parse the json string and order the JsonElements by the Datetime property
                JsonCache = JsonDocument.Parse(fileContent, options).RootElement.EnumerateArray()
                    .OrderBy(element => element.GetProperty("Datetime").GetDateTime()).ToArray();
            }
            catch (Exception ex) when (ex is JsonException || ex is InvalidOperationException)
            {
                Dialogs.GenericWarning($"'{Settings.Instance.PathServicesJson}'" +
                    $" heeft niet de juiste structuur.\n\nDe volgende foutmelding werd gegeven: {ex.Message}");
            }
        }

        /// <summary>
        /// Get the service object for the service at <paramref name="dateTime"/>
        /// </summary>
        public static Service GetCurrent(DateTime dateTime)
        {
            Service current = new Service();
            if (JsonCache == null) return current;

            // Get the JsonElement which matches dateTime from JsonCache
            JsonElement currentElement = (from JsonElement service in JsonCache
                                          where service.GetProperty("Datetime").GetDateTime() == dateTime
                                          select service).FirstOrDefault();

            if (currentElement.ValueKind != JsonValueKind.Undefined)
                current = JsonSerializer.Deserialize<Service>(currentElement.GetRawText());

            return current;
        }

        /// <summary>
        /// Get the service object for the first service after <paramref name="dateTime"/>
        /// </summary>
        public static Service GetNext(DateTime dateTime)
        {
            Service next = new Service();
            if (JsonCache == null) return next;

            // Get the first JsonElement later than dateTime from JsonCache
            JsonElement nextElement = (from JsonElement service in JsonCache
                                       where service.GetProperty("Datetime").GetDateTime() > dateTime
                                       select service).FirstOrDefault(); // JsonCache is ordered

            if (nextElement.ValueKind != JsonValueKind.Undefined)
                next = JsonSerializer.Deserialize<Service>(nextElement.GetRawText());

            return next;
        }

        public static DateTime GetPreviousDatetime(DateTime current)
        {
            if (JsonCache == null) return DateTime.MinValue;

            List<DateTime> dateTimes = (from JsonElement service in JsonCache
                                        let dateTime = service.GetProperty("Datetime").GetDateTime()
                                        where dateTime < current
                                        select dateTime).ToList();

            if (dateTimes.Count == 0) return DateTime.MinValue;
            else return dateTimes.Last();
        }

        public static DateTime GetNextDatetime(DateTime current)
        {
            if (JsonCache == null) return DateTime.MinValue;

            List<DateTime> dateTimes = (from JsonElement service in JsonCache
                                        let dateTime = service.GetProperty("Datetime").GetDateTime()
                                        where dateTime > current
                                        select dateTime).ToList();

            if (dateTimes.Count == 0) return DateTime.MinValue;
            else return dateTimes.First();
        }
    }
}
