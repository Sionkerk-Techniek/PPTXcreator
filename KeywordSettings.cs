using System.Collections.Generic;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace PPTXcreator
{
    public class KeywordSettings
    {
        [JsonPropertyName("Begintijd dienst")]
        public string ServiceTime { get; set; } = "[tijd]";

        [JsonPropertyName("Begintijd volgende dienst")]
        public string ServiceNextTime { get; set; } = "[tijd volgende]";

        [JsonPropertyName("Datum volgende dienst")]
        public string ServiceNextDate { get; set; } = "[datum volgende]";

        [JsonPropertyName("Voorganger")]
        public string Pastor { get; set; } = "[voorganger]";

        [JsonPropertyName("Plaats voorganger")]
        public string PastorPlace { get; set; } = "[voorganger plaats]";

        [JsonPropertyName("Volgende voorganger")]
        public string PastorNext { get; set; } = "[volgende voorganger]";

        [JsonPropertyName("Plaats volgende voorganger")]
        public string PastorNextPlace { get; set; } = "[volgende voorganger plaats]";

        public string Organist { get; set; } = "[organist]";

        [JsonPropertyName("Thema")]
        public string Theme { get; set; } = "[thema]";

        [JsonPropertyName("Thema header identifier")]
        public string ThemeHeaderIdentifier { get; set; } = "thema van de preek";

        [JsonPropertyName("Collectedoel 1")]
        public string Collection1 { get; set; } = "[collectedoel 1]";

        [JsonPropertyName("Collectedoel 2")]
        public string Collection2 { get; set; } = "[collectedoel 2]";

        [JsonPropertyName("Collectedoel 3")]
        public string Collection3 { get; set; } = "[collectedoel 3]";

        [JsonPropertyName("Liturgie gezangen")]
        public string Songs { get; set; } = "[zingen]";

        [JsonPropertyName("Liturgie lezingen")]
        public string Readings { get; set; } = "[lezingen]";

        [JsonPropertyName("Tijdens dienst hoofdelement")]
        public string ServiceElementTitle { get; set; } = "[titel]";

        [JsonPropertyName("Tijdens dienst subelement")]
        public string ServiceElementSubtitle { get; set; } = "[subtitel]";
    }
}
