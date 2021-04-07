using System.Text.Json.Serialization;

namespace PPTXcreator
{
    public class CollectionsInfo
    {
        [JsonPropertyName("Collecte 1")]
        public string First { get; set; }

        [JsonPropertyName("Collecte 2")]
        public string Second { get; set; } = "Algemeen kerkenwerk";

        [JsonPropertyName("Collecte 3")]
        public string Third { get; set; }
    }
}
