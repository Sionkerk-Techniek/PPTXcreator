using System.Text.Json.Serialization;

namespace PPTXcreator
{
    public class PastorInfo
    {
        [JsonPropertyName("Naam")]
        public string TitleName { get; set; } = "";

        [JsonIgnore]
        public string Title
        {
            get
            {
                // Split the string at the first space and return the first part
                if (!TitleName.Contains(" ")) return "titel";
                if (TitleName == "Nog niet bekend") return "";
                return TitleName.Split(new char[] { ' ' }, 2)[0];
            }
            set => Title = value;   
        }

        [JsonIgnore]
        public string Name
        {
            get
            {
                // Split the string at the first space and return the second part
                if (!TitleName.Contains(" ")) return "naam";
                if (TitleName == "Nog niet bekend") return TitleName;
                return TitleName.Split(new char[] { ' ' }, 2)[1];
            }
        }

        [JsonPropertyName("Plaats")]
        public string Place { get; set; } = "plaats";
    }
}
