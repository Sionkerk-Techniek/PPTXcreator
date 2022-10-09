using System.Text.Json.Serialization;

namespace PPTXcreator
{
    public class ImageSettings
    {
        [JsonPropertyName("X-offset")]
        public int OffsetX { get; set; } = 20;

        [JsonPropertyName("Y-offset")]
        public int OffsetY { get; set; } = 20;

        [JsonPropertyName("Breedte")]
        public int Width { get; set; } = 360;

        [JsonPropertyName("Hoogte")]
        public int Height { get; set; } = 360;

        public float Threshold { get; set; } = 0.6f;

        public bool IsValid()
        {
            return OffsetX > 0 && OffsetY > 0 && Width > 0 && Height > 0
                && 0 <= Threshold && Threshold <= 1;
        }

        public override string ToString()
        {
            return $"{Width}x{Height}+{OffsetX}+{OffsetY}";
        }
    }
}
