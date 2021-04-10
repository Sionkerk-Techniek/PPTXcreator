using System.Windows.Forms;

namespace PPTXcreator
{
    public enum ElementType
    {
        Reading,
        PsalmOB,
        PsalmWK,
        SongWK,
        SongOther
    }

    /// <summary>
    /// Basic object which contains information about a
    /// part of the service (a song or reading)
    /// </summary>
    public class ServiceElement
    {
        public ElementType Type { get; }
        /// <summary>
        /// Appears in the presentation before the service and on the
        /// first line in the presentation during the service
        /// </summary>
        public string Title { get; }
        /// <summary>
        /// Appears in the second line in the presentation during the service
        /// </summary>
        public string Subtitle { get; }
        public bool IsSong { get => Type != ElementType.Reading; }

        /// <summary>
        /// Construct a ServiceElement instance from a DataGridViewRow
        /// </summary>
        /// <param name="row">A row from <see cref="Window.dataGridView"/></param>
        public ServiceElement(DataGridViewRow row)
        {
            string type = (string)row.Cells[0].Value;
            string selection = (string)row.Cells[1].Value;
            string songname = (string)row.Cells[2].Value;

            switch (type)
            {
                case "Lezing":
                    Type = ElementType.Reading;
                    Title = FormatTitle(selection);
                    break;
                case "Psalm":
                    Type = ElementType.PsalmOB;
                    Title = "Psalm " + FormatTitle(selection);
                    Subtitle = "Oude berijming";
                    break;
                case "Psalm (WK)":
                    Type = ElementType.PsalmWK;
                    Title = "Psalm " + FormatTitle(selection);
                    Subtitle = "Weerklank";
                    break;
                case "Lied (WK)":
                    Type = ElementType.SongWK;
                    Title = "Lied " + FormatTitle(selection);
                    if (string.IsNullOrWhiteSpace(songname)) Subtitle = "";
                    else Subtitle = songname.Trim();
                    break;
                case "Lied (Overig)":
                    Type = ElementType.SongOther;
                    if (string.IsNullOrWhiteSpace(songname)) Title = "";
                    else Title = songname.Trim();
                    break;
            }
        }

        private string ReplaceLastComma(string input)
        {
            int lastComma = input.LastIndexOf(",");
            if (lastComma == -1) return input;
            return input.Remove(lastComma, 1).Insert(lastComma, " en ");
        }

        private string FormatTitle(string title)
        {
            if (!string.IsNullOrWhiteSpace(title))
            {
                title = ReplaceLastComma(title);  // Replace last comma with 'en'
                title = title.Replace(",", ", ")  // Add space after comma
                             .Replace(":", " : ") // Add space before and after colon
                             .Replace("-", " - ") // Add space before and after dash
                             .Replace("  ", " ")  // Remove double spaces
                             .Trim();             // Remove leading/trailing whitespace
            }

            return title;
        }
    }
}
