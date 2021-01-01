using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using System.Globalization;

namespace PPTXcreator
{
    public partial class Window : Form
    {
        public Window()
        {
            InitializeComponent();
        }

        private void ShowWarning(string message)
        {
            MessageBox.Show(
                message,
                "Er is een fout opgetreden",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }

        /// <summary>
        /// Selects a file and sets this.textBoxQRPath.Text and
        /// QRFile to the path
        /// </summary>
        private void ButtonSelectQR(object sender, EventArgs e)
        {
            // Create an OpenFileDialog object which can open .jpeg and .png files
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "JPEG (*.jpeg)|*.jpeg|PNG (*.png)|*.png|Alle bestanden (*.*)|*.*",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer),
                Title = "Selecteer de QR-code",
                RestoreDirectory = true
            };

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(fileDialog.FileName))
                {
                    this.textBoxQRPath.Text = Settings.GetPath(fileDialog.FileName);
                    Settings.ImagePath = fileDialog.FileName;
                }
                else
                {
                    ShowWarning("Het geselecteerde bestand kan niet worden gevonden");
                }
            }
        }

        /// <summary>
        /// Selects a directory and sets <see cref="textBoxOutputFolder"/>.Text
        /// and <see cref="Settings.OutputFolder"/> to the path
        /// </summary>
        private void ButtonSelectOutputFolder(object sender, EventArgs e)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog
            {
                Description = "Selecteer de output folder",
            };

            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                if (Directory.Exists(folderDialog.SelectedPath))
                {
                    textBoxOutputFolder.Text = Settings.GetPath(folderDialog.SelectedPath);
                    Settings.OutputFolderPath = folderDialog.SelectedPath;
                }
                else
                {
                    ShowWarning("De geselecteerde folder kan niet worden gevonden");
                }
            }
        }

        private void ButtonSelectTemplateBefore(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "PowerPoint (*.pptx)|*.pptx|Alle bestanden (*.*)|*.*",
                InitialDirectory = Directory.GetCurrentDirectory(),
                Title = "Selecteer de 'voor de dienst' template",
                RestoreDirectory = true,
            };

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxTemplatePre.Text = Settings.GetPath(fileDialog.FileName);
                Settings.TemplatePathBefore = fileDialog.FileName;
            }
        }

        private void ButtonSelectTemplateDuring(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "PowerPoint (*.pptx)|*.pptx|Alle bestanden (*.*)|*.*",
                InitialDirectory = Directory.GetCurrentDirectory(),
                Title = "Selecteer de 'tijdens de dienst' template",
                RestoreDirectory = true,
            };

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxTemplateDuring.Text = Settings.GetPath(fileDialog.FileName);
                Settings.TemplatePathDuring = fileDialog.FileName;
            }
        }

        private void ButtonSelectTemplateAfter(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "PowerPoint (*.pptx)|*.pptx|Alle bestanden (*.*)|*.*",
                InitialDirectory = Directory.GetCurrentDirectory(),
                Title = "Selecteer de 'na de dienst' template",
                RestoreDirectory = true,
            };

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxTemplateAfter.Text = Settings.GetPath(fileDialog.FileName);
                Settings.TemplatePathAfter = fileDialog.FileName;
            }
        }

        private void ButtonSelectXmlServices(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "XML (*.xml)|*.xml|Alle bestanden (*.*)|*.*",
                InitialDirectory = Directory.GetCurrentDirectory(),
                Title = "Selecteer het XML-bestand met diensten",
                RestoreDirectory = true,
            };

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxXmlServices.Text = Settings.GetPath(fileDialog.FileName);
                Settings.ServicesXml = fileDialog.FileName;
            }
        }

        private void ButtonSelectXmlOrganists(object sender, EventArgs e)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = "XML (*.xml)|*.xml|Alle bestanden (*.*)|*.*",
                InitialDirectory = Directory.GetCurrentDirectory(),
                Title = "Selecteer het XML-bestand met organisten",
                RestoreDirectory = true,
            };

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                textBoxXmlOrganist.Text = Settings.GetPath(fileDialog.FileName);
                Settings.OrganistXml = fileDialog.FileName;
            }
        }

        private void DateTimePickerNu_Leave(object sender, EventArgs e)
        {
            string dateTimeString = dateTimePickerNu.Value.ToString("yyyy-MM-dd H:mm",
                CultureInfo.InvariantCulture);

            if (Settings.AutoPopulate)
            {
                Service currentService = Service.GetService(Settings.ServicesXml,
                    Settings.OrganistXml, dateTimeString);

                SetFormData(currentService);
            }
        }

        private void SetFormData(Service service)
        {
            (string dsNowTitle, string dsNowName) = Service.SplitName(service.DsName);
            (string dsNextTitle, string dsNextName) = Service.SplitName(service.NextDsName);

            if (service.Time != DateTime.MinValue)
            {
                textBoxVoorgangerNuTitel.Text = dsNowTitle;
                textBoxVoorgangerNuNaam.Text = dsNowName;
                textBoxVoorgangerNuPlaats.Text = service.DsPlace;

                textBoxCollecte1.Text = service.Collection_1;
                textBoxCollecte3.Text = service.Collection_3;
            }

            if (service.NextTime != DateTime.MinValue)
            {
                dateTimePickerNext.Value = service.NextTime;
                textBoxVoorgangerNextTitel.Text = dsNextTitle;
                textBoxVoorgangerNextNaam.Text = dsNextName;
                textBoxVoorgangerNextPlaats.Text = service.NextDsPlace;
            }

            if (!string.IsNullOrWhiteSpace(service.Organist))
            {
                textBoxOrganist.Text = service.Organist;
            }
        }

        public Dictionary<string, string> GetFormKeywords()
        {
            Dictionary<string, string> keywords = new Dictionary<string, string>
            {
                { "[tijd]", GetTime(dateTimePickerNu) },
                { "[tijd next]", GetTime(dateTimePickerNext) },
                { "[datum next]", GetDateLong(dateTimePickerNext) },
                { "[voorganger]", $"{textBoxVoorgangerNuTitel.Text} {textBoxVoorgangerNuNaam.Text}" },
                { "[voorganger plaats]", textBoxVoorgangerNuPlaats.Text },
                { "[voorganger next]", $"{textBoxVoorgangerNextTitel.Text} {textBoxVoorgangerNextNaam.Text}" },
                { "[voorganger next plaats]", textBoxVoorgangerNextPlaats.Text },
                { "[organist]", textBoxOrganist.Text },
                { "[thema]", textBoxThema.Text }
            };

            return keywords;
        }

        public Dictionary<string, string> GetSongKeywords()
        {
            Dictionary<string, string> keywords = new Dictionary<string, string>();

            byte songCounter = 0;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                string type = (string)row.Cells[0].Value;
                string song = (string)row.Cells[1].Value;
                string name = (string)row.Cells[2].Value;

                if (string.IsNullOrWhiteSpace(type)) continue;

                if (type != "Lezing")
                {
                    songCounter++;
                    keywords[$"[lied {songCounter}]"] = GetSongSelection(song, type, name);
                    keywords[$"[liedbundel {songCounter}]"] = GetSongBundle(type, name);
                }
            }

            return keywords;
        }

        public Dictionary<string, string> GetReadingKeywords()
        {
            Dictionary<string, string> keywords = new Dictionary<string, string>();

            byte readingCounter = 0;
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                string type = (string)row.Cells[0].Value;
                string reading = (string)row.Cells[1].Value;

                if (type == "Lezing")
                {
                    readingCounter++;
                    keywords[$"[lezing {readingCounter}]"] = reading.Replace(":", " : ").Replace("-", " - ");
                }
            }

            return keywords;
        }

        private string SongReplaceLastComma(string input)
        {
            int lastComma = input.LastIndexOf(",");
            if (lastComma == -1) return input;
            return input.Remove(lastComma, 1).Insert(lastComma, " en ");
        }

        private string GetSongSelection(string song, string type, string name)
        {
            if (!string.IsNullOrWhiteSpace(song))
            {
                song = SongReplaceLastComma(song);
                song = song.Replace(",", ", ").Replace(":", " : ");
            }

            switch (type)
            {
                case "Psalm":
                    return $"Psalm {song}";
                case "Psalm (WK)":
                    return $"Psalm {song}   WK";
                case "Lied (WK)":
                    return $"Lied {song}";
                default:
                    return name;
            }
        }

        private string GetSongBundle(string type, string name)
        {
            switch(type)
            {
                case "Psalm":
                    return "Oude berijming";
                case "Psalm (WK)":
                    return "Weerklank";
                default:
                    return name;
            }
        }

        private string GetTime(DateTimePicker dateTimePicker)
        {
            return dateTimePicker.Value.ToString("H:mm", CultureInfo.InvariantCulture);
        }

        private string GetDateLong(DateTimePicker dateTimePicker)
        {
            // Dictionary to avoid not having the nl-NL resource when using System.Globalization
            Dictionary<int, string> monthNames = new Dictionary<int, string>()
            {
                { 1, "januari" }, { 2, "februari" }, { 3, "maart" }, { 4, "april" },
                { 5, "mei" }, { 6, "juni" }, { 7, "juli" }, { 8, "augustus"},
                { 9, "september" }, { 10, "oktober" }, { 11, "november" }, { 12, "december" }
            };
            return $"{dateTimePicker.Value.Day} {monthNames[dateTimePicker.Value.Month]}";
        }

        private void ButtonAddDatagridviewRow(object sender, EventArgs e)
        {
            dataGridView.Rows.Add();
        }

        private void ButtonRemoveDatagridviewRow(object sender, EventArgs e)
        {
            if (dataGridView.SelectedRows.Count == 1)
            {
                try
                {
                    dataGridView.Rows.Remove(dataGridView.SelectedRows[0]);
                }
                catch (InvalidOperationException exception)
                {
                    ShowWarning("De rij kan niet worden verwijderd." +
                        $"De volgende foutmelding werd gegeven: {exception.Message}");
                }
            }
            else if (dataGridView.CurrentCell != null)
            {
                try
                {
                    dataGridView.Rows.Remove(dataGridView.CurrentCell.OwningRow);
                }
                catch (InvalidOperationException exception)
                {
                    ShowWarning("De rij kan niet worden verwijderd. " +
                        $"De volgende foutmelding werd gegeven: {exception.Message}");
                }
            }
        }

        private void ButtonMoveUpDatagridviewRow(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow is null) return;

            int index = dataGridView.CurrentRow.Index;
            int rowCount = dataGridView.Rows.Count;
            DataGridViewRow row = dataGridView.CurrentRow;

            if (index == 0 || index == rowCount - 1 || rowCount == 1) return;

            dataGridView.Rows.Remove(row);
            dataGridView.Rows.Insert(index - 1, row);

            UpdateDatagridviewSelection(index - 1);
        }

        private void ButtonMoveDownDatagridviewRow(object sender, EventArgs e)
        {
            if (dataGridView.CurrentRow is null) return;

            int index = dataGridView.CurrentRow.Index;
            int rowCount = dataGridView.Rows.Count;
            DataGridViewRow row = dataGridView.CurrentRow;

            if (index >= rowCount - 2 || rowCount == 1) return;

            dataGridView.Rows.Remove(row);
            dataGridView.Rows.Insert(index + 1, row);

            UpdateDatagridviewSelection(index + 1);
        }

        private void UpdateDatagridviewSelection(int index)
        {
            dataGridView.ClearSelection();
            dataGridView.Rows[index].Selected = true;
            dataGridView.CurrentCell = dataGridView.Rows[index].Cells[0];
        }

        private void ButtonGotoSettingsTab(object sender, EventArgs e)
        {
            tabControl.SelectTab("tabInstellingen");
        }

        private void ButtonNextTab(object sender, EventArgs e)
        {
            int index = tabControl.SelectedIndex;
            tabControl.SelectTab(index + 1);
        }

        private void ButtonPreviousTab(object sender, EventArgs e)
        {
            int index = tabControl.SelectedIndex;
            tabControl.SelectTab(index - 1);
        }

        private bool CheckValidInputs()
        {
            StringBuilder warning = new StringBuilder();
            List<string> invalidInputs = new List<string>();

            string[] fieldNames = new string[]
            {
                "titel van de voorganger van deze dienst", "titel van de voorganger van de volgende dienst",
                "naam van de voorganger van deze dienst", "naam van de voorganger van de volgende dienst",
                "plaats van de voorganger van deze dienst", "plaats van de voorganger van de volgende dienst",
                "collectedoel 1", "collectedoel 3", "naam van de organist", "bestandspad QR-code", "liturgie"
            };
            string[] inputValues = new string[]
            {
                textBoxVoorgangerNuTitel.Text, textBoxVoorgangerNextTitel.Text,
                textBoxVoorgangerNuNaam.Text, textBoxVoorgangerNextNaam.Text,
                textBoxVoorgangerNuPlaats.Text, textBoxVoorgangerNextPlaats.Text,
                textBoxCollecte1.Text, textBoxCollecte3.Text, textBoxOrganist.Text,
                textBoxQRPath.Text, dataGridView.RowCount.ToString()
            };
            string[] defaultValues = new string[]
            {
                "titel", "titel", "naam", "naam", "plaats", "plaats",
                "doel 1", "doel 3", "naam", "", "1"
            };
            
            for (int i = 0; i < inputValues.Length; i++)
            {
                if (inputValues[i] == defaultValues[i] || string.IsNullOrWhiteSpace(inputValues[i]))
                {
                    invalidInputs.Add(fieldNames[i]);
                }
            }

            if (invalidInputs.Count > 0)
            {
                if (invalidInputs.Count == 1) warning.Append("Het volgende veld is nog niet ingevuld: ");
                else warning.Append("De volgende velden zijn nog niet ingevuld:\r\n    - ");
                warning.Append(string.Join("\r\n    - ", invalidInputs));
                warning.Append("\r\n\r\nWil je doorgaan met het maken van de presentaties?");

                DialogResult result = MessageBox.Show(warning.ToString(), "Waarschuwing", 
                    MessageBoxButtons.YesNo, MessageBoxIcon.Warning);
                return result == DialogResult.Yes;
            }

            return true;
        }

        public void CreatePresentations(object sender, EventArgs e)
        {
            if (!CheckValidInputs()) return;

            // TODO: methods should probably return key and value lists, easier to merge
            Dictionary<string, string> keywords = GetFormKeywords();
            keywords = keywords.Concat(GetSongKeywords()).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);
            keywords = keywords.Concat(GetReadingKeywords()).ToDictionary(kvp => kvp.Key, kvp => kvp.Value);

            foreach (var kvp in keywords)
            {
                Console.WriteLine($"{kvp.Key} = {kvp.Value}");
            }
        }
    }
}
