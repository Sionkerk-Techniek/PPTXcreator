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

        /// <summary>
        /// Selects a file and sets this.textBoxQRPath.Text and
        /// QRFile to the path
        /// </summary>
        private void ButtonSelectQR(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFile(
                "JPEG (*.jpeg)|*.jpeg|PNG (*.png)|*.png|Alle bestanden (*.*)|*.*",
                "Selecteer de QR-code"
            );

            if (!string.IsNullOrEmpty(path))
            {
                textBoxQRPath.Text = path;
                Settings.Instance.PathQRImage = path;
            }
        }

        /// <summary>
        /// Selects a directory and sets <see cref="textBoxOutputFolder"/>.Text
        /// and <see cref="Settings.OutputFolder"/> to the path
        /// </summary>
        private void ButtonSelectOutputFolder(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFolder("Selecteer de output folder");
            if (!string.IsNullOrEmpty(path))
            {
                textBoxOutputFolder.Text = path;
                Settings.Instance.PathOutputFolder = path;
            }
        }

        private void ButtonSelectTemplateBefore(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFile(
                "PowerPoint (*.pptx)|*.pptx",
                "Selecteer de 'voor de dienst' template"
            );
            if (!string.IsNullOrEmpty(path))
            {
                textBoxTemplatePre.Text = path;
                Settings.Instance.PathTemplateBefore = path;
            }
        }

        private void ButtonSelectTemplateDuring(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFile(
                "PowerPoint (*.pptx)|*.pptx",
                "Selecteer de 'tijdens de dienst' template"
            );
            if (!string.IsNullOrEmpty(path))
            {
                textBoxTemplateDuring.Text = path;
                Settings.Instance.PathTemplateDuring = path;
            }
        }

        private void ButtonSelectTemplateAfter(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFile(
                "PowerPoint (*.pptx)|*.pptx",
                "Selecteer de 'voor de dienst' template"
            );
            if (!string.IsNullOrEmpty(path))
            {
                textBoxTemplateAfter.Text = path;
                Settings.Instance.PathTemplateAfter = path;
            }
        }

        private void ButtonSelectXmlServices(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFile(
               "XML (*.xml)|*.xml",
                "Selecteer het diensten XML-bestand"
            );

            if (!string.IsNullOrEmpty(path))
            {
                textBoxXmlServices.Text = path;
                Settings.Instance.PathServicesJson = path;
            }
        }

        private void ButtonSelectXmlOrganists(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFile(
               "XML (*.xml)|*.xml",
                "Selecteer het organisten XML-bestand"
            );

            if (!string.IsNullOrEmpty(path))
            {
                textBoxXmlOrganist.Text = path;
                Settings.Instance.PathOrganistsJson = path;
            }
        }

        private void DateTimePickerNu_Leave(object sender, EventArgs e)
        {
            string dateTimeString = dateTimePickerNu.Value.ToString("yyyy-MM-dd H:mm",
                CultureInfo.InvariantCulture);

            if (Settings.Instance.EnableAutoPopulate)
            {
                Service currentService = Service.GetService(Settings.Instance.PathServicesJson,
                    Settings.Instance.PathOrganistsJson, dateTimeString);

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
            KeywordSettings tags = Settings.Instance.Keywords;
            Dictionary<string, string> keywords = new Dictionary<string, string>
            {
                { tags.ServiceTime, GetTime(dateTimePickerNu) },
                { tags.ServiceNextTime, GetTime(dateTimePickerNext) },
                { tags.ServiceNextDate, GetDateLong(dateTimePickerNext) },
                { tags.Pastor, $"{textBoxVoorgangerNuTitel.Text} {textBoxVoorgangerNuNaam.Text}" },
                { tags.PastorPlace, textBoxVoorgangerNuPlaats.Text },
                { tags.PastorNext, $"{textBoxVoorgangerNextTitel.Text} {textBoxVoorgangerNextNaam.Text}" },
                { tags.PastorNextPlace, textBoxVoorgangerNextPlaats.Text },
                { tags.Organist, textBoxOrganist.Text },
                { tags.Theme, textBoxThema.Text },
                { tags.Collection1, textBoxCollecte1.Text },
                { tags.Collection2, textBoxCollecte2.Text },
                { tags.Collection3, textBoxCollecte3.Text }
            };

            return keywords;
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
                    Dialogs.GenericWarning("De rij kan niet worden verwijderd." +
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
                    Dialogs.GenericWarning("De rij kan niet worden verwijderd. " +
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

            Dictionary<string, string> keywords = GetFormKeywords();
            List<ServiceElement> elements = new List<ServiceElement>();
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                elements.Add(new ServiceElement(row));
            }

            PowerPoint beforeService = CreatePowerpoint(Settings.Instance.PathTemplateBefore, Settings.Instance.PathOutputFolder + "/outputbefore.pptx");
            if (beforeService == null) return;
            beforeService.ReplaceKeywords(keywords);
            beforeService.ReplaceImage(textBoxQRPath.Text);
            beforeService.ReplaceMultilineKeywords(
                (from ServiceElement element in elements where element.IsSong select element).ToList(),
                (from ServiceElement element in elements where !element.IsSong select element).ToList()
            );
            beforeService.SaveClose();

            PowerPoint duringService = CreatePowerpoint(Settings.Instance.PathTemplateDuring, Settings.Instance.PathOutputFolder + "/outputduring.pptx");
            if (duringService == null) return;
            duringService.ReplaceKeywords(keywords);
            duringService.ReplaceImage(textBoxQRPath.Text);
            KeywordSettings tags = Settings.Instance.Keywords;
            foreach (ServiceElement element in elements)
            {
                duringService.DuplicateAndReplace(new Dictionary<string, string> {
                    { tags.ServiceElementTitle, element.Title },
                    { tags.ServiceElementSubtitle, element.Subtitle }
                });
            }
            duringService.SaveClose();

            PowerPoint afterService = CreatePowerpoint(Settings.Instance.PathTemplateAfter, Settings.Instance.PathOutputFolder + "/outputafter.pptx");
            if (afterService == null) return;
            afterService.ReplaceKeywords(keywords);
            afterService.ReplaceImage(textBoxQRPath.Text);
            afterService.SaveClose();
        }

        private PowerPoint CreatePowerpoint(string templatePath, string outputPath)
        {
            try
            {
                PowerPoint powerpoint = new PowerPoint(templatePath, outputPath);
                return powerpoint;
            }
            catch (IOException)
            {
                Dialogs.GenericWarning("Niet alle presentaties konden worden gemaakt omdat een of meer " +
                    "bestanden op de outputlocatie geopend zijn. Sluit PowerPoint en probeer het opnieuw.");
                return null;
            }
        }
    }
}
