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

            // Load settings into the form
            textBoxTemplateBefore.Text = Settings.Instance.PathTemplateBefore;
            textBoxTemplateDuring.Text = Settings.Instance.PathTemplateDuring;
            textBoxTemplateAfter.Text = Settings.Instance.PathTemplateAfter;
            textBoxJsonServices.Text = Settings.Instance.PathServicesJson;
            textBoxOutputFolder.Text = Settings.Instance.PathOutputFolder;
            checkBoxQRedit.Checked = Settings.Instance.EnableEditQR;
            checkBoxAutoPopulate.Checked = Settings.Instance.EnableAutoPopulate;

            if (Settings.Instance.NextService != DateTime.MinValue)
                dateTimePickerCurrent.Value = Settings.Instance.NextService; // automatically triggers the ValueChanged event
        }

        /// <summary>
        /// Selects a file and sets this.textBoxQRPath.Text and
        /// QRFile to the path
        /// </summary>
        private void ButtonSelectQR(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFile(
                "Afbeelding (*.jpeg; *.jpg; *.png)|*.jpeg;*.jpg;*.png|Alle bestanden (*.*)|*.*",
                "Selecteer de QR-code"
            );

            if (!string.IsNullOrEmpty(path)) textBoxQRPath.Text = path;
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
                textBoxTemplateBefore.Text = path;
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

        private void ButtonSelectJsonServices(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFile(
               "JSON (*.json)|*.json",
                "Selecteer het diensten JSON-bestand"
            );

            if (!string.IsNullOrEmpty(path))
            {
                textBoxJsonServices.Text = path;
                Settings.Instance.PathServicesJson = path;
            }
        }

        private void DateTimePickerCurrentChanged(object sender, EventArgs e)
        {
            if (Settings.Instance.EnableAutoPopulate)
            {
                (Service current, Service next) = Service.GetCurrentAndNext(dateTimePickerCurrent.Value);
                SetFormDataCurrent(current);
                SetFormDataNext(next);
            }
        }

        private void DateTimePickerNextChanged(object sender, EventArgs e)
        {
            if (Settings.Instance.EnableAutoPopulate)
            {
                (Service next, _) = Service.GetCurrentAndNext(dateTimePickerNext.Value);
                SetFormDataNext(next);
            }
        }

        private void SetFormDataCurrent(Service service)
        {
            textBoxVoorgangerNuTitel.Text = service.Pastor.Title;
            textBoxVoorgangerNuNaam.Text = service.Pastor.Name;
            textBoxVoorgangerNuPlaats.Text = service.Pastor.Place;

            textBoxCollecte1.Text = service.Collections.First;
            textBoxCollecte2.Text = service.Collections.Second;
            textBoxCollecte3.Text = service.Collections.Third;

            textBoxOrganist.Text = service.Organist;
        }

        private void SetFormDataNext(Service nextService)
        {
            if (nextService.Datetime != DateTime.MinValue)
            {
                dateTimePickerNext.Value = nextService.Datetime;
            }

            textBoxVoorgangerNextTitel.Text = nextService.Pastor.Title;
            textBoxVoorgangerNextNaam.Text = nextService.Pastor.Name;
            textBoxVoorgangerNextPlaats.Text = nextService.Pastor.Place;
        }

        public Dictionary<string, string> GetFormKeywords()
        {
            KeywordSettings tags = Settings.Instance.Keywords;
            Dictionary<string, string> keywords = new Dictionary<string, string>
            {
                { tags.ServiceTime, GetTime(dateTimePickerCurrent) },
                { tags.ServiceNextTime, GetTime(dateTimePickerNext) },
                { tags.ServiceNextDate, GetDateLong(dateTimePickerNext) },
                { tags.Pastor, $"{TitleCase(textBoxVoorgangerNuTitel.Text)} {textBoxVoorgangerNuNaam.Text}" },
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

        private static string TitleCase(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            return char.ToUpper(input[0]) + input.Substring(1);
        }

        private static string GetTime(DateTimePicker dateTimePicker)
        {
            return dateTimePicker.Value.ToString("H:mm", CultureInfo.InvariantCulture);
        }

        private static string GetDateLong(DateTimePicker dateTimePicker)
        {
            // Dictionary to avoid not having the nl-NL resource when using System.Globalization
            Dictionary<int, string> monthNames = new Dictionary<int, string>
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
            // I have no idea why this prevents a crash when using multiple screens at different scaling
            if (index == 1 && dataGridView.CurrentCell is null)
            {
                dataGridView.Rows[0].Cells[1].Selected = true;
            }
            tabControl.SelectTab(index + 1);
        }

        private void ButtonPreviousTab(object sender, EventArgs e)
        {
            int index = tabControl.SelectedIndex;
            tabControl.SelectTab(index - 1);
        }

        private void FocusLeaveSettingsTab(object sender, EventArgs e)
        {
            Settings.Save();
        }

        private void CheckBoxAutoPopulateChanged(object sender, EventArgs e)
        {
            Settings.Instance.EnableAutoPopulate = ((CheckBox)sender).Checked;
        }

        private void CheckBoxEnableEditQRChanged(object sender, EventArgs e)
        {
            Settings.Instance.EnableEditQR = ((CheckBox)sender).Checked;
        }

        public void CheckBoxEnableEditQRSetValue(bool value)
        {
            Settings.Instance.EnableEditQR = value;
            checkBoxQRedit.Checked = value;
        }

        /// <summary>
        /// Get a list of ServiceElements from the DataGridView
        /// </summary>
        private List<ServiceElement> GetServiceElements()
        {
            List<ServiceElement> elements = new List<ServiceElement>();
            foreach (DataGridViewRow row in dataGridView.Rows)
            {
                ServiceElement newElement = new ServiceElement(row);
                if (elements.Count > 0)
                {
                    // If the previous element was a song and this one is a reading, show a QR on the last one
                    if (elements.Last().IsReading && newElement.IsSong) newElement.ShowQR = true;
                }
                elements.Add(newElement);
            }
            elements.RemoveAt(elements.Count - 1); // Last row in the dataframe is a placeholder

            return elements;
        }

        /// <summary>
        /// Check if the values are not the default values, and ask the user if they want to continue
        /// if there are more than one default values present
        /// </summary>
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

        /// <summary>
        /// Create the three presentations with information from the user interface
        /// </summary>
        public void CreatePresentations(object sender, EventArgs e)
        {
            // Check if fields are filled in and if the output folder is right
            if (!CheckValidInputs()) return;
            if (!Directory.Exists(Settings.Instance.PathOutputFolder))
            {
                Dialogs.GenericWarning("De outputfolder bestaat niet, " +
                    "selecteer een bestaande folder in de instellingen");
                return;
            }

            // Get the filled in information from the window
            KeywordSettings tags = Settings.Instance.Keywords;
            Dictionary<string, string> keywords = GetFormKeywords();
            List<ServiceElement> elements = GetServiceElements();
            string filenamepart = GetFilenamePart(dateTimePickerCurrent);

            // Create the presentation before the service
            PowerPoint beforeService = CreatePowerpoint(Settings.Instance.PathTemplateBefore,
                Settings.Instance.PathOutputFolder + $"/voor {filenamepart}.pptx");

            if (beforeService == null) return;
            if (string.IsNullOrWhiteSpace(keywords[tags.Theme]))
                beforeService.RemoveThemeRuns();

            beforeService.ReplaceKeywords(keywords);
            beforeService.ReplaceImage(textBoxQRPath.Text);
            beforeService.ReplaceMultilineKeywords(
                from ServiceElement element in elements where element.IsSong select element,
                from ServiceElement element in elements where element.IsReading select element
            );

            beforeService.SaveClose();

            // Create the presentation during the service
            PowerPoint duringService = CreatePowerpoint(Settings.Instance.PathTemplateDuring,
                Settings.Instance.PathOutputFolder + $"/tijdens {filenamepart}.pptx");

            if (duringService == null) return;

            duringService.ReplaceKeywords(keywords);
            duringService.ReplaceImage(textBoxQRPath.Text);

            foreach (ServiceElement element in elements)
            {
                duringService.DuplicateAndReplace(new Dictionary<string, string> {
                    { tags.ServiceElementTitle, element.Title },
                    { tags.ServiceElementSubtitle, element.Subtitle }
                }, element.ShowQR);
            }

            duringService.SaveClose();

            // Create the presentation after the service
            PowerPoint afterService = CreatePowerpoint(Settings.Instance.PathTemplateAfter,
                Settings.Instance.PathOutputFolder + $"/na {filenamepart}.pptx");

            if (afterService == null) return;

            afterService.ReplaceKeywords(keywords);
            afterService.ReplaceImage(textBoxQRPath.Text);

            afterService.SaveClose();

            // Done, message the user
            Dialogs.GenericInformation("Voltooid", $"De presentaties zijn gemaakt " +
                $"en staan in de folder '{Settings.Instance.PathOutputFolder}'.");
            Settings.Instance.NextService = dateTimePickerNext.Value;
        }

        /// <summary>
        /// Try to create a new <see cref="PowerPoint"/> object, handle exceptions
        /// </summary>
        private static PowerPoint CreatePowerpoint(string templatePath, string outputPath)
        {
            try
            {
                PowerPoint powerpoint = new PowerPoint(templatePath, outputPath);
                return powerpoint;
            }
            catch (FileNotFoundException ex)
            {
                Dialogs.GenericWarning($"Het bestand '{ex.FileName}' is niet gevonden. Controleer het bestandspad en probeer opnieuw.");
            }
            catch (DirectoryNotFoundException)
            {
                Dialogs.GenericWarning($"Het bestandspad '{templatePath}' of '{outputPath}' bestaat niet. Controleer het bestandspad " +
                    $"en probeer opnieuw");
            }
            catch (IOException ex) when ((ex.HResult & 0x0000FFFF) == 32)
            {
                Dialogs.GenericWarning($"Het bestand '{outputPath}' kon niet worden bewerkt omdat het geopend is in een ander " +
                    "programma. Sluit het bestand en probeer opnieuw.");
            }
            catch (Exception ex) when (ex is IOException
                || ex is UnauthorizedAccessException
                || ex is NotSupportedException
                || ex is System.Security.SecurityException)
            {
                Dialogs.GenericWarning($"'{templatePath}' of '{outputPath}' kon niet worden geopend.\n\n" +
                    $"De volgende foutmelding werd gegeven: {ex.Message}");
            }
            return null;
        }

        private static string GetFilenamePart(DateTimePicker dateTimePicker)
        {
            string output = "";
            DateTime dateTime = dateTimePicker.Value;
            if (dateTime.Hour < 12) output += "ochtend ";
            else if (dateTime.Hour < 18) output += "middag ";
            else output += "avond ";
            output += dateTime.ToString("yyyy-MM-dd");
            return output;
        }
    }
}
