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
            Settings settings = Settings.Instance;

            textBoxTemplateBefore.Text = settings.PathTemplateBefore;
            textBoxTemplateDuring.Text = settings.PathTemplateDuring;
            textBoxTemplateAfter.Text = settings.PathTemplateAfter;
            textBoxJsonServices.Text = settings.PathServicesJson;
            textBoxOutputFolder.Text = settings.PathOutputFolder;

            checkBoxQRedit.Checked = settings.EnableEditQR;
            checkBoxAutoPopulate.Checked = settings.EnableAutoPopulate;

            numericXMinQR.Value = settings.ImageParameters.OffsetX;
            numericYMinQR.Value = settings.ImageParameters.OffsetY;
            numericXMaxQR.Value = settings.ImageParameters.OffsetX + settings.ImageParameters.Width;
            numericYMaxQR.Value = settings.ImageParameters.OffsetY + settings.ImageParameters.Height;

            if (Settings.Instance.NextService != DateTime.MinValue)
                dateTimePickerCurrent.Value = Settings.Instance.NextService; // triggers the ValueChanged event

            _initialized = true;
        }

        private readonly bool _initialized = false;

        /// <summary>
        /// Select a file and sets this.textBoxQRPath.Text and
        /// QRFile to the path
        /// </summary>
        private void ButtonSelectQR(object sender, EventArgs e)
        {
            string path = Dialogs.SelectFile(
                "Afbeelding (*.jpeg; *.jpg; *.png)|*.jpeg;*.jpg;*.png|Alle bestanden (*.*)|*.*",
                "Selecteer de QR-code"
            );

            if (!string.IsNullOrEmpty(path))
            {
                textBoxQRPath.Text = path;

                checkBoxQRedit.Enabled = true;
                numericXMinQR.Enabled = checkBoxQRedit.Checked;
                numericXMaxQR.Enabled = checkBoxQRedit.Checked;
                numericYMinQR.Enabled = checkBoxQRedit.Checked;
                numericYMaxQR.Enabled = checkBoxQRedit.Checked;
            }
        }

        /// <summary>
        /// Select a directory and sets <see cref="textBoxOutputFolder"/>.Text
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

        /// <summary>
        /// Select a template file for the first presentation
        /// and sets <see cref="Settings.PathTemplateBefore"/>
        /// </summary>
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

        /// <summary>
        /// Select a template file for the second presentation
        /// and sets <see cref="Settings.PathTemplateDuring"/>
        /// </summary>
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

        /// <summary>
        /// Select a template file for the third presentation
        /// and sets <see cref="Settings.PathTemplateAfter"/>
        /// </summary>
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

        /// <summary>
        /// Select the json file for autofilling fields
        /// and sets <see cref="Settings.PathServicesJson"/>
        /// </summary>
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
                Service.UpdateJsonCache();
            }
        }

        /// <summary>
        /// Get the previous service date and updates the datetime picker
        /// </summary>
        private void ButtonPreviousDatetime(object sender, EventArgs e)
        {
            DateTime dateTime = Service.GetPreviousDatetime(dateTimePickerCurrent.Value);
            if (dateTime != DateTime.MinValue) 
                dateTimePickerCurrent.Value = dateTime;
        }

        /// <summary>
        /// Get the next service date and updates the datetime picker
        /// </summary>
        private void ButtonNextDatetime(object sender, EventArgs e)
        {
            DateTime dateTime = Service.GetNextDatetime(dateTimePickerCurrent.Value);
            if (dateTime != DateTime.MinValue)
                dateTimePickerCurrent.Value = dateTime;
        }

        /// <summary>
        /// Autofill fields for the current and next service
        /// if <see cref="Settings.EnableAutoPopulate"/> is true
        /// </summary>
        private void DateTimePickerCurrentChanged(object sender, EventArgs e)
        {
            if (Settings.Instance.EnableAutoPopulate)
            {
                SetFormDataCurrent(Service.GetCurrent(dateTimePickerCurrent.Value));
                SetFormDataNext(Service.GetNext(dateTimePickerCurrent.Value));
            }
        }

        /// <summary>
        /// Autofill fields for the next service
        /// if <see cref="Settings.EnableAutoPopulate"/> is true
        /// </summary>
        private void DateTimePickerNextChanged(object sender, EventArgs e)
        {
            if (Settings.Instance.EnableAutoPopulate)
                SetFormDataNext(Service.GetCurrent(dateTimePickerNext.Value));
        }

        /// <summary>
        /// Autofill fields for the current and next service
        /// </summary>
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

        /// <summary>
        /// Autofill fields for the next service
        /// </summary>
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

        /// <summary>
        /// Retrieve form information and return it as a dictionary of strings, where every key is
        /// a placeholder string in a template and its value the value it should be replaced by
        /// </summary>
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

        /// <summary>
        /// Make the first character in a string uppercase
        /// </summary>
        private static string TitleCase(string input)
        {
            if (string.IsNullOrEmpty(input)) return input;
            return char.ToUpper(input[0]) + input.Substring(1);
        }

        /// <summary>
        /// Convert a DateTimePicker to a H:mm string (24-hour time without leading zero)
        /// </summary>
        private static string GetTime(DateTimePicker dateTimePicker)
        {
            return dateTimePicker.Value.ToString("H:mm", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Convert a DateTimePicker to a date in the long notation without year, e.g. '30 februari'
        /// </summary>
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

        /// <summary>
        /// Add a row to the datagridview
        /// </summary>
        private void ButtonAddDatagridviewRow(object sender, EventArgs e)
        {
            dataGridView.Rows.Add();
        }

        /// <summary>
        /// Remove the selected row(s) from the datagridview
        /// </summary>
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

        /// <summary>
        /// Decrement the current DataGridViewRow's index
        /// </summary>
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

        /// <summary>
        /// Increment the current DataGridViewRow's index
        /// </summary>
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

        /// <summary>
        /// Set the CurrentRow of the datagridview to the one at index <paramref name="index"/>
        /// </summary>
        private void UpdateDatagridviewSelection(int index)
        {
            dataGridView.ClearSelection();
            dataGridView.Rows[index].Selected = true;
            dataGridView.CurrentCell = dataGridView.Rows[index].Cells[0];
        }

        /// <summary>
        /// Color cells of a DataGridViewRow based on the type of service
        /// </summary>
        /// <param name="currentrow"></param>
        /// <param name="serviceElementType"></param>
        private void UpdateDatagridviewRowStyle(DataGridViewRow currentrow, string serviceElementType)
        {
            DataGridViewCellStyle disabledStyle = dataGridView.DefaultCellStyle.Clone();
            disabledStyle.BackColor = Color.LightGray;

            switch (serviceElementType)
            {
                case null:
                    return;
                case "Lezing":
                case "Psalm":
                case "Psalm (WK)":
                    currentrow.Cells[1].Style = dataGridView.DefaultCellStyle;
                    currentrow.Cells[2].Style = disabledStyle;
                    return;
                case "Lied (Overig)":
                    currentrow.Cells[1].Style = disabledStyle;
                    currentrow.Cells[2].Style = dataGridView.DefaultCellStyle;
                    return;
                case "Lied (WK)":
                case "Lied (OTH)":
                default:
                    currentrow.Cells[1].Style = dataGridView.DefaultCellStyle;
                    currentrow.Cells[2].Style = dataGridView.DefaultCellStyle;
                    return;
            }
        }

        /// <summary>
        /// Autofill the song's name if applicable and call <see cref="UpdateDatagridviewRowStyle"/>
        /// </summary>
        private void AutofillDatagridview(object sender, DataGridViewCellEventArgs e)
        {
            // Prevent annoying NRE's
            DataGridViewRow currentrow;
            string serviceElementType;
            try
            {
                currentrow = dataGridView.Rows[e.RowIndex];
                serviceElementType = (string)currentrow.Cells[0].Value;
            }
            catch (NullReferenceException)
            {
                return;
            }
            catch (Exception)
            {
                return;
            }

            // Show which fields have to be filled in
            UpdateDatagridviewRowStyle(currentrow, serviceElementType);

            // Get the song number, convert it to a name and fill it in in the form
            if (serviceElementType != "Lied (WK)") return;
            string songnumber = ((string)dataGridView.CurrentCell.Value).Split(':')[0].Trim();
            string songname = Songnames.GetName(songnumber);
            if (songname is null) return;

            currentrow.Cells[2].Value = songname;
        }

        /// <summary>
        /// Magic, or maybe sufficiently advanced technology.
        /// Prevents certain crashes when accessing the datagridviews dropdown
        /// </summary>
        private void PreventBlackDropdown(object sender, EventArgs e)
        {
            if (dataGridView.CurrentCell is null)
            {
                dataGridView.Rows[0].Cells[0].Selected = true;
            }
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

            if (_initialized)
            {
                numericXMinQR.Enabled = Settings.Instance.EnableEditQR;
                numericYMinQR.Enabled = Settings.Instance.EnableEditQR;
                numericXMaxQR.Enabled = Settings.Instance.EnableEditQR;
                numericYMaxQR.Enabled = Settings.Instance.EnableEditQR;
            }
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
                "collectedoel 1", "collectedoel 3", "naam van de organist", "liturgie"
            };
            string[] inputValues = new string[]
            {
                textBoxVoorgangerNuTitel.Text, textBoxVoorgangerNextTitel.Text,
                textBoxVoorgangerNuNaam.Text, textBoxVoorgangerNextNaam.Text,
                textBoxVoorgangerNuPlaats.Text, textBoxVoorgangerNextPlaats.Text,
                textBoxCollecte1.Text, textBoxCollecte3.Text, textBoxOrganist.Text,
                dataGridView.RowCount.ToString()
            };
            string[] defaultValues = new string[]
            {
                "titel", "titel", "naam", "naam", "plaats", "plaats",
                "doel 1", "doel 3", "naam", "1"
            };

            for (int i = 0; i < inputValues.Length; i++)
                if (inputValues[i] == defaultValues[i] || string.IsNullOrWhiteSpace(inputValues[i]))
                    invalidInputs.Add(fieldNames[i]);

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

            // Create the presentation during the service
            PowerPoint duringService = CreatePowerpoint(Settings.Instance.PathTemplateDuring,
                Settings.Instance.PathOutputFolder + $"/tijdens {filenamepart}.pptx");

            if (duringService == null) return;
            
            duringService.ReplaceImage(textBoxQRPath.Text);

            foreach (ServiceElement element in elements)
            {
                duringService.DuplicateAndReplace(new Dictionary<string, string> {
                    { tags.ServiceElementTitle, element.Title },
                    { tags.ServiceElementSubtitle, element.Subtitle }
                }, element.ShowQR);
            }

            duringService.CopyQRSlideFromPresentation(beforeService);
            duringService.ReplaceKeywords(keywords);
            beforeService.SaveClose();
            duringService.RemoveSlide(index: 0); // delete template slide
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
                Dialogs.GenericWarning($"Het bestand '{ex.FileName}' is niet gevonden." +
                    "Controleer het bestandspad en probeer opnieuw.");
            }
            catch (DirectoryNotFoundException)
            {
                Dialogs.GenericWarning($"Het bestandspad '{templatePath}' of '{outputPath}'" +
                    "bestaat niet. Controleer het bestandspad en probeer opnieuw");
            }
            catch (IOException ex) when ((ex.HResult & 0x0000FFFF) == 32)
            {
                Dialogs.GenericWarning($"Het bestand '{outputPath}' kon niet worden bewerkt omdat " +
                    "het geopend is in een ander programma. Sluit het bestand en probeer opnieuw.");
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

        /// <summary>
        /// Convert <paramref name="dateTimePicker"/> to a yyyy-MM-dd string 
        /// and prefix what part of the day it is
        /// </summary>
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

        /// <summary>
        /// Check if the crop region is valid and save it to the imageparameters
        /// </summary>
        /// <param name="sender"></param>
        /// <param name="e"></param>
        private void CropRegionChanged(object sender, EventArgs e)
        {
            if (numericXMinQR.Value >= numericXMaxQR.Value)
                numericXMaxQR.Value = numericXMinQR.Value + 1;
            if (numericYMinQR.Value >= numericYMaxQR.Value)
                numericYMaxQR.Value = numericYMinQR.Value + 1;

            Settings.Instance.ImageParameters.OffsetX = (int)numericXMinQR.Value;
            Settings.Instance.ImageParameters.OffsetY = (int)numericYMinQR.Value;
            Settings.Instance.ImageParameters.Width  = (int)(numericXMaxQR.Value - numericXMinQR.Value);
            Settings.Instance.ImageParameters.Height = (int)(numericYMaxQR.Value - numericYMinQR.Value);
        }

        private void CropRegionSelectOnFocus(object sender, EventArgs e)
        {
            // for some reason this doesn't happen automatically like with other fields
            ((NumericUpDown)sender).Select(start: 0, length: 4); // max length is 4
        }
    }
}
