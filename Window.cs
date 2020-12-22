﻿using System;
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
                    this.textBoxQRPath.Text = fileDialog.FileName;
                    Settings.ImagePath = fileDialog.FileName;
                }
                else
                {
                    MessageBox.Show(
                        "Het geselecteerde bestand kan niet worden gevonden",
                        "Er is een fout opgetreden",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
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
                    textBoxOutputFolder.Text = folderDialog.SelectedPath;
                    Settings.OutputFolderPath = folderDialog.SelectedPath;
                }
                else
                {
                    MessageBox.Show(
                        "De geselecteerde folder kan niet worden gevonden",
                        "Er is een fout opgetreden",
                        MessageBoxButtons.OK,
                        MessageBoxIcon.Warning
                    );
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
                textBoxTemplatePre.Text = fileDialog.FileName;
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
                textBoxTemplateDuring.Text = fileDialog.FileName;
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
                textBoxTemplateAfter.Text = fileDialog.FileName;
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
                textBoxXmlServices.Text = fileDialog.FileName;
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
                textBoxXmlOrganist.Text = fileDialog.FileName;
                Settings.OrganistXml = fileDialog.FileName;
            }
        }

        private void DateTimePickerNu_Leave(object sender, EventArgs e)
        {
            string dateTimeString = dateTimePickerNu.Value.ToString("yyyy-MM-dd H:mm", CultureInfo.InvariantCulture);
            if (Settings.AutoPopulate)
            {
                Service currentService = Service.GetService(Settings.ServicesXml, Settings.OrganistXml, dateTimeString);
                UpdateForm(currentService);
            }
        }

        private void UpdateForm(Service service)
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
    }
}
