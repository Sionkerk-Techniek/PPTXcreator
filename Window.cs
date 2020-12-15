using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

namespace PPTXcreator
{
    public partial class Window : Form
    {
        public string QRFile;

        public Window()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Selects the content of the textbox for quick editing
        /// </summary>
        private void TextboxSelectOnEnter(object sender, EventArgs e)
        {
            TextBox box = (TextBox)sender;
            box.SelectAll();
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
                Filter = "JPEG (*.jpeg)|*.jpeg|PNG (*.png)|*.png|Overig (*.*)|*.*",
                InitialDirectory = Environment.GetFolderPath(Environment.SpecialFolder.MyComputer),
                Title = "Selecteer de QR-code",
                RestoreDirectory = true
            };

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(fileDialog.FileName))
                {
                    this.textBoxQRPath.Text = fileDialog.FileName;
                    QRFile = fileDialog.FileName;
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
        /// Selects a directory and sets this.textBoxOutputFolder.Text
        /// and Outputfolder to the path
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
                    this.textBoxOutputFolder.Text = folderDialog.SelectedPath;
                    Settings.OutputFolder = folderDialog.SelectedPath;
                    Settings.ChangeSetting("OUTPUT_FOLDER", folderDialog.SelectedPath);
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
    }
}
