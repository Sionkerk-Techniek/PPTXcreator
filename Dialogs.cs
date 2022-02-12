﻿using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace PPTXcreator
{
    public static class Dialogs
    {
        public static void GenericWarning(string message)
        {
            MessageBox.Show(
                message,
                "Er is een fout opgetreden",
                MessageBoxButtons.OK,
                MessageBoxIcon.Warning
            );
        }

        public static void CrashNotification(string message)
        {
            MessageBox.Show(
                "De volgende foutmelding werd gegeven: " + message,
                "Het programma is gestopt",
                MessageBoxButtons.OK,
                MessageBoxIcon.Error
            );
        }

        public static void GenericInformation(string title, string message)
        {
            MessageBox.Show(
                message,
                title,
                MessageBoxButtons.OK,
                MessageBoxIcon.Information
            );
        }

        /// <summary>
        /// Show a folder selection window and check if the selected folder exists
        /// </summary>
        /// <returns>The path to the selected folder, or null if the path is invalid</returns>
        public static string SelectFolder(string description)
        {
            FolderBrowserDialog folderDialog = new FolderBrowserDialog
            {
                Description = description,
            };

            if (folderDialog.ShowDialog() == DialogResult.OK)
            {
                if (Directory.Exists(folderDialog.SelectedPath))
                {
                    return Settings.GetPath(folderDialog.SelectedPath);
                }
                else
                {
                    GenericWarning("De geselecteerde folder kan niet worden gevonden");
                }
            }
            
            return null;
        }

        /// <summary>
        /// Show a file selection window and check if the selected file exists
        /// </summary>
        /// <returns>The path to the selected file, or null if the path is invalid</returns>
        public static string SelectFile(string filter, string title)
        {
            OpenFileDialog fileDialog = new OpenFileDialog
            {
                Filter = filter,
                InitialDirectory = Directory.GetCurrentDirectory(),
                Title = title,
                RestoreDirectory = true,
            };

            if (fileDialog.ShowDialog() == DialogResult.OK)
            {
                if (File.Exists(fileDialog.FileName))
                {
                    return Settings.GetPath(fileDialog.FileName);
                }
                else
                {
                    GenericWarning("Het geselecteerde bestand kan niet worden gevonden");
                }
            }

            return null;
        }

        public static void UpdateAvailable(string newversion)
        {
            DialogResult result = MessageBox.Show(
                $"Er is een update beschikbaar: versie {newversion}. Wil je deze nu downloaden?",
                "Update beschikbaar",
                MessageBoxButtons.YesNo,
                MessageBoxIcon.Information
            );

            if (result == DialogResult.Yes)
            {
                System.Diagnostics.Process.Start("https://github.com/Sionkerk-Techniek/PPTXcreator/releases/latest");
            }
        }
    }
}
