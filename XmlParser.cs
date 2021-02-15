using System;
using System.Collections.Generic;
using System.IO;
using System.Xml;
using System.Windows.Forms;

namespace PPTXcreator
{
    static class XmlParser
    {
        /// <summary>
        /// Gets the XmlNode from <paramref name="filePath"/> containing service or
        /// organist information about the service on <paramref name="dateTimeString"/>.
        /// </summary>
        /// <param name="filePath">Path to the XML file</param>
        /// <param name="dateTimeString">DateTime formatted as "yyyy-MM-dd H:mm"</param>
        /// <param name="xpath">The path to the root object in the XML file</param>
        /// <returns>The relevant <see cref="XmlNode"/> if the service was found, null otherwise</returns>
        public static XmlNode GetNode(string filePath, string dateTimeString, string xpath)
        {
            if (filePath is null || !File.Exists(filePath)) return null;
            // Load the XML file and get the node containing all services
            XmlNode root = Load(filePath, xpath);
            if (CheckNodeIsNull(root, filePath, xpath)) return null;

            // Iterate over all child nodes to find the node with the right datetime value
            foreach (XmlNode node in root.ChildNodes)
            {
                if (node.SelectSingleNode("datetime").InnerText == dateTimeString)
                {
                    return node;
                }
            }

            return null;
        }

        /// <summary>
        /// Loads an xml file from <paramref name="filePath"/>
        /// and returns the node at <paramref name="xpath"/>
        /// </summary>
        private static XmlNode Load(string filePath, string xpath)
        {
            XmlDocument document = new XmlDocument();
            try
            {
                document.Load(filePath);
                return document.SelectSingleNode(xpath);
            }
            catch (IOException)
            {
                return null;
            }
            catch (XmlException)
            {
                return null;
            }
        }

        private static bool CheckNodeIsNull(XmlNode node, string filePath, string xpath)
        {
            if (node is null)
            {
                MessageBox.Show("Het geselecteerde XML-bestand heeft niet de juiste interne structuur: "
                    + $"'{xpath}' is niet gevonden in het bestand '{filePath}'.",
                    "Er is een fout opgetreden",
                    MessageBoxButtons.OK,
                    MessageBoxIcon.Error);
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}
