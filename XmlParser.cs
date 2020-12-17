using System;
using System.Collections.Generic;
using System.Globalization;
using System.IO;
using System.Xml;

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
            if (filePath is null) return null;
            // Load the XML file and get the node containing all services
            XmlNode root = Load(filePath, xpath);

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
        }
    }

    public class Service
    {
        public DateTime Time { get; }
        public DateTime NextTime { get; }
        public string DsName { get; }
        public string NextDsName { get; }
        public string DsPlace { get; }
        public string NextDsPlace { get; }
        public string Collecte_1 { get; }
        public string Collecte_3 { get; }
        public string Organist { get; }

        /// <summary>
        /// Data class containing information about a service
        /// </summary>
        /// <param name="serviceNode">The XmlNode object containing information about the service</param>
        /// <param name="organistNode">The XmlNode object containing information about the organist</param>
        public Service(XmlNode serviceNode, XmlNode organistNode)
        {
            if (!(serviceNode is null))
            {
                XmlNode nextServiceNode = serviceNode.NextSibling;

                (DsName, DsPlace) = GetDsAttributes(serviceNode.SelectSingleNode("voorganger"));
                (NextDsName, NextDsPlace) = GetDsAttributes(nextServiceNode.SelectSingleNode("voorganger"));

                Time = DateTime.ParseExact(serviceNode.SelectSingleNode("datetime").InnerText,
                    "yyyy-MM-dd H:mm", CultureInfo.InvariantCulture);
                NextTime = DateTime.ParseExact(nextServiceNode.SelectSingleNode("datetime").InnerText,
                    "yyyy-MM-dd H:mm", CultureInfo.InvariantCulture);

                Collecte_1 = serviceNode.SelectSingleNode("collecte_1").InnerText;
                Collecte_3 = serviceNode.SelectSingleNode("collecte_3").InnerText;
            }
            if (!(organistNode is null))
            {
                Organist = organistNode.SelectSingleNode("organist").InnerText;
            }
        }

        /// <summary>
        /// Builds a Service instance for the relevant service at <paramref name="dateTimeString"/>
        /// from the XML files located at <paramref name="infoXML"/> and <paramref name="organistXML"/>
        /// </summary>
        public static Service GetService(string infoXML, string organistXML, string dateTimeString)
        {
            XmlNode dsNode = XmlParser.GetNode(infoXML, dateTimeString, "diensten");
            XmlNode organistNode = XmlParser.GetNode(organistXML, dateTimeString, "organisten");
            return new Service(dsNode, organistNode);
        }

        /// <summary>
        /// Returns the 'naam' and 'plaats' attribute values for the XmlNode <paramref name="node"/>
        /// </summary>
        private (string, string) GetDsAttributes(XmlNode node)
        {
            XmlAttributeCollection dsInfo = node.Attributes;
            string ds = dsInfo.GetNamedItem("naam").Value;
            string dsPlace = dsInfo.GetNamedItem("plaats").Value;
            return (ds, dsPlace);
        }

        /// <summary>
        /// Returns the time in 'H:mm' notation
        /// </summary>
        public static string GetTime(DateTime timeProperty)
        {
            return timeProperty.ToString("H:mm", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Returns the date in 'yyyy-MM-dd' notation
        /// </summary>
        public string GetDate(DateTime timeProperty)
        {
            return timeProperty.ToString("yyyy-MM-dd", CultureInfo.InvariantCulture);
        }

        /// <summary>
        /// Returns the date in 'd MMMM' notation
        /// </summary>
        public string GetDateLong(DateTime timeProperty)
        {
            // Dictionary to avoid not having the nl-NL resource when using System.Globalization
            Dictionary<int, string> monthNames = new Dictionary<int, string>()
            {   { 1, "januari" }, { 2, "februari" }, { 3, "maart" }, { 4, "april" },
                { 5, "mei" }, { 6, "juni" }, { 7, "juli" }, { 8, "augustus"},
                { 9, "september" }, { 10, "oktober" }, { 11, "november" }, { 12, "december" } };

            return $"{timeProperty.Day} {monthNames[Time.Month]}";
        }
    }
}
