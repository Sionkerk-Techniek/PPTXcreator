using System;
using System.Collections.Generic;
using System.Linq;
using System.Xml;
using System.Globalization;

namespace PPTXcreator
{
    public class Service
    {
        // TODO: create method which updates properties when Time or NextTime change
        public DateTime Time { get; }
        public DateTime NextTime { get; }
        public string DsName { get; private set; }
        public string NextDsName { get; private set; }
        public string DsPlace { get; private set; }
        public string NextDsPlace { get; private set; }
        public string Collection_1 { get; private set; }
        public string Collection_3 { get; private set; }
        public string Organist { get; private set; }

        /// <summary>
        /// Data class containing information about a service
        /// </summary>
        /// <param name="serviceNode">The XmlNode object containing information about the service</param>
        /// <param name="organistNode">The XmlNode object containing information about the organist</param>
        public Service(XmlNode serviceNode, XmlNode organistNode)
        {
            if (!(serviceNode is null))
            {
                // Get the next service
                XmlNode nextServiceNode = serviceNode.NextSibling;

                (DsName, DsPlace) = GetDsAttributes(serviceNode.SelectSingleNode("voorganger"));
                if (!(nextServiceNode is null))
                {
                    // Fill in GUI textboxes with the xml info
                    (NextDsName, NextDsPlace) = GetDsAttributes(nextServiceNode.SelectSingleNode("voorganger"));
                    (Time, NextTime) = GetServiceTimes(serviceNode, nextServiceNode);
                }
                else
                {
                    // Set the values for the next service to default if the next service is not found
                    (NextDsName, NextDsPlace) = ("titel naam", "plaats");
                    (Time, NextTime) = GetServiceTimes(serviceNode, serviceNode);
                }

                Collection_1 = serviceNode.SelectSingleNode("collecte_1").InnerText;
                Collection_3 = serviceNode.SelectSingleNode("collecte_3").InnerText;
            }
            else
            {
                // Reset to default values if the node is not found in the form

            }

            if (!(organistNode is null))
            {
                Organist = organistNode.SelectSingleNode("organist").InnerText;
            }
        }

        private void ResetFields(bool onlyNextService = false)
        {
            DsName = NextDsName = "titel naam";
            DsPlace = NextDsPlace = "plaats";
            Collection_1 = "doel 1";
            Collection_3 = "doel 3";
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
        /// Get the time of this service and the next service from the relevant XmlNodes
        /// </summary>
        /// <returns></returns>
        private (DateTime, DateTime) GetServiceTimes(XmlNode serviceNode, XmlNode nextServiceNode)
        {
            DateTime time = DateTime.ParseExact(serviceNode.SelectSingleNode("datetime").InnerText,
                "yyyy-MM-dd H:mm", CultureInfo.InvariantCulture);
            DateTime nextTime = DateTime.ParseExact(nextServiceNode.SelectSingleNode("datetime").InnerText,
                "yyyy-MM-dd H:mm", CultureInfo.InvariantCulture);
            return (time, nextTime);
        }

        /// <summary>
        /// Takes the title and name as a single string and splits it
        /// </summary>
        public static (string, string) SplitName(string dsTitleName)
        {
            if (string.IsNullOrWhiteSpace(dsTitleName)) return ("titel", "naam");
            else if (!dsTitleName.Contains(" ")) return ("titel", "naam");

            string[] titleName = dsTitleName.Split(" ".ToCharArray(), 2);
            return (titleName[0], titleName[1]);
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
            {
                { 1, "januari" }, { 2, "februari" }, { 3, "maart" }, { 4, "april" },
                { 5, "mei" }, { 6, "juni" }, { 7, "juli" }, { 8, "augustus"},
                { 9, "september" }, { 10, "oktober" }, { 11, "november" }, { 12, "december" }
            };

            return $"{timeProperty.Day} {monthNames[Time.Month]}";
        }
    }
}
