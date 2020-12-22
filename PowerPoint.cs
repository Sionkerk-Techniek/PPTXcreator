using System;
using System.IO;
using System.Collections.Generic;
using System.Linq; // for ToArray()
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Drawing.Pictures;
using DocumentFormat.OpenXml.Packaging;
using System.Reflection;

namespace PPTXcreator
{
    class PowerPoint
    {
        private PresentationDocument Document { get; }
        private SlidePart[] Slides { get => Document.PresentationPart.SlideParts.ToArray(); }
        public Dictionary<string, string> Keywords = new Dictionary<string, string>()
        {
            { "#TIJD_1", "13:37" },
            { "#TIJD_2", "16:15" },
            { "#DS_PLAATS_1", "Null Island" },
            { "#DS_NAAM_1", "ds. Voornaam Achternaam" },
            { "#DS_PLAATS_2", "Null Island" },
            { "#DS_NAAM_2", "ds. Voornaam Achternaam" },
            { "#ORGANIST", "Abc van de Defgijklmnopqrstuvwxyz" },
            { "#ZINGEN_1", "Ps 151 : 1, 2 en 3" },
            { "#ZINGEN_2", "Ps 152 : 1, 2 en 3" },
            { "#ZINGEN_3", "Ps 153 : 1, 2 en 3 WK" },
            { "#ZINGEN_4", "Ps 154 : 1, 2 en 3" },
            { "#ZINGEN_5", "Ps 155 : 1, 2 en 3" },
            { "#ZINGEN_6", "Ps 156 : 1, 2 en 3" },
            { "#ZINGEN_7", "Ps 157 : 1, 2 en 3 WK" },
            { "#LEZING", "asdf" },
            { "#THEMA", "testthema" },
            { "#COLLECTE_1", "doel 1" },
            { "#COLLECTE_2", "doel 2" }
        };
        
        /// <summary>
        /// Constructs a Presentation object from the .pptx file at <paramref name="openPath"/>,
        /// and sets it to save to <paramref name="savePath"/>.
        /// </summary>
        public PowerPoint(string openPath, string savePath)
        {
            PresentationDocument document = PresentationDocument.Open(openPath, true);
            Document = (PresentationDocument)document.Clone(savePath, true);
        }

        /// <summary>
        /// Replaces all keywords in the presentation with their respective values
        /// </summary>
        /// <param name="keywords">A dictionary containing replaceable strings
        /// and what they should be replaced by</param>
        public void ReplaceKeywords(Dictionary<string, string> keywords)
        {
            // Loop over slides
            foreach (SlidePart slide in Slides)
            {
                // Loop over text in slides
                foreach (Drawing.Text text in slide.Slide.Descendants<Drawing.Text>())
                {
                    StringBuilder sb = new StringBuilder(text.Text);

                    // Loop over replacable keywords
                    foreach (KeyValuePair<string, string> kvp in keywords)
                    {
                        sb.Replace(kvp.Key, kvp.Value);
                    }

                    text.Text = sb.ToString();
                }
            }
        }

        /// <summary>
        /// Replaces the image located at "/ppt/media/image4.png"
        /// with the image at <paramref name="imagePath"/>.
        /// </summary>
        public void ReplaceImage(string imagePath)
        {
            // Loop over slides (not really necessary if the slide number is known)
            foreach (SlidePart slide in Slides)
            {
                // Loop over all ImageParts in the slide
                foreach (ImagePart image in slide.ImageParts)
                {
                    // TODO: write all URIs to logfile so they can be easily known by the user

                    // Check if the image's URI is equal to a known URI
                    if (image.Uri.ToString() == "/ppt/media/image44.png")
                    {
                        FileStream imageStream = File.OpenRead(imagePath);
                        image.FeedData(imageStream);
                        imageStream.Close();
                    }
                }
            }
            // TODO?: check image dimensions maybe instead of a URI
            // slide.GetIdOfPart(image) -> possibly useful
        }

        /// <summary>
        /// Save and close the presentation
        /// </summary>
        public void SaveClose()
        {
            Document.Close();
            Document.Dispose(); // Is this necessary? Probably not but it feels appropiate
        }
    }
}
