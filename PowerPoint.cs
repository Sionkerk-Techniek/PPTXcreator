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
        public static Dictionary<string, string> Keywords { get; set; }
        
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
