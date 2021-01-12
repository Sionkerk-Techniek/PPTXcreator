using System;
using System.IO;
using System.Collections.Generic;
using System.Linq; // for ToArray()
using System.Text;
using DocumentFormat.OpenXml;
using Presentation = DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

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
            if (!File.Exists(imagePath)) return;

            // Loop over slides (not really necessary if the slide number is known)
            foreach (SlidePart slide in Slides)
            {
                // Loop over all picture objects
                foreach (Presentation.Picture pic in slide.Slide.Descendants<Presentation.Picture>())
                {
                    // Get the description and rId from the object
                    string description = pic.NonVisualPictureProperties.NonVisualDrawingProperties.Description;
                    string rId = pic.BlipFill.Blip.Embed.Value;
                    Console.WriteLine($"rid: {rId}, description: {description}");

                    if (description == Settings.QRdescription)
                    {
                        // Get the ImagePart by id, and replace the image
                        ImagePart imagePart = (ImagePart)slide.GetPartById(rId);
                        FileStream imageStream = File.OpenRead(imagePath);
                        imagePart.FeedData(imageStream);
                        imageStream.Close();
                    }
                }
            }
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
