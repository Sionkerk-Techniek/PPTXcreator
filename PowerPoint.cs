using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace PPTXcreator
{
    class PowerPoint
    {
        private PresentationDocument Document { get; }
        private PresentationPart PresPart { get => Document.PresentationPart; }
        private IEnumerable<SlidePart> Slides { get => PresPart.SlideParts; }
        
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
        /// Replaces all keywords in a slide with their respective values
        /// </summary>
        /// <param name="keywords">A dictionary containing replaceable strings
        /// and what they should be replaced by</param>
        /// <param name="slide">The slide the keywords have to be replaced in</param>
        private void ReplaceKeywords(Dictionary<string, string> keywords, Slide slide)
        {
            // Loop over text in the slide
            foreach (Drawing.Text text in slide.Descendants<Drawing.Text>())
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

        /// <summary>
        /// Replaces all keywords in the presentation with their respective values
        /// </summary>
        /// <param name="keywords">A dictionary containing replaceable strings
        /// and what they should be replaced by</param>
        public void ReplaceKeywords(Dictionary<string, string> keywords)
        {
            // Loop over slides
            foreach (SlidePart slidePart in Slides)
            {
                ReplaceKeywords(keywords, slidePart.Slide);
            }
        }

        /// <summary>
        /// Replaces the image with description <see cref="Settings.QRdescription"/>
        /// with the image at <paramref name="imagePath"/>.
        /// </summary>
        /// <param name="imagePath">Path to the image</param>
        /// <param name="slidePart">The slidePart the image has to be replaced in</param>
        private void ReplaceImage(string imagePath, SlidePart slidePart)
        {
            // Loop over all picture objects
            foreach (Picture pic in slidePart.Slide.Descendants<Picture>())
            {
                // Get the description and rId from the object
                string description = pic.NonVisualPictureProperties.NonVisualDrawingProperties.Description;
                string rId = pic.BlipFill.Blip.Embed.Value;
                Console.WriteLine($"rid: {rId}, description: {description}");

                if (description == Settings.QRdescription)
                {
                    // Get the ImagePart by id, and replace the image
                    ImagePart imagePart = (ImagePart)slidePart.GetPartById(rId);
                    FileStream imageStream = File.OpenRead(imagePath);
                    imagePart.FeedData(imageStream);
                    imageStream.Close();
                }
            }
        }

        /// <summary>
        /// Replaces the image with description <see cref="Settings.QRdescription"/>
        /// with the image at <paramref name="imagePath"/>.
        /// </summary>
        public void ReplaceImage(string imagePath)
        {
            if (!File.Exists(imagePath)) return;

            // Loop over slideparts (not really necessary if the slide number is known)
            foreach (SlidePart slidePart in Slides)
            {
                ReplaceImage(imagePath, slidePart);
            }
        }

        /// <summary>
        /// Copy the first slide of this document and paste it at the end,
        /// and make the copied slide visible if it was hidden
        /// </summary>
        public SlidePart DuplicateFirstSlide()
        {
            // Get the SlideIdList and the largest id in it
            SlideIdList idList = PresPart.Presentation.SlideIdList;
            uint maxId = (from slideId in idList select ((SlideId)slideId).Id).Max();

            // Get the first SlidePart from the SlideIdList
            SlideId sourceSlideId = (SlideId)idList.FirstChild;
            SlidePart sourceSlidePart = (SlidePart)PresPart.GetPartById(sourceSlideId.RelationshipId);

            // Copy the slide and SlideLayoutPart to a new slidepart, set it to visible
            SlidePart targetSlidePart = PresPart.AddNewPart<SlidePart>();
            sourceSlidePart.Slide.Save(targetSlidePart);
            targetSlidePart.AddPart(sourceSlidePart.SlideLayoutPart);
            targetSlidePart.Slide.Show = true;

            // Append a new id for the slidepart to the SlideIdList
            SlideId targetSlideId = idList.AppendChild(new SlideId());
            targetSlideId.Id = maxId + 1;
            targetSlideId.RelationshipId = PresPart.GetIdOfPart(targetSlidePart);

            PresPart.Presentation.Save();
            return targetSlidePart;
        }

        /// <summary>
        /// Duplicates the first slide, then calls
        /// <see cref="ReplaceKeywords(Dictionary{string, string}, Slide)"/>
        /// </summary>
        /// <param name="keywords">A dictionary containing replaceable strings
        /// and what they should be replaced by</param>
        public void DuplicateAndReplace(Dictionary<string, string> keywords)
        {
            SlidePart slidePart = DuplicateFirstSlide();
            ReplaceKeywords(keywords, slidePart.Slide);
            // ReplaceImage can also be used if necessary
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
