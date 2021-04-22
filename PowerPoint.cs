using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using Presentation = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
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
            using (PresentationDocument document = PresentationDocument.Open(openPath, false))
            {
                Document = (PresentationDocument)document.Clone(savePath, true);
            }
        }

        /// <summary>
        /// Replaces all keywords in a slide with their respective values
        /// </summary>
        /// <param name="keywords">A dictionary containing replaceable strings
        /// and what they should be replaced by</param>
        /// <param name="slide">The slide the keywords have to be replaced in</param>
        private void ReplaceKeywords(Dictionary<string, string> keywords, Presentation.Slide slide)
        {
            // Loop over text in the slide
            foreach (Text text in slide.Descendants<Text>())
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
        /// Replaces all keywords in the presentation with their respective values,
        /// except for '[zingen]' and '[lezen]'
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
        /// Replace a <see cref="Run"/> element and the paragraph it's in
        /// with multiple <see cref="Paragraph"/> objects
        /// </summary>
        /// <param name="run">The run to replace</param>
        /// <param name="elements">The ServiceElements to replace the run by</param>
        public void ReplaceMultilineKeywords(Run run, IEnumerable<ServiceElement> elements)
        {
            // Get the paragraph and textbody the run is a child of
            Paragraph par = (Paragraph)run.Parent;
            Presentation.TextBody textBody = (Presentation.TextBody)par.Parent;
            OpenXmlElement lastPar = par;

            // Add a new run to the paragraph for every line in the input list
            foreach (ServiceElement element in elements)
            {
                string text;
                if (element.Type == ElementType.PsalmWK) text = element.Title + "   WK";
                else text = element.Title;

                // A new deep copy of the runproperties is needed for every new run
                RunProperties runProperties = (RunProperties)run.RunProperties.CloneNode(true);
                ParagraphProperties parProperties = (ParagraphProperties)par.ParagraphProperties.CloneNode(true);
                Paragraph newPar = new Paragraph(
                    new Run(
                        new Text(text)
                    ) { RunProperties = runProperties }
                ) { ParagraphProperties = parProperties };
                lastPar = textBody.InsertAfter(newPar, lastPar);
            }

            // Remove the placeholder
            par.Remove();
        }

        /// <summary>
        /// Replace placeholder text in the document with songs and readings
        /// </summary>
        public void ReplaceMultilineKeywords(IEnumerable<ServiceElement> songs, IEnumerable<ServiceElement> readings)
        {
            // Prevent removing the paragraph if there are no songs/readings
            if (readings.Count() == 0) readings = readings.Append(new ServiceElement());

            // Loop over all Drawing.Run elements in the document
            foreach (SlidePart slidePart in Slides)
            {
                foreach (Run run in slidePart.Slide.Descendants<Run>())
                {
                    // Replace the run with the contents of the relevant list
                    KeywordSettings tags = Settings.Instance.Keywords;
                    if (run.InnerText.Contains(tags.Songs)) ReplaceMultilineKeywords(run, songs);
                    else if (run.InnerText.Contains(tags.Readings)) ReplaceMultilineKeywords(run, readings);
                }
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
            foreach (Presentation.Picture pic in slidePart.Slide.Descendants<Presentation.Picture>())
            {
                // Get the description and rId from the object
                string description = pic.NonVisualPictureProperties.NonVisualDrawingProperties.Description;
                string rId = pic.BlipFill.Blip.Embed.Value;
                
                if (description == Settings.Instance.ImageDescription)
                {
                    // Get the ImagePart by id, and replace the image
                    ImagePart imagePart = (ImagePart)slidePart.GetPartById(rId);
                    FileStream imageStream;
                    if (!Program.TryGetFileStream(imagePath, out imageStream)) return;
                    imagePart.FeedData(imageStream);
                    imageStream.Close();
                }
            }
        }

        private void RemoveImage(SlidePart slidePart)
        {
            // Loop over all picture objects
            foreach (Presentation.Picture pic in slidePart.Slide.Descendants<Presentation.Picture>())
            {
                // Remove the image if the description matches the ImageDescription setting
                string description = pic.NonVisualPictureProperties.NonVisualDrawingProperties.Description;
                if (description == Settings.Instance.ImageDescription) pic.Remove();
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
            Presentation.SlideIdList idList = PresPart.Presentation.SlideIdList;
            uint maxId = (from slideId in idList select ((Presentation.SlideId)slideId).Id).Max();

            // Get the first SlidePart from the SlideIdList
            Presentation.SlideId sourceSlideId = (Presentation.SlideId)idList.FirstChild;
            SlidePart sourceSlidePart = (SlidePart)PresPart.GetPartById(sourceSlideId.RelationshipId);

            // Copy the slide and SlideLayoutPart to a new slidepart, set it to visible
            SlidePart targetSlidePart = PresPart.AddNewPart<SlidePart>();
            sourceSlidePart.Slide.Save(targetSlidePart);
            targetSlidePart.AddPart(sourceSlidePart.SlideLayoutPart);
            targetSlidePart.Slide.Show = true;

            // Append a new id for the slidepart to the SlideIdList
            Presentation.SlideId targetSlideId = idList.AppendChild(new Presentation.SlideId());
            targetSlideId.Id = maxId + 1;
            targetSlideId.RelationshipId = PresPart.GetIdOfPart(targetSlidePart);
            
            // Copy all ImageParts from the source slide to the new one
            IEnumerable<ImagePart> imageParts = sourceSlidePart.ImageParts;
            foreach (ImagePart img in imageParts)
            {
                string rId = sourceSlidePart.GetIdOfPart(img);
                ImagePart newImagePart = targetSlidePart.AddImagePart(img.ContentType, rId);
                newImagePart.FeedData(img.GetStream(FileMode.Open));
            }

            PresPart.Presentation.Save();
            return targetSlidePart;
        }

        /// <summary>
        /// Duplicates the first slide, then calls
        /// <see cref="ReplaceKeywords(Dictionary{string, string}, Slide)"/>
        /// </summary>
        /// <param name="keywords">A dictionary containing replaceable strings
        /// and what they should be replaced by</param>
        public void DuplicateAndReplace(Dictionary<string, string> keywords, bool ShowQR)
        {
            SlidePart slidePart = DuplicateFirstSlide();
            ReplaceKeywords(keywords, slidePart.Slide);
            if (!ShowQR) RemoveImage(slidePart);
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
