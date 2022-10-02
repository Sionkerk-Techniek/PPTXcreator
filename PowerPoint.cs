using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Drawing;
using System.Drawing.Imaging;
using DocumentFormat.OpenXml;
using Presentation = DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace PPTXcreator
{
    class PowerPoint
    {
        private PresentationDocument Document { get; }
        private PresentationPart PresPart => Document.PresentationPart;
        private IEnumerable<SlidePart> Slides => PresPart.SlideParts;
        
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
        private void ReplaceMultilineKeywords(Run run, IEnumerable<ServiceElement> elements)
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
        /// Removes the runs containing <see cref="KeywordSettings.ThemeHeaderIdentifier"/>
        /// </summary>
        /// <param name="slide"></param>
        public void RemoveThemeRuns()
        {
            // Loop over slides
            foreach (SlidePart slide in Slides)
            {
                // Loop over text in the slide
                foreach (Text text in slide.Slide.Descendants<Text>())
                {
                    if (text.Text.ToLower().Contains(Settings.Instance.Keywords.ThemeHeaderIdentifier)
                        || text.Text.Contains(Settings.Instance.Keywords.Theme))
                    {
                        text.Parent.Remove();
                    }
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

                    using (Stream imageStream = GetEditedImage(imagePath))
                        if (imageStream != null) imagePart.FeedData(imageStream);
                }
            }
        }

        /// <summary>
        /// Replaces the image with description <see cref="Settings.ImageDescription"/>
        /// with the image at <paramref name="imagePath"/>.
        /// </summary>
        public void ReplaceImage(string imagePath)
        {
            if (!File.Exists(imagePath) && imagePath != "standaard QR")
                return;

            // Loop over slideparts (not really necessary if the slide number is known)
            foreach (SlidePart slidePart in Slides)
            {
                ReplaceImage(imagePath, slidePart);
            }
        }

        /// <summary>
        /// Removes the image with description <see cref="Settings.ImageDescription"/>
        /// </summary>
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
        /// Loads the file located at <paramref name="path"/> and applies crop and threshold
        /// </summary>
        public static Stream GetEditedImage(string path)
        {
            FileStream originalImageStream = null;
            Bitmap originalImage;
            ImageSettings settings = Settings.Instance.ImageParameters;

            // Try to get the filestream, return it without editing if EnableEditQR is false
            if (path != "standaard QR")
            {
                if (!Program.TryGetFileStream(path, out originalImageStream))
                    return null;
                if (!Settings.Instance.EnableEditQR)
                    return originalImageStream;

                originalImage = new Bitmap(originalImageStream);
            }
            else
            {
                originalImage = Properties.Resources.qr_code;
                settings = new ImageSettings();
            }

            // Check if the cropping region is valid, don't edit if it isn't
            if (!settings.IsValid()
                || settings.OffsetX + settings.Width > originalImage.Width
                || settings.OffsetY + settings.Height > originalImage.Height)
            {
                Dialogs.GenericWarning("De afbeelding kan niet worden bewerkt omdat het gebied " +
                    $"waar naar bijgesneden moet worden ({settings}) niet bestaat in de afbeelding " +
                    $"met resolutie {originalImage.Width}x{originalImage.Height}.");
                Program.MainWindow.CheckBoxEnableEditQRSetValue(false);

                if (originalImageStream == null)
                {
                    MemoryStream imageStream = new MemoryStream();
                    Properties.Resources.qr_code.Save(imageStream, ImageFormat.Png);
                    imageStream.Position = 0;
                    return imageStream;
                }
                else
                {
                    return originalImageStream;
                }
            }

            // Crop the image
            System.Drawing.Rectangle crop = new System.Drawing.Rectangle(
                settings.OffsetX, settings.OffsetY, settings.Width, settings.Height);
            Bitmap editedImage = originalImage.Clone(crop, PixelFormat.DontCare);

            if (!originalImage.PixelFormat.HasFlag(PixelFormat.Indexed))
            {
                // Manually iterate over the pixels to apply the threshold
                for (int x = 0; x < settings.Width; x++)
                {
                    for (int y = 0; y < settings.Height; y++)
                    {
                        // GetPixel and SetPixel are slow, but it shouldn't be a problem for a single image
                        Color pixel = editedImage.GetPixel(x, y);
                        if (pixel.GetBrightness() >= settings.Threshold)
                            editedImage.SetPixel(x, y, Color.White);
                        else
                            editedImage.SetPixel(x, y, Color.Black);
                    }
                }
            }

            // Convert to 1-bit indexed png and save to stream
            MemoryStream memoryStream = new MemoryStream();
            crop.X = crop.Y = 0; // include full image, .Clone always needs a rectangle for some reason
            editedImage.Clone(crop, PixelFormat.Format1bppIndexed).Save(memoryStream, ImageFormat.Png);
            memoryStream.Position = 0;
            originalImageStream?.Dispose();
            return memoryStream;
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

        public void RemoveSlide(int index)
        {
            if (Slides.Count() == 1) return; // Prevent illegal slideless presentations

            // Get the SlideIdList
            Presentation.SlideIdList idList = PresPart.Presentation.SlideIdList;

            if (index >= idList.ChildElements.Count) return; // slide does not exist

            // Get the SlideId from the SlideIdList and delete it from the list
            Presentation.SlideId sourceSlideId = (Presentation.SlideId)idList.ChildElements[index];
            string sourceSlideRId = sourceSlideId.RelationshipId;
            idList.RemoveChild(sourceSlideId);

            // Delete the actual slide
            PresPart.Presentation.Save();
            SlidePart sourceSlidePart = (SlidePart)PresPart.GetPartById(sourceSlideId.RelationshipId);
            PresPart.DeletePart(sourceSlidePart);
        }
    }
}
