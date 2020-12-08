using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using Drawing = DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;

namespace PPTXcreator
{
    static class Program
    {
        /// <summary>
        /// The path at which the .pptx file is located. Will be replaced by a path from the form in the future
        /// </summary>
        public static string Path = "../../../PPTXcreatorfiles/template_voor-dienst.pptx";


        /// <summary>
        /// Dictionary containing the values with which the keys will be replaced in the slides.
        /// This dictionary will eventually be filled with form values instead of this way.
        /// </summary>
        public static Dictionary<string, string> TemplateContents = new Dictionary<string, string>()
        {
            { "#TIJD_1", "13:37" },
            { "#TIJD_2", "16:15" },
            { "#DS_PLAATS_1", "Null Island" },
            { "#DS_NAAM_1", "ds. Voornaam Achternaam" },
            { "#DS_PLAATS_2", "Null Island" },
            { "#DS_NAAM_2", "ds. Voornaam Achternaam" },
            { "#ZINGEN_7X", "wat we hier zingen\r\nwerkt deze enter?\nen deze?\r\n4\r\n5\r\n6\r\n7" },
            { "#LEZING", "asdf" },
            { "#THEMA", "testthema" }
        };


        /// <summary>
        /// Replaces all instances of keys in TemplateContents with their respective values
        /// </summary>
        /// <param name="document">The PresentationDocument object in which to replace the placeholders</param>
        private static void ReplacePlaceholders(PresentationDocument document)
        {
            PresentationPart part = document.PresentationPart;

            // Loop over slides
            foreach (SlidePart slide in part.SlideParts)
            {
                // Loop over text in slides
                foreach (Drawing.Text text in slide.Slide.Descendants<Drawing.Text>())
                {
                    StringBuilder sb = new StringBuilder(text.Text);

                    // loop over replacable keywords
                    foreach (KeyValuePair<string, string> kvp in TemplateContents)
                    {
                        sb.Replace(kvp.Key, kvp.Value);
                    }
                    text.Text = sb.ToString();
                }
            }
        }


        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        public static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Window());

            
            PresentationDocument pptx = PresentationDocument.Open(Path, true); // open pptx file
            ReplacePlaceholders(pptx);
            pptx.Close();
        }
    }
}
