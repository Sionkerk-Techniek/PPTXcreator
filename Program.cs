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
            { "#COLLECTE_2", "doel 2" },
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
            Settings.Load();

            // start UI
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Window());

            // edit de powerpoints met ReplacePlaceholders() en sla het resultaat op
            PresentationDocument pptx = PresentationDocument.Open(Settings.PPTXTemplatePre, true); // open pptx file. TODO: exception handling
            ReplacePlaceholders(pptx);
            pptx.SaveAs("../../edited.pptx");
            pptx.Close();
        }
    }
}
