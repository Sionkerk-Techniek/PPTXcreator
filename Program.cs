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
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Window());

            string path = "../../../PPTXcreatorfiles/template_voor-dienst.pptx";
            PresentationDocument pptx = PresentationDocument.Open(path, true); // open pptx file
            PresentationPart part = pptx.PresentationPart;

            
            // This dictionary will eventually be replaced by values from the form
            Dictionary<String, String> templateContents = new Dictionary<String, String>()
            {
                { "#TIJD_NU", "13:37" },
                { "#DS_NU", "ds. Voornaam Achternaam" },
                { "#DS_NU_PLAATS", "Null Island" },
                { "#THEMA", "testthema" }
            };
            

            // Loop over slides
            foreach (SlidePart slide in part.SlideParts)
            {
                Console.Write("Slide: "); // for debugging
                Console.WriteLine(slide.Uri.ToString()); // for debugging
                
                // Loop over text in slides
                foreach (Drawing.Text text in slide.Slide.Descendants<Drawing.Text>())
                {
                    StringBuilder sb = new StringBuilder(text.Text);
                    Console.WriteLine(text.Text); // for debugging

                    // loop over replacable keywords
                    foreach (KeyValuePair<String, String> kvp in templateContents)
                    {
                        sb.Replace(kvp.Key, kvp.Value);
                    }
                    text.Text = sb.ToString(); // hopefully this saves to the file

                    Console.WriteLine(text.Text); // for debugging
                }
            }

            pptx.Close();
        }
    }
}
