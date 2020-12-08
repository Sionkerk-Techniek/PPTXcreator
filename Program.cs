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
            //OpenXmlElementList slideIDs = part.Presentation.SlideIdList.ChildElements;
            
            // Loop over slides
            foreach (SlidePart slide in part.SlideParts)
            {
                Console.Write("Slide: "); // for debugging
                Console.WriteLine(slide.Uri.ToString()); // for debugging
                
                // Loop over text in slides
                foreach (Drawing.Text text in slide.Slide.Descendants<Drawing.Text>())
                {
                    Console.WriteLine(text.Text);
                }
            }

        }
    }
}
