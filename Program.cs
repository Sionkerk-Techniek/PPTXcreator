using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Presentation;
using DocumentFormat.OpenXml.Drawing;
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
            /*Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Form1());*/

            string path = "../../template_voor-dienst.pptx";
            PresentationDocument pptx = PresentationDocument.Open(path, true); // open pptx file
            PresentationPart part = pptx.PresentationPart;
            OpenXmlElementList slideIDs = part.Presentation.SlideIdList.ChildElements;

            string relID = (slideIDs[index] as SlideId).RelationshipId;

            // Get the slide part from the relationship ID.
            SlidePart slide = (SlidePart)part.GetPartById(relID);


            /*
            foreach (SlidePart part in pptx.PresentationPart.SlideParts)
            {
                Console.WriteLine(part.SlideParts);
            }

            

            foreach (var part in pptx.PresentationPart.SlideParts.ToList()[0].SlideParts)
            {
                Console.WriteLine(part);
            }*/
            //Console.WriteLine(pptx.PresentationPart.SlideParts.ToList().ToString());
        }
    }
}
