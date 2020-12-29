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
        public static void Main()
        {
            Settings.Load();
            Service service = Service.GetService(Settings.ServicesXml, Settings.OrganistXml, "2020-12-20 9:30");
            // TODO: set UI content to service properties

            Console.WriteLine(Settings.OutputFolderPath + "test.pptx");
            Console.WriteLine(Settings.OutputFolderPath);
            PowerPoint pptx = new PowerPoint(Settings.TemplatePathBefore, "test.pptx"); // + "test.pptx");
            Console.WriteLine(Settings.ImagePath);
            pptx.ReplaceImage(Settings.ImagePath);

            // start UI
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new Window());

            Settings.Save();
            /*// edit de powerpoints met ReplacePlaceholders() en sla het resultaat op
            PresentationDocument pptx = PresentationDocument.Open(Settings.PPTXTemplatePre, true); // open pptx file. TODO: exception handling
            ReplacePlaceholders(pptx);
            pptx.SaveAs("../../edited.pptx");
            pptx.Close();*/
        }
    }
}
