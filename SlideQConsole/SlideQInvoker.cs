using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using slideQ;
using System.IO;
using slideQ.SmellDetectors;
using slideQ.Model;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

namespace SlideQConsole
{
    class SlideQInvoker
    {
        static void Main(string[] args)
        {

            List<string> filepaths = new List<string>();
            while (true)
            {
                Console.WriteLine("Enter PPT Path Path or for exit enter 0");
                string temp = Console.ReadLine();
                if(temp.Equals("0"))
                {
                    break;
                }
                else
                {
                    filepaths.Add(temp);
                }
            }

            Console.WriteLine("Enter Destination Directory Path");
            string dest = Console.ReadLine();
            foreach (string s in filepaths)
            {

                Application ppApp = new Application();
                ppApp.Visible = MsoTriState.msoTrue;
                Presentations oPresSet = ppApp.Presentations;
                _Presentation PPTObject = oPresSet.Open(@s, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
                SmellDetector detector = new SmellDetector();
                List<PresentationSmell> SlideDataModelList = detector.detectPresentationSmells(PPTObject.Slides);
                PPTObject.Close();
            }
        }
    }
}
