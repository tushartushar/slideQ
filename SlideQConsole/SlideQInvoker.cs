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
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace SlideQConsole
{
    class SlideQInvoker
    {
        static void Main(string[] args)
        {
                Application ppApp = new Application();
                ppApp.Visible = MsoTriState.msoTrue;
                Presentations oPresSet = ppApp.Presentations;
                _Presentation PPTObject = oPresSet.Open(@args[0], MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
                SmellDetector detector = new SmellDetector();
                List<PresentationSmell> SlideDataModelList = detector.detectPresentationSmells(PPTObject.Slides);
                PPTObject.Close();
                try
                {
                    Excelgenerator(SlideDataModelList, args[1]);
                }
            catch(Exception ex)
                {
                    Console.WriteLine("message " + ex.Message );
                }
        }

        private static void Excelgenerator(List<PresentationSmell> SlideDataModelList,string path)
        {
            Excel.Application xlApp = new Microsoft.Office.Interop.Excel.Application();

            if (xlApp == null)
            {
                Console.WriteLine("Excel is not properly installed!!");
                return;
            }


            Excel.Workbook xlWorkBook;
            Excel.Worksheet xlWorkSheet;
            object misValue = System.Reflection.Missing.Value;

            xlWorkBook = xlApp.Workbooks.Add(misValue);
            xlWorkSheet = (Excel.Worksheet)xlWorkBook.Worksheets.get_Item(1);

            xlWorkSheet.Cells[1, 1] = "Smell Name";
            xlWorkSheet.Cells[1, 2] = "slide No";
            xlWorkSheet.Cells[1, 3] = "Cause";
            int k = 2;
            foreach(PresentationSmell smell in SlideDataModelList)
            {
                xlWorkSheet.Cells[k, 1] = smell.SmellName;
                xlWorkSheet.Cells[k, 2] = smell.SlideNo;
                xlWorkSheet.Cells[k, 3] = smell.Cause;
                k++;
            }


            xlWorkBook.SaveAs(@path, Excel.XlFileFormat.xlWorkbookNormal, misValue, misValue, misValue, misValue, Excel.XlSaveAsAccessMode.xlExclusive, misValue, misValue, misValue, misValue, misValue);
            xlWorkBook.Close(true, misValue, misValue);
            xlApp.Quit();

            Marshal.ReleaseComObject(xlWorkSheet);
            Marshal.ReleaseComObject(xlWorkBook);
            Marshal.ReleaseComObject(xlApp);

            Console.WriteLine("Excel file created , you can find the file " + path);
        }
    }
}
