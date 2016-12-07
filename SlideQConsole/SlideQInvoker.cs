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
                if (!CheckInputFilePath(args[0]) || !CheckOutputFilePathDirectory(args[1]))
                {
                    Console.WriteLine("Input Error : make sure you are providing correct input in console");
                    Console.WriteLine("Help : your first argument is your input Presentation file. It must be a full file path like ");
                    Console.WriteLine("Help : your second argument is your output file. It must be a full file path and in this make sure the parent directory Exists on your system, file will be created automatically like ");
                }
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

        private static bool CheckInputFilePath(string path)
        {
            bool flag = false;
            if(File.Exists(path))
            {
                flag = true;
            }
            return flag;
        }
        private static bool CheckOutputFilePathDirectory(string path)
        {
            bool flag = false;
            if (Directory.Exists(path))
            {
                flag = true;
            }
            return flag;
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
