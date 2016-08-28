using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using slideQ.Properties;
using PowerPointAddInTest_CountSlides.Model;
using PowerPointAddInTest_CountSlides;
using slideQ;

namespace PowerPointAddInTest_CountSlides.SmellDetectors
{
    class TexthellSmellDetector
    {
        public void GetTextHellSmells(MasterModel Data, MasterResultModelPerSlide TextHell)
        {
            List<string> Text = Data.TextContainingShapes.Select(x => x.Text).ToList();
            int Words = 0;
            int Chars = 0;
            IterateTextList(Text, ref Words, ref Chars);

            TexthellSmellModel obj = FillTextHellSmellObject(Data, Words, Chars);

            TextHell.TextHell = obj;
        }

        private static TexthellSmellModel FillTextHellSmellObject(MasterModel Data, int Words, int Chars)
        {
            TexthellSmellModel obj = new TexthellSmellModel();
            obj.CharCount = Chars;
            obj.SlideNmae = Data.MetaData.SlideNmae;
            obj.WordCount = Words;
            obj.SlideNumber = Data.MetaData.SlideNumber;
            if (obj.CharCount > Constants.TexthellThresholdValue)
            {
                obj.IsTexthellSmellpresent = true;
            }
            else
            {
                obj.IsTexthellSmellpresent = false;
            }
            return obj;
        }

        private static void IterateTextList(List<string> Text, ref int Words, ref int Chars)
        {
            foreach (string str in Text)
            {
                Chars = Chars + str.Length;

                try
                {
                    string[] WordsCountArray = str.Split(' ');
                    Words = Words + WordsCountArray.Where(x=>x != null && x != string.Empty).Count();
                }
                catch
                {

                }
            }
        }

        
    }


}
