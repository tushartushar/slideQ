using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using NHunspell;
using System.Reflection;
using System.IO;

namespace slideQ.Model
{
    public class SlideDataModel
    {
        private Slide slide;

        public SlideDataModel(Slide slide)
        {
            this.slide = slide;
        }

        public void build()
        {
            TotalSpellingMistake = 0;
            IndentLevel = 0;
            countText();
            NoOfAniMationINTheSlide= NoOfEnimationInSlide();
            CheckSlideHeaderFooter();
            GetSlideLayout();
            SlideNo = slide.SlideNumber;
        }

        private void countText()
        {
            
            int count = 0;
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
            {
                
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    if (shape.GroupItems.Count > 0)
                    {
                        count = CheckEveryShapeINGroup(count, shape);
                    }
                }
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    count = CheckForHavingText(count, shape);
                }



            }
            TotalTextCount = count;
        }

        private int CheckEveryShapeINGroup(int count, Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Shape InheritShape in shape.GroupItems)
            {
                if (InheritShape.Type == MsoShapeType.msoGroup)
                {
                    if (InheritShape.GroupItems.Count > 0)
                    {
                        count = CheckEveryShapeINGroup(count, InheritShape);
                    }
                }
                if (InheritShape.HasTextFrame == MsoTriState.msoTrue)
                {
                    count = CheckForHavingText(count, InheritShape);
                }

            }
            return count;
        }

        private int CheckForHavingText(int count, Microsoft.Office.Interop.PowerPoint.Shape InheritShape)
        {

            string name = InheritShape.Name;
            if(InheritShape.AnimationSettings.Animate==MsoTriState.msoTrue)
            {

            }
            if (InheritShape.TextFrame.HasText == MsoTriState.msoTrue)
            {

                count = GetCharCount(count, InheritShape);
                TotalSpellingMistake = TotalSpellingMistake + GetspellingmistakeCount(InheritShape);
                GetTextAttribute(InheritShape);
                CheckIndentLevelForBullet(InheritShape);
                if (ExtractSlideTitlefromShape(InheritShape))
                {
                    TitleHavingUnderLine = IsHavingUnderLine(InheritShape);
                }
            }
            return count;
        }
        private bool ExtractSlideTitlefromShape(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            bool isTitleShape = shape.Name.ToLower().Contains("title");
            return isTitleShape;
        }

        private bool IsHavingUnderLine(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            bool IsHavingUnderLineCounter = false;
            Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
            for (int index = 0; index < Textrange.Text.Count(); index++)
            {
                Microsoft.Office.Interop.PowerPoint.TextRange text = Textrange.Find(Textrange.Text[index].ToString(), index);
                if(text.Font.Underline==MsoTriState.msoTrue)
                {
                    return true;
                }
            }
            return IsHavingUnderLineCounter;
        }

        private int GetCharCount(int count, Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
            count += Textrange.Text.Trim().Count();
            return count;
        }

        private int GetspellingmistakeCount( Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            int count = 0;
            try
            {
                Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
                string line = Textrange.Text.Trim();
                char[] spliter =  { ' ', '\r','\n',')','(',',',';','.'};
                string[] words = line.Split(spliter);
                //add-in  http://www.nuget.org/packages/NHunspell/


                string afffilepath;
                string dicfilepath;
                WritefilesintempFolder(out afffilepath, out dicfilepath);

                using (Hunspell hunspell = new Hunspell(afffilepath, dicfilepath))
                {
                    foreach (string word in words)
                    {
                        try
                        {
                            if (!hunspell.Spell(word))
                            {
                                count++;
                            }
                        }
                        catch
                        {

                        }
                    }
                }
            }
            catch
            { }
            return count;
        }

        private static void WritefilesintempFolder(out string afffilepath, out string dicfilepath)
        {
            afffilepath = Path.Combine(Path.GetTempPath(), "en_us.aff");
            using (StreamWriter sw = new StreamWriter(afffilepath))
            {
                sw.WriteLine(slideQ.Properties.Resources.en_US_a);
            }

            dicfilepath = Path.Combine(Path.GetTempPath(), "en_us.dic");
            using (StreamWriter sw = new StreamWriter(dicfilepath))
            {
                sw.WriteLine(slideQ.Properties.Resources.en_US);
            }
        }


     
      
        private void GetTextAttribute(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;


            for (int index = 0; index < Textrange.Text.Count(); index++)
            {
                Microsoft.Office.Interop.PowerPoint.TextRange text = Textrange.Find(Textrange.Text[index].ToString(), index);
                float sz = text.Font.Size;
                CharAttribute attr = new CharAttribute();
                attr.Size = sz;
                attr.Ch = Textrange.Text[index];
                attr.Color = text.Font.Color.RGB;
                attr.FontNameofChar = text.Font.Name;
                TextFontSize.Add(attr);
            }


        }
        public void GetSlideLayout()
        {
            Theme = slide.ThemeColorScheme;
            ThemeObjCount = Theme.Count;
            Layout = slide.Layout;
            ColorSchem = slide.ColorScheme;
            ColorSchemCount = ColorSchem.Count;
            SlideShapeRange = slide.Background;
           
        }

        private int NoOfEnimationInSlide()
        {
            int numOfAnimations = 0;
            foreach (Effect e in slide.TimeLine.MainSequence)
            {
                if (e.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                {
                    numOfAnimations++;
                }
            }
            return numOfAnimations;
        }
        public void CheckSlideHeaderFooter()
        {
            try
            {
                Microsoft.Office.Interop.PowerPoint.HeaderFooter Header = slide.HeadersFooters.Header;
                HeaderText = Header.Text;
            }
            catch
            { }
            try
            {
                Microsoft.Office.Interop.PowerPoint.HeaderFooter Footer = slide.HeadersFooters.Footer;
                FooterText = Footer.Text;
            }
            catch
            {

            }
        }

        public void CheckIndentLevelForBullet(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {

            List<int> indentlevelList = new List<int>();
            Microsoft.Office.Interop.PowerPoint.TextFrame2 Textframe2 = shape.TextFrame2;  
            foreach (Microsoft.Office.Core.TextRange2 text in Textframe2.TextRange.Lines) 
            {
              int i=  text.ParagraphFormat.IndentLevel;
                if(indentlevelList.IndexOf(i)==-1)
                {
                    indentlevelList.Add(i);
                }
            }
            
            if(indentlevelList.Count>2)
            {
                IndentLevel++;
            }
        }

        public int TotalTextCount { get; set; }

        public int SlideNo { get; set; }
    
        public ThemeColorScheme Theme { get; set; }
        public int ThemeObjCount { get; set; }
        public PpSlideLayout Layout { get; set; }

        public int TotalSpellingMistake { get; set; }
        public Microsoft.Office.Interop.PowerPoint.ShapeRange SlideShapeRange { get; set; }

        public ColorScheme ColorSchem { get; set; }
        public int ColorSchemCount { get; set; }
        public List<CharAttribute> TextFontSize = new List<CharAttribute>();

        public int NoOfAniMationINTheSlide { get; set; }

        public bool TitleHavingUnderLine { get; set; }

        public int IndentLevel { get; set; }
        public string HeaderText { get; set; }

        public string FooterText { get; set; }
    }

    public class CharAttribute
    {
        public char Ch { get; set; }
        public float Size { get; set; }

        public int Color { get; set; }

        public string FontNameofChar { get; set; }
    }
}
