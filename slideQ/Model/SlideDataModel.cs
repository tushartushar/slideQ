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

        private static string LogFilePath = @"F:\Rohit\SlideQ\Log\"+DateTime.Now.ToString("ddMMyyyyTHHmmss");
        public SlideDataModel(Slide slide)
        {
            this.slide = slide;
        }


        public void build()
        {
            Log("analyze slide no " + slide.SlideNumber);
            TotalSpellingMistake = 0;
            IndentLevel = 0;
            countText();
            NoOfAnimationsInTheSlide= NoOfEnimationInSlide();
            CheckSlideHeaderFooter();
            GetSlideLayout();
            SlideNo = slide.SlideNumber;
        }

        void CreatetestlogFile()
        {
            if (!File.Exists(LogFilePath))
            {
                StreamWriter writer = new StreamWriter(LogFilePath);
                writer.Close();
            }
        }
        public static void Log(string str)
        {
            try
            {

                StreamWriter Tex = File.AppendText(SlideDataModel.LogFilePath);
                Tex.WriteLine(DateTime.Now.ToString() + " " + str);
                Tex.Close();

            }
            catch
            {

            }
        }
        private void countText()
        {
            
            int count = 0;
            int i = 0;
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
            {
                int gt = 0;
                try
                {
                  gt=  shape.GroupItems.Count ;
                  Log("analyze slide no " + slide.SlideNumber + " found group i=" + i + " total group items "+ gt);
                }
                catch 
                {
                    Log("analyze slide no " + slide.SlideNumber + " exception in count text on group count i = "+i );
                }
                if (shape.Type == MsoShapeType.msoGroup && gt!=0  )
                {
                    Log("analyze slide no " + slide.SlideNumber+ " found group i="+i);
                        count = CheckEveryShapeINGroup(count, shape);
                    
                }

                else if (shape.Type == MsoShapeType.msoSmartArt || shape.Type == MsoShapeType.msoPlaceholder || gt > 0)
                {

                    try
                    {
                        Log("analyze slide no " + slide.SlideNumber + " found msoSmartArt i=" + i);
                        SmartArtNodes nodes = shape.SmartArt.AllNodes;
                        foreach (SmartArtNode node in nodes)
                        {
                            try
                            {
                                if (node.TextFrame2.HasText == MsoTriState.msoTrue)
                                {
                                    GetNodeCharAttribute(node.TextFrame2);
                                    count = count + node.TextFrame2.TextRange.Text.Count();
                                    TotalSpellingMistake = TotalSpellingMistake + spellcheckCore(node.TextFrame2.TextRange.Text);
                                }
                            }
                            catch(Exception ex)
                            {
                                Log("analyze slide no " + slide.SlideNumber + " found msoSmartArt => HasText cond. i=" + i + " Exception" + ex.Message);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("analyze slide no " + slide.SlideNumber + " found msoSmartArt => SmartArt.AllNodes cond. i=" + i + " Exception" + ex.Message);
                 
                    }
                   // count = CheckEveryShapeINGroup(count, shape);

                }
           
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    count = CheckForHavingText(count, shape);
                    Log("analyze slide no " + slide.SlideNumber + " found group i =" + i + " count "+ count);
                }


                i++;
            }
            TotalTextCount = count;
        }

        private int CheckEveryShapeINGroup(int count, Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            int i = 0;
            foreach (Microsoft.Office.Interop.PowerPoint.Shape InheritShape in shape.GroupItems)
            {
                int gt = 0;
                try
                {
                    gt = InheritShape.GroupItems.Count;
                    Log("\t analyze slide no " + slide.SlideNumber + " found group inner shape  i=" + i + " total group items " + gt);
                }
                catch
                {
                    Log("\t analyze slide no " + slide.SlideNumber + " exception in inner shape count text on group count i = " + i);
                }

                if (InheritShape.Type == MsoShapeType.msoGroup && gt != 0)
                {
                    if (InheritShape.GroupItems.Count > 0)
                    {
                        Log("\t analyze slide no " + slide.SlideNumber + " found group i=" + i);
            
                        count = CheckEveryShapeINGroup(count, InheritShape);
                    }
                }

                else if (InheritShape.Type == MsoShapeType.msoSmartArt || InheritShape.Type == MsoShapeType.msoPlaceholder || gt > 0)
                {

                    try
                    {
                        Log("analyze slide no " + slide.SlideNumber + " found msoSmartArt i=" + i);
                        SmartArtNodes nodes = InheritShape.SmartArt.AllNodes;
                        foreach (SmartArtNode node in nodes)
                        {
                            try
                            {
                                if (node.TextFrame2.HasText == MsoTriState.msoTrue)
                                {
                                    GetNodeCharAttribute(node.TextFrame2);
                                    count = count + node.TextFrame2.TextRange.Text.Count();
                                    TotalSpellingMistake = TotalSpellingMistake + spellcheckCore(node.TextFrame2.TextRange.Text);
                                }
                            }
                            catch (Exception ex)
                            {
                                Log("analyze slide no " + slide.SlideNumber + " found msoSmartArt => HasText cond. i=" + i + " Exception" + ex.Message);
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Log("analyze slide no " + slide.SlideNumber + " found msoSmartArt => SmartArt.AllNodes cond. i=" + i + " Exception" + ex.Message);

                    }
                    // count = CheckEveryShapeINGroup(count, shape);

                }



                if (InheritShape.HasTextFrame == MsoTriState.msoTrue)
                {
                    count = CheckForHavingText(count, InheritShape);
                    Log("analyze slide no " + slide.SlideNumber + " found group i =" + i + " count " + count);
                }
                i++;
            }
            return count;
        }

        private int CheckForHavingText(int count, Microsoft.Office.Interop.PowerPoint.Shape InheritShape)
        {

            if (InheritShape.TextFrame.HasText == MsoTriState.msoTrue)
            {

                count = GetCharCount(count, InheritShape);
                Log("\tanalyze slide no " + slide.SlideNumber + " CheckForHavingText get count " + count);
                TotalSpellingMistake = TotalSpellingMistake + GetspellingmistakeCount(InheritShape);
                Log("\tanalyze slide no " + slide.SlideNumber + " CheckForHavingText get TotalSpellingMistake  " + TotalSpellingMistake);
                GetTextAttribute(InheritShape);
                Log("\tanalyze slide no " + slide.SlideNumber + " CheckForHavingText get TextAttribute  " );
             
                CheckIndentLevelForBullet(InheritShape);
                Log("\tanalyze slide no " + slide.SlideNumber + " CheckForHavingText get CheckIndentLevelForBullet  ");
                Log("\tanalyze slide no " + slide.SlideNumber + " CheckForHavingText get check if condition  for title");
             
                if (ExtractSlideTitlefromShape(InheritShape))
                {
                    Log("\tanalyze slide no " + slide.SlideNumber + " CheckForHavingText title found  ");
             
                    TitleHavingUnderLine = IsHavingUnderLine(InheritShape);
                    Log("\tanalyze slide no " + slide.SlideNumber + " CheckForHavingText get check if condition  for title under line ");
             
                }
            }
            return count;
        }
        private bool ExtractSlideTitlefromShape(Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            try
            {
                bool isTitleShape = shape.Name.ToLower().Contains("title");
                return isTitleShape;
            }
            catch
            {
                return false;
            }
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
            Log("\tanalyze slide no " + slide.SlideNumber + " GetCharCount get char count   ");
             
            Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
            count += Textrange.Text.Trim().Count();
            Log("\tanalyze slide no " + slide.SlideNumber + " GetCharCount  char count  = "+ count );
           
            return count;
        }

        private int GetspellingmistakeCount( Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            Log("\tanalyze slide no " + slide.SlideNumber + " GetspellingmistakeCount  ");
          
            int count = 0;
            try
            {
                Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
                string line = Textrange.Text.Trim();
                count = spellcheckCore( line);
            }
            catch
            {
                Log("\tanalyze slide no " + slide.SlideNumber + " GetspellingmistakeCount exception occur ");
          
            
            }
            return count;
        }

        private int spellcheckCore( string line)
        {int count=0;
            char[] spliter = { ' ', '\r', '\n', ')', '(', ',', ';', '.' };
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
                        Log("\tanalyze slide no " + slide.SlideNumber + " GetspellingmistakeCount exception occur in hunspell.Spell(word)condition");
                    }
                }
            }
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

            Log("\tanalyze slide no " + slide.SlideNumber + " GetTextAttribute  total char  " + Textrange.Text.Count());
            GetCharAttribute(Textrange);


        }


        private void GetNodeCharAttribute(Microsoft.Office.Core.TextFrame2 TextFrame)
        {
            Microsoft.Office.Core.TextRange2 Textrange = TextFrame.TextRange;

            for (int index = 0; index < Textrange.Text.Count(); index++)
            {
                try
                {
                    Microsoft.Office.Core.TextRange2 text = Textrange.Find(Textrange.Text[index].ToString(), index);
                    float sz = text.Font.Size;
                    CharAttribute attr = new CharAttribute();
                    attr.Size = sz;
                    attr.Ch = Textrange.Text[index];
                    attr.Color = text.Font.Fill.ForeColor.RGB;
                    attr.FontNameofChar = text.Font.Name;
                    TextFontSize.Add(attr);
                }
                catch (Exception ex)
                {
                    Log("\tanalyze slide no " + slide.SlideNumber + " GetCharAttribute exception occur : " + ex.Message);

                }
            }
        }


        private void GetCharAttribute(Microsoft.Office.Interop.PowerPoint.TextRange Textrange)
        {

            for (int index = 0; index < Textrange.Text.Count(); index++)
            {
                try
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
                catch (Exception ex)
                {
                    Log("\tanalyze slide no " + slide.SlideNumber + " GetCharAttribute exception occur : " + ex.Message);

                }
            }
        }
        public void GetSlideLayout()
        {
            Log("\tanalyze slide no " + slide.SlideNumber + " GetSlideLayout ");
     
            try
            {
                Theme = slide.ThemeColorScheme;
                ThemeObjCount = Theme.Count;
                Layout = slide.Layout;
                ColorSchem = slide.ColorScheme;
                ColorSchemCount = ColorSchem.Count;
                SlideShapeRange = slide.Background;
            }
            catch (Exception ex)
            {
                Log("\tanalyze slide no " + slide.SlideNumber + " GetSlideLayout exception occur : " + ex.Message);
          
            }
           
        }

        private int NoOfEnimationInSlide()
        {
            int numOfAnimations = 0;
            Log("\tanalyze slide no " + slide.SlideNumber + " NoOfEnimationInSlide");
          
            try
            {
                foreach (Effect e in slide.TimeLine.MainSequence)
                {
                    if (e.Timing.TriggerType == MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                    {
                        numOfAnimations++;
                    }
                }
            }
            catch (Exception ex)
            {
                Log("\tanalyze slide no " + slide.SlideNumber + " NoOfEnimationInSlide exception occur : " + ex.Message);
         
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

        public int NoOfAnimationsInTheSlide { get; set; }

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
