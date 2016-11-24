using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using PPT = Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using NHunspell;
using System.Reflection;
using System.IO;

namespace slideQ.Model
{
    public class SlideDataModel
    {
        private const string LOG_FILE_PATH = "slideQ_log.txt";
        private PPT.Slide slide;

        //private static string LogFilePath = @"F:\Rohit\SlideQ\Log\"+DateTime.Now.ToString("ddMMyyyyTHHmmss");
        public SlideDataModel(PPT.Slide slide)
        {
            this.slide = slide;
        }

        public void build()
        {
            TotalSpellingMistake = 0;
            IndentLevel = 0;
            countText();
            NoOfAnimationsInTheSlide= NoOfEnimationInSlide();
            CheckSlideHeaderFooter();
            GetSlideLayout();
            SlideNo = slide.SlideNumber;
        }

        public static void Log(string str)
        {
            try
            {
                StreamWriter Tex = File.AppendText(LOG_FILE_PATH);
                Tex.WriteLine(DateTime.Now.ToString() + " " + str);
                Tex.Close();
            }
            catch(Exception)
            {
                
            }
        }
        private void countText()
        {
            int count = 0;
            foreach (PPT.Shape shape in slide.Shapes)
                count += countTextFromShape(shape);
            
            TotalTextCount = count;
        }

        private int countTextFromShape(PPT.Shape shape)
        {
            int count = 0;
            if (shape.Type == MsoShapeType.msoGroup)
                foreach (PPT.Shape myShape in shape.GroupItems)
                    count += countTextFromShape(myShape);
            else if (shape.Type == MsoShapeType.msoSmartArt || shape.Type == MsoShapeType.msoPlaceholder)
            {
                try
                {
                    SmartArtNodes nodes = shape.SmartArt.AllNodes;
                    foreach (SmartArtNode node in nodes)
                    {
                        try
                        {
                            if (node.TextFrame2.HasText == MsoTriState.msoTrue)
                            {
                                GetNodeCharAttribute(node.TextFrame2);
                                string text = String.Join("", node.TextFrame2.TextRange.Text.Split('\t', '\r'));
                                count += text.Count();
                                TotalSpellingMistake = TotalSpellingMistake + spellcheckCore(node.TextFrame2.TextRange.Text);
                            }
                        }
                        catch (Exception ex)
                        {
                            Log("Exception occurred. " + ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Log("Exception occurred. " + ex.Message);
                }
            }

            if (shape.HasTextFrame == MsoTriState.msoTrue)
                count += CheckForHavingText(count, shape);
            return count;
        }

        
        private int CheckForHavingText(int count, PPT.Shape InheritShape)
        {
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
        private bool ExtractSlideTitlefromShape(PPT.Shape shape)
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

        private bool IsHavingUnderLine(PPT.Shape shape)
        {
            bool IsHavingUnderLineCounter = false;
            PPT.TextRange Textrange = shape.TextFrame.TextRange;
            for (int index = 0; index < Textrange.Text.Count(); index++)
            {
                PPT.TextRange text = Textrange.Find(Textrange.Text[index].ToString(), index);
                if(text.Font.Underline==MsoTriState.msoTrue)
                {
                    return true;
                }
            }
            return IsHavingUnderLineCounter;
        }

        private int GetCharCount(int count, PPT.Shape shape)
        {
            PPT.TextRange Textrange = shape.TextFrame.TextRange;
            //We need to remove \t and \r for accurate text count
            string text = String.Join("", Textrange.Text.Split('\t', '\r'));

            count += text.Count();
            return count;
        }

        private int GetspellingmistakeCount( PPT.Shape shape)
        {
            int count = 0;
            try
            {
                PPT.TextRange Textrange = shape.TextFrame.TextRange;
                string line = Textrange.Text.Trim();
                count = spellcheckCore( line);
            }
            catch(Exception ex)
            {
                Log("Exception occurred. " + ex.Message);
            }
            return count;
        }

        private int spellcheckCore( string line)
        {
            int count=0;
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
                    catch(Exception ex)
                    {
                        Log("Exception occurred. " + ex.Message);
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

        private void GetTextAttribute(PPT.Shape shape)
        {
            PPT.TextRange Textrange = shape.TextFrame.TextRange;
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
                    Log("Exception occurred. " + ex.Message);
                }
            }
        }

        private void GetCharAttribute(PPT.TextRange Textrange)
        {
            for (int index = 0; index < Textrange.Text.Count(); index++)
            {
                try
                {
                    PPT.TextRange text = Textrange.Find(Textrange.Text[index].ToString(), index);
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
                    Log("Exception occurred. " + ex.Message);
                }
            }
        }
        public void GetSlideLayout()
        {
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
                Log("Exception occurred. " + ex.Message);
            }
        }

        private int NoOfEnimationInSlide()
        {
            int numOfAnimations = 0;
            try
            {
                foreach (PPT.Effect e in slide.TimeLine.MainSequence)
                {
                    if (e.Timing.TriggerType == PPT.MsoAnimTriggerType.msoAnimTriggerOnPageClick)
                    {
                        numOfAnimations++;
                    }
                }
            }
            catch (Exception ex)
            {
                Log("Exception occurred. " + ex.Message);
            }
            return numOfAnimations;
        }
        public void CheckSlideHeaderFooter()
        {
            try
            {
                PPT.HeaderFooter Header = slide.HeadersFooters.Header;
                HeaderText = Header.Text;
            }
            catch(Exception ex)
            { 
                Log("Exception occurred. " + ex.Message); 
            }
            try
            {
                PPT.HeaderFooter Footer = slide.HeadersFooters.Footer;
                FooterText = Footer.Text;
            }
            catch(Exception ex)
            {
                Log("Exception occurred. " + ex.Message);
            }
        }

        public void CheckIndentLevelForBullet(PPT.Shape shape)
        {
            List<int> indentlevelList = new List<int>();
            PPT.TextFrame2 Textframe2 = shape.TextFrame2;  
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
        public PPT.PpSlideLayout Layout { get; set; }

        public int TotalSpellingMistake { get; set; }
        public PPT.ShapeRange SlideShapeRange { get; set; }

        public PPT.ColorScheme ColorSchem { get; set; }
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
