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
        private const string LOG_FILE_PATH = @"C:\temp\slideQ_log.txt";
        private PPT.Slide slide;

        private List<PPT.TextRange> slideText = new List<PPT.TextRange>();

        private List<string> spellingMistakes = new List<string>();

        //private static string LogFilePath = @"F:\Rohit\SlideQ\Log\"+DateTime.Now.ToString("ddMMyyyyTHHmmss");
        public SlideDataModel(PPT.Slide slide)
        {
            this.slide = slide;
        }

        public void build()
        {
            MaxIndentLevel = 0;

            extractSlideInfo();
            countSlideText();

            extractSpellingmistakes();
            TotalSpellingMistake = spellingMistakes.Count;

            extractCharacterStyles();

            NoOfAnimationsInTheSlide= noOfEnimationInSlide();
            CheckSlideHeaderFooter();
            GetSlideLayout();
            SlideNo = slide.SlideNumber;
        }

        private void countSlideText()
        {
            TotalTextCount = 0;
            foreach (PPT.TextRange str in slideText)
            {
                TotalTextCount += str.Text.Length;
            }
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
            {}
        }
        private void extractSlideInfo()
        {
            foreach (PPT.Shape shape in slide.Shapes)
                extractInfoFromShape(shape);
        }

        private void extractInfoFromShape(PPT.Shape shape)
        {
            if (shape.Type == MsoShapeType.msoGroup)
                foreach (PPT.Shape myShape in shape.GroupItems)
                    extractInfoFromShape(myShape);
            else if (shape.Type == MsoShapeType.msoSmartArt)
            {
                try
                {
                    SmartArtNodes nodes = shape.SmartArt.AllNodes;
                
                foreach (SmartArtNode node in nodes)
                {
                    try
                    {
                        if (node.TextFrame2.HasText == MsoTriState.msoTrue)
                            slideText.Add((PPT.TextRange)node.TextFrame2.TextRange);
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
                extractInfo(shape);
        }

        
        private void extractInfo(PPT.Shape shape)
        {
            if (shape.TextFrame.HasText == MsoTriState.msoTrue)
            {
                slideText.Add(shape.TextFrame.TextRange);
                extractIndentLevels(shape);
                if (isSlideContainsTitle(shape))
                {
                    TitleHavingUnderLine = IsHavingUnderLine(shape);
                }
            }
        }
        private bool isSlideContainsTitle(PPT.Shape shape)
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
            PPT.TextRange Textrange = shape.TextFrame.TextRange;
            for (int index = 0; index < Textrange.Text.Count(); index++)
            {
                PPT.TextRange text = Textrange.Find(Textrange.Text[index].ToString(), index);
                if (text.Font.Underline == MsoTriState.msoTrue)
                {
                    return true;
                }
            }
            return false;
        }

        private void extractSpellingmistakes()
        {
            foreach (PPT.TextRange str in slideText)
                spellcheck(str.Text);
        }

        private void spellcheck(string line)
        {
            char[] spliter = { ' ', '\r', '\n', ')', '(', ',', ';', '.' };
            string[] words = line.Split(spliter);
            //nuget   http://www.nuget.org/packages/NHunspell/

            string afffilepath;
            string dicfilepath;
            WritefilesintempFolder(out afffilepath, out dicfilepath);

            try
            {
                using (Hunspell hunspell = new Hunspell(afffilepath, dicfilepath))
                {
                    foreach (string word in words)
                    {
                        if (!hunspell.Spell(word))
                        {
                            spellingMistakes.Add(word);
                        }

                    }
                }
            }
            catch (Exception ex)
            {
                Log("Exception occurred. " + ex.Message);
            }
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

        private void extractCharacterStyles()
        {
            foreach (PPT.TextRange str in slideText)
                getCharacterStyles(str);
        }

        private void getCharacterStyles(PPT.TextRange str)
        {
            for (int index = 0; index < str.Text.Count(); index++)
            {
                try
                {
                    PPT.TextRange text = str.Find(str.Text[index].ToString(), index);
                    float sz = text.Font.Size;
                    TextStyle style = new TextStyle();
                    style.Size = sz;
                    style.Character = str.Text[index];
                    style.Color = text.Font.Color.RGB;
                    style.FontName = text.Font.Name;
                    TextStlyeList.Add(style);
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

        private int noOfEnimationInSlide()
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

        public void extractIndentLevels(PPT.Shape shape)
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
            
            if(indentlevelList.Count>MaxIndentLevel)
            {
                MaxIndentLevel = indentlevelList.Count;
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
        public List<TextStyle> TextStlyeList = new List<TextStyle>();

        public int NoOfAnimationsInTheSlide { get; set; }

        public bool TitleHavingUnderLine { get; set; }

        public int MaxIndentLevel { get; set; }
        public string HeaderText { get; set; }

        public string FooterText { get; set; }
    }

    public class TextStyle
    {
        public char Character { get; set; }
        public float Size { get; set; }

        public int Color { get; set; }

        public string FontName { get; set; }
    }
}
