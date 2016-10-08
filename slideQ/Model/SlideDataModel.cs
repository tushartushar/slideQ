using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;

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
            countText();
            NoOfAniMationINTheSlide= NoOfEnimationInSlide();
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
                GetTextAttribute(InheritShape);
            }
            return count;
        }

        private int GetCharCount(int count, Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
            count += Textrange.Text.Trim().Count();
            return count;
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
                TextFontSize.Add(attr);


            }


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
 

        public int TotalTextCount { get; set; }

        public int SlideNo { get; set; }

        public List<CharAttribute> TextFontSize = new List<CharAttribute>();

        public int NoOfAniMationINTheSlide { get; set; }

    }

    public class CharAttribute
    {
        public char Ch { get; set; }
        public float Size { get; set; }

        public int Color { get; set; }
    }
}
