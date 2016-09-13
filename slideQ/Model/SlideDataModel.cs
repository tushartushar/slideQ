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

        private static int CheckEveryShapeINGroup(int count, Microsoft.Office.Interop.PowerPoint.Shape shape)
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

        private static int CheckForHavingText(int count, Microsoft.Office.Interop.PowerPoint.Shape InheritShape)
        {
            if (InheritShape.TextFrame.HasText == MsoTriState.msoTrue)
            {
                count = GetCharCount(count, InheritShape);

            }
            return count;
        }

        private static int GetCharCount(int count, Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
            count += Textrange.Text.Trim().Count();
            return count;
        }

        public int TotalTextCount { get; set; }

        public int SlideNo { get; set; }
    }
}
