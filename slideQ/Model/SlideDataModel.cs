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
                if (shape.HasTextFrame == MsoTriState.msoTrue)
                {
                    if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                    {
                        Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
                        count += Textrange.Text.Count();
                    }
                }
            }
            TotalTextCount = count;
        }

        public int TotalTextCount { get; set; }

        public int SlideNo { get; set; }
    }
}
