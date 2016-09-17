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

                ShapeSize sizeobj = new Model.ShapeSize();
                sizeobj.Height = shape.Height;
                sizeobj.Width = shape.Width;
                sizeobj.Name = shape.Name;
                ShapeSize.Add(sizeobj);

            }
            TotalTextCount = count;
        }

        private  int CheckEveryShapeINGroup(int count, Microsoft.Office.Interop.PowerPoint.Shape shape)
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
                ShapeSize sizeobj = new Model.ShapeSize();
                sizeobj.Height = InheritShape.Height;
                sizeobj.Width = InheritShape.Width;
                sizeobj.Name = InheritShape.Name;
                ShapeSize.Add(sizeobj);
            }
            return count;
        }

        private  int CheckForHavingText(int count, Microsoft.Office.Interop.PowerPoint.Shape InheritShape)
        {
            if (InheritShape.TextFrame.HasText == MsoTriState.msoTrue)
            {
                count = GetCharCount(count, InheritShape);

            }
            return count;
        }

        private  int GetCharCount(int count, Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
            count += Textrange.Text.Trim().Count();

            for (int index = 0; index < Textrange.Text.Count(); index++)
            {
                Microsoft.Office.Interop.PowerPoint.TextRange text = Textrange.Find(Textrange.Text[index].ToString(), index);
                float sz = text.Font.Size;

                Microsoft.Office.Interop.PowerPoint.ColorFormat clr = text.Font.Color;
               int flag= TextFontSize.Where(x => x.Size == sz).Count();
               if (flag == 0)
                {
                    CharAttribute attr = new CharAttribute();
                    attr.Size = sz;
                    attr.Ch = Textrange.Text[index];
                    TextFontSize.Add(attr);

                }
            }

            return count;
        }

        public int TotalTextCount { get; set; }

        public int SlideNo { get; set; }

        public List<CharAttribute> TextFontSize = new List<CharAttribute>();
        public List<ShapeSize> ShapeSize = new List<ShapeSize>();
    }
    public class ShapeSize
    {
        public float Height { get; set; }
        public float Width { get; set; }
        public string Name { get; set; }
    }
    public class CharAttribute
    {
        public char Ch { get; set; }
        public float Size { get; set; }
    }
}
