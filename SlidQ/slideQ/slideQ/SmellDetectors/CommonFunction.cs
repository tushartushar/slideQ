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
using Microsoft.Office.Core;
using Microsoft.Office.Interop;

namespace PowerPointAddInTest_CountSlides.SmellDetectors
{
  public  class CommonFunction
    {
        public int GetSlideCount(Microsoft.Office.Interop.PowerPoint.Slides Slides)
        {
            int NumberOfSlide = 0;
            try
            {

                NumberOfSlide = Slides.Count;
            }
            catch
            {
                NumberOfSlide = 0;
            }
            return NumberOfSlide;
        }

        public void IterateSlide(ConsolidateMasterModel Data, Microsoft.Office.Interop.PowerPoint.Slides Slides)
        {
           
            foreach (Microsoft.Office.Interop.PowerPoint.Slide slide in Slides)
            {
                MasterModel Master = new MasterModel();
                Master.TextContainingShapes = new List<TextShapes>();
                IterateShapes(slide, Master);
                SlideMeaData Smeta = FillSlideMetaData(slide, Master);
                Master.MetaData = Smeta;
                Data.AnalyzedData.Add(Master);
            }
        }

        private static void IterateShapes(Microsoft.Office.Interop.PowerPoint.Slide slide, MasterModel Master)
        {
            foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
            {
                CheckForTextShapes(Master, shape);

            }
        }

        private static void CheckForTextShapes(MasterModel Master, Microsoft.Office.Interop.PowerPoint.Shape shape)
        {
            if (shape.HasTextFrame == MsoTriState.msoTrue)
            {

                if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    TextShapes Shapeobj = new TextShapes();
                    Shapeobj.Name = shape.Name;
                    Microsoft.Office.Interop.PowerPoint.TextRange Textrange = shape.TextFrame.TextRange;
                    Shapeobj.Text = Textrange.Text;
                    Shapeobj.Shapeobj = shape;
                    Master.TextContainingShapes.Add(Shapeobj);
                }
            }
        }

        private static SlideMeaData FillSlideMetaData(Microsoft.Office.Interop.PowerPoint.Slide slide, MasterModel Master)
        {
            SlideMeaData Smeta = new SlideMeaData();
            Smeta.SlideNmae = slide.Name;
            Smeta.SlideNumber = slide.SlideNumber;
            return Smeta;
        }
    }
}
