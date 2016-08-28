using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddInTest_CountSlides.Model
{
   public  class TextShapes
    {
       public Microsoft.Office.Interop.PowerPoint.Shape Shapeobj;
       public  string Text { get; set; }

       public string Name { get; set; }
    }
}
