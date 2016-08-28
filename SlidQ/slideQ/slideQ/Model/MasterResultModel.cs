using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddInTest_CountSlides.Model
{
    public class MasterResultModelPerSlide
    {
        public TexthellSmellModel TextHell = new TexthellSmellModel();
    }

    public class MaterModelResult
    {
        public List<MasterResultModelPerSlide> FinalResult = new List<MasterResultModelPerSlide>();
   
    }

}
