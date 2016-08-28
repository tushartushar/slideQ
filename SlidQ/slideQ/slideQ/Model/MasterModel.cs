using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace PowerPointAddInTest_CountSlides.Model
{
    public class MasterModel
    {
        public SlideMeaData MetaData { get; set; }
        public  List<TextShapes> TextContainingShapes { get; set; } 
    }

    public  class ConsolidateMasterModel
    {
        public List<MasterModel> AnalyzedData = new List<MasterModel>();
    }
}
