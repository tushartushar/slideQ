using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using PowerPointAddInTest_CountSlides.Model;

namespace PowerPointAddInTest_CountSlides.SmellDetectors
{
    class MasterFunction
    {
        public void MasterFunctionCall (Microsoft.Office.Interop.PowerPoint.Slides Slides)
        {
            MaterModelResult Result = new MaterModelResult();
            ConsolidateMasterModel Data=new ConsolidateMasterModel();
            CommonFunction Func = new CommonFunction();
            Func.IterateSlide(Data, Slides);
            foreach(MasterModel item in Data.AnalyzedData)
            {
                MasterResultModelPerSlide slidesmell = new MasterResultModelPerSlide();
                TexthellSmellDetector THD = new TexthellSmellDetector();
                THD.GetTextHellSmells(item,slidesmell);
                Result.FinalResult.Add(slidesmell);
            }
        }
    }
}
