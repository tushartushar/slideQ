using slideQ.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace slideQ.SmellDetectors
{
    class ColormaniaSmellDetector :ISmellDetector
    {
         private MasterDataModel dataModel;

         public ColormaniaSmellDetector(MasterDataModel dataModel)
        {
            this.dataModel = dataModel;
        }

         public List<PresentationSmell> detect()
         {
             List<PresentationSmell> smellList = new List<PresentationSmell>();

             foreach (SlideDataModel slide in dataModel.SlideDataModelList)
             {
                 int ColorCount = slide.TextFontSize.GroupBy(x=>x.Color).Select(x=>x.FirstOrDefault()).Count();
                 if (ColorCount > Constants.COLOR_MANIA_THRESHOLD)
                 {
                     PresentationSmell smell = new PresentationSmell();
                     smell.SmellName = Constants.COLORMANIA;
                     string Cause = "The tool detected the smell since the slide contains ( " + ColorCount + " ) " + "Multiple Color";
                     smell.Cause = Cause;
                     smell.SlideNo = slide.SlideNo;
                     smellList.Add(smell);
                 }
             }
             return smellList;
         }

    }
}
