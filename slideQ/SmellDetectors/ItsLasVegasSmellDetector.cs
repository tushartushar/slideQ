using slideQ.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace slideQ.SmellDetectors
{
    class ItsLasVegasSmellDetector : ISmellDetector
    {
          private MasterDataModel dataModel;

          public ItsLasVegasSmellDetector(MasterDataModel dataModel)
        {
            this.dataModel = dataModel;
        }

        public List<PresentationSmell> detect()
        {
            List<PresentationSmell> smellList = new List<PresentationSmell>();

            foreach (SlideDataModel slide in dataModel.SlideDataModelList)
            {
                if (slide.NoOfAnimationsInTheSlide > Constants.ITSLASVEGAS_ANIMATION_THRESHOLD)
                {
                    PresentationSmell smell = new PresentationSmell();
                    smell.SmellName = Constants.ITS_LAS_VEGAS;
                    string Cause = "The tool detected the smell since the slide contains " + slide.NoOfAnimationsInTheSlide + " animations.";
                    smell.Cause = Cause;
                    smell.SlideNo = slide.SlideNo;
                    smellList.Add(smell);
                }
            }
            return smellList;
        }
    }
}
