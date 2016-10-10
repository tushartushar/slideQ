using slideQ.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace slideQ.SmellDetectors
{
    class SecretaryofHiPUSmellDetector :ISmellDetector
    {
          private MasterDataModel dataModel;

          public SecretaryofHiPUSmellDetector(MasterDataModel dataModel)
        {
            this.dataModel = dataModel;
        }

        public List<PresentationSmell> detect()
        {
            List<PresentationSmell> smellList = new List<PresentationSmell>();

            foreach (SlideDataModel slide in dataModel.SlideDataModelList)
            {
                if (slide.TitleHavingUnderLine==true )
                {
                    PresentationSmell smell = new PresentationSmell();
                    smell.SmellName = Constants.SecretaryofHiPU;
                    string Cause = "The tool detected the smell since the slide contains under the slide Title";
                    smell.Cause = Cause;
                    smell.SlideNo = slide.SlideNo;
                    smellList.Add(smell);
                }
            }
            return smellList;
        }
    }
}
