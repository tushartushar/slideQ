using slideQ.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace slideQ.SmellDetectors
{
    class BYOBSmellDetector : ISmellDetector
    {
        private MasterDataModel dataModel;

        public BYOBSmellDetector(MasterDataModel dataModel)
        {
            this.dataModel = dataModel;
        }

        public List<PresentationSmell> detect()
        {
            List<PresentationSmell> smellList = new List<PresentationSmell>();

            foreach (SlideDataModel slide in dataModel.SlideDataModelList)
            {
                List<CharAttribute> SmallTextSmell = slide.TextFontSize.Where(x => x.Size < Constants.BYOB_THRESHOLD_TEXT_SIZE).ToList();
                if (SmallTextSmell.Count > 0)
                {
                    PresentationSmell smell = new PresentationSmell();
                    smell.SmellName = Constants.BYOB;
                    string Cause = "The tool detected the smell since the slide contains ( " + SmallTextSmell.Count + " ) " + "small character";
                    smell.Cause = Cause;
                    smell.SlideNo = slide.SlideNo;
                    smellList.Add(smell);
                }
            }
            return smellList;
        }
    }
}
