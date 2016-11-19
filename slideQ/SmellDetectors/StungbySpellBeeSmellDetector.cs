using slideQ.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace slideQ.SmellDetectors
{
    class StungbySpellBeeSmellDetector : ISmellDetector
    {
        private MasterDataModel dataModel;

        public StungbySpellBeeSmellDetector(MasterDataModel dataModel)
        {
            this.dataModel = dataModel;
        }

        public List<PresentationSmell> detect()
        {
            List<PresentationSmell> smellList = new List<PresentationSmell>();

            foreach (SlideDataModel slide in dataModel.SlideDataModelList)
            {
                 if (slide.TotalSpellingMistake > Constants.STUNG_BY_SPELLBEE_THRESHOLD)
                {
                    PresentationSmell smell = new PresentationSmell();
                    smell.SmellName = Constants.STUNG_BY_SPELLBEE;
                    string Cause = "The tool detected the smell since the slide contains ( " + slide.TotalSpellingMistake + " ) " + "Spelling Mistakes";
                    smell.Cause = Cause;
                    smell.SlideNo = slide.SlideNo;
                    smellList.Add(smell);
                }
            }
            return smellList;
        }
    }
}
