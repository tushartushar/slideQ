using slideQ.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace slideQ.SmellDetectors
{
    class SubSubBulletSmellDetector : ISmellDetector
    {
        private MasterDataModel dataModel;

        public SubSubBulletSmellDetector(MasterDataModel dataModel)
        {
            this.dataModel = dataModel;
        }

        public List<PresentationSmell> detect()
        {
            List<PresentationSmell> smellList = new List<PresentationSmell>();

            foreach (SlideDataModel slide in dataModel.SlideDataModelList)
            {
                if (slide.MaxIndentLevel > Constants.SUBSUB_BULLET_THRESHOLD)
                {
                    PresentationSmell smell = new PresentationSmell();
                    smell.SmellName = Constants.SUBSUB_BULLET;
                    string Cause = "The tool detected the smell since the slide contains " + slide.MaxIndentLevel + " indentation levels.";
                    smell.Cause = Cause;
                    smell.SlideNo = slide.SlideNo;
                    smellList.Add(smell);
                }
            }
            return smellList;
        }
    }
}
