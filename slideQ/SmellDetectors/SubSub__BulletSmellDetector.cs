using slideQ.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace slideQ.SmellDetectors
{
    class SubSub__BulletSmellDetector : ISmellDetector
    {
        private MasterDataModel dataModel;

        public SubSub__BulletSmellDetector(MasterDataModel dataModel)
        {
            this.dataModel = dataModel;
        }

        public List<PresentationSmell> detect()
        {
            List<PresentationSmell> smellList = new List<PresentationSmell>();

            foreach (SlideDataModel slide in dataModel.SlideDataModelList)
            {
                if (slide.IndentLevel > Constants.SUBSUB_BULLET_THRESHOLD)
                {
                    PresentationSmell smell = new PresentationSmell();
                    smell.SmellName = Constants.SUBSUB_BULLET;
                    string Cause = "The tool detected the smell since the slide contains ( " + slide.IndentLevel + " ) " + "Shapes which contains bullet indent level more then two";
                    smell.Cause = Cause;
                    smell.SlideNo = slide.SlideNo;
                    smellList.Add(smell);
                }
            }
            return smellList;
        }
    }
}
