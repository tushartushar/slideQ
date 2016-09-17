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
           
            foreach(SlideDataModel slide in dataModel.SlideDataModelList)
            {
              List<ShapeSize> SmelldShape=  slide.ShapeSize.Where(x => x.Height < Constants.BYOB_THRESHOLD_HEIGHT || x.Width < Constants.BYOB_THRESHOLD_WIDTH).ToList();
              List<CharAttribute> SmallTextSmell = slide.TextFontSize.Where(x => x.Size < Constants.BYOB_THRESHOLD_TEXT_SIZE).ToList();
              if (SmelldShape.Count > 0 || SmallTextSmell.Count >0)
                {
                    PresentationSmell smell = new PresentationSmell();
                    smell.SmellName = Constants.BYOB;
                    string Cause = "The tool detected the smell since the slide contains";
                    if (SmelldShape.Count > 0)
                    {
                        Cause=Cause+"(" + SmelldShape.Count + ")" +"small shapes(" + string.Join(",", SmelldShape.Select(x=>x.Name).ToArray()) + ")";
                    }
                    if (SmelldShape.Count > 0 && SmallTextSmell.Count > 0)
                    {
                        Cause = Cause + " and ";
                    }
                    if (SmelldShape.Count > 0)
                    {
                        Cause = Cause + "(" + SmallTextSmell.Count + ")" + "small character";
                    }
                    smell.Cause = Cause;
                    smell.SlideNo = slide.SlideNo;
                    smellList.Add(smell);
                }
            }
            return smellList;
        }
    }
}
