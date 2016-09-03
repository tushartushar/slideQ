using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using slideQ.Model;
using Microsoft.Office.Interop.PowerPoint;
namespace slideQ.SmellDetectors
{
    public class SmellDetector
    {
        private List<PresentationSmell> smellsList;
        public List<PresentationSmell> detectPresentationSmells (Slides slides)
        {
            smellsList = new List<PresentationSmell>();
            MasterDataModel dataModel = new MasterDataModel(slides);
            dataModel.build();

            detectTextHellSmell(dataModel);
            return smellsList;
        }

        private void detectTextHellSmell(MasterDataModel dataModel)
        {
            TexthellSmellDetector detector = new TexthellSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }
    }
}
