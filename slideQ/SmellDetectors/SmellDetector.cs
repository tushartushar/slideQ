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
            detectByobSmell(dataModel);
            ColormaniaSmell(dataModel);
            ItsLasVegasSmell(dataModel);
            SecretaryofHiPUSmell(dataModel);
            ChaoticStylistSmell(dataModel);
            StungbySpellBeeSmell(dataModel);
            SubSub__BulletSmell(dataModel);
            return smellsList;
        }

        private void detectTextHellSmell(MasterDataModel dataModel)
        {
            TexthellSmellDetector detector = new TexthellSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }
        private void detectByobSmell(MasterDataModel dataModel)
        {
            BYOBSmellDetector detector = new BYOBSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }
        private void ColormaniaSmell(MasterDataModel dataModel)
        {
            ColormaniaSmellDetector detector = new ColormaniaSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }
        private void ItsLasVegasSmell(MasterDataModel dataModel)
        {
            ItsLasVegasSmellDetector detector = new ItsLasVegasSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }

        private void SecretaryofHiPUSmell(MasterDataModel dataModel)
        {
            SecretaryofHiPUSmellDetector detector = new SecretaryofHiPUSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }

        private void ChaoticStylistSmell(MasterDataModel dataModel)
        {
            ChaoticStylistSmellDetectors detector = new ChaoticStylistSmellDetectors(dataModel);
            smellsList.AddRange(detector.detect());
        }

        private void StungbySpellBeeSmell(MasterDataModel dataModel)
        {
            StungbySpellBeeSmellDetector detector = new StungbySpellBeeSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }

        private void SubSub__BulletSmell(MasterDataModel dataModel)
        {
            SubSub__BulletSmellDetector detector = new SubSub__BulletSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }
    }
}
