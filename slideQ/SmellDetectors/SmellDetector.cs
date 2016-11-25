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
            detectColormaniaSmell(dataModel);
            detectItsLasVegasSmell(dataModel);
            detectSecretaryofHiPUSmell(dataModel);
            detectChaoticStylistSmell(dataModel);
            detectStungbySpellBeeSmell(dataModel);
            detectSubSubBulletSmell(dataModel);
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
        private void detectColormaniaSmell(MasterDataModel dataModel)
        {
            ColormaniaSmellDetector detector = new ColormaniaSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }
        private void detectItsLasVegasSmell(MasterDataModel dataModel)
        {
            ItsLasVegasSmellDetector detector = new ItsLasVegasSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }

        private void detectSecretaryofHiPUSmell(MasterDataModel dataModel)
        {
            SecretaryofHiPUSmellDetector detector = new SecretaryofHiPUSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }

        private void detectChaoticStylistSmell(MasterDataModel dataModel)
        {
            ChaoticStylistSmellDetectors detector = new ChaoticStylistSmellDetectors(dataModel);
            smellsList.AddRange(detector.detect());
        }

        private void detectStungbySpellBeeSmell(MasterDataModel dataModel)
        {
            StungbySpellBeeSmellDetector detector = new StungbySpellBeeSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }

        private void detectSubSubBulletSmell(MasterDataModel dataModel)
        {
            SubSubBulletSmellDetector detector = new SubSubBulletSmellDetector(dataModel);
            smellsList.AddRange(detector.detect());
        }
    }
}
