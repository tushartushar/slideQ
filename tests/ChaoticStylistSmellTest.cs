using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using NUnit.Framework;
using slideQ.Model;
using slideQ.SmellDetectors;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideQTests.TestFile
{
    [TestFixture]
    class ChaoticStylistSmellTest1
    {

        public List<PresentationSmell> SlideDataModelList { get; set; }

        _Presentation PPTObject { get; set; }

          [SetUp]
        public void GetPPTObject()
        {
            string solution_dir = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.WorkDirectory));
            string path = @solution_dir + @"\TestFile\ChaoticStylistWithSmell.pptx";
            string absolute = Path.GetFullPath(path);
            Application ppApp = new Application();
            ppApp.Visible = MsoTriState.msoTrue;
            Presentations oPresSet = ppApp.Presentations;
             PPTObject = oPresSet.Open(@path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
            SmellDetector detector = new SmellDetector();
            SlideDataModelList = detector.detectPresentationSmells(PPTObject.Slides);
        }

          [TearDown]
          public void tearDown()
          {
              PPTObject.Close();
          }

        [Test]
        public void ChaoticStylistWithSmell()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 1 && x.SmellName.Equals(slideQ.Constants.CHAOTIC_STYLIST)).Count() != 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }


        
    }
}
