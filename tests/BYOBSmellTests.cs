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

namespace SlideQTests
{
    class BYOBSmellTests
    {
        private List<PresentationSmell> SlideDataModelList = null;
        _Presentation PPTObject { get; set; }
        [SetUp]
        public void GetPPTObject()
        {
            string solution_dir = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.WorkDirectory));
            string path = @solution_dir + @"\TestFile\Byob.pptx";
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
        public void NormalTest()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 1 && x.SmellName.Equals(slideQ.Constants.BYOB)).Count() != 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }

        [Test]
        public void SmartArtTest()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 2 && x.SmellName.Equals(slideQ.Constants.BYOB)).Count() != 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }

        [Test]
        public void GroupedSmartArtTest()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 3 && x.SmellName.Equals(slideQ.Constants.BYOB)).Count() != 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }

        [Test]
        public void ChartTest()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 4 && x.SmellName.Equals(slideQ.Constants.BYOB)).Count() == 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }

        [Test]
        public void BulletDiffrentStyleTextTest()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 6 && x.SmellName.Equals(slideQ.Constants.BYOB)).Count() != 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }

        [Test]
        public void nested()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 7 && x.SmellName.Equals(slideQ.Constants.BYOB)).Count() != 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }

        [Test]
        public void EmptySlideTestTest()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 8 && x.SmellName.Equals(slideQ.Constants.BYOB)).Count() == 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }
    }
}
