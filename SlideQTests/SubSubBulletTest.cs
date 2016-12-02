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
    class SubSubBulletTest
    {
        private List<PresentationSmell> SlideDataModelList = null;
        private _Presentation PPTObject;

        [SetUp]
        public void GetPPTObject()
        {
            string solution_dir = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.WorkDirectory));
            string path = @solution_dir + @"\TestFile\SubSubBullet.pptx";
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
        public void NegativeTest1()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 1 && x.SmellName.Equals(slideQ.Constants.SUBSUB_BULLET)).Count() == 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }

        [Test]
        public void NegativeTest2()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 5 && x.SmellName.Equals(slideQ.Constants.SUBSUB_BULLET)).Count() == 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }

        [Test]
        public void PositiveTest1()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 2 && x.SmellName.Equals(slideQ.Constants.SUBSUB_BULLET)).Count() != 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }

        [Test]
        public void PositiveTest2()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 3 && x.SmellName.Equals(slideQ.Constants.SUBSUB_BULLET)).Count() != 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }
        [Test]
        public void PositiveTest3()
        {
            bool flag = false;
            if (SlideDataModelList.Where(x => x.SlideNo == 4 && x.SmellName.Equals(slideQ.Constants.SUBSUB_BULLET)).Count() != 0)
                flag = true;
            Assert.AreEqual(true, flag);
        }


    }
}
