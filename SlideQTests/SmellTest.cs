using System;
using NUnit.Framework;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;
using slideQ.Model;
using slideQ.SmellDetectors;

namespace SlideQTests
{
    [TestFixture]
    public class SmellTest
    {
        private _Presentation PPTObject;

        [SetUp]
        public void GetPPTObject()
        {
            string solution_dir = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.WorkDirectory));
            string path = @solution_dir+@"\TestFile\SlideQTestCount.pptx";
            string absolute = Path.GetFullPath(path);
            Application ppApp = new Application();
            ppApp.Visible = MsoTriState.msoTrue;
            Presentations oPresSet = ppApp.Presentations;
            PPTObject = oPresSet.Open(@path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
        }
        [Test]
        public void TextCountTest()
        {
            List<SlideDataModel> SlideDataModelList = new List<SlideDataModel>();
            foreach (Slide slide in PPTObject.Slides)
            {
                SlideDataModel slideModel = new SlideDataModel(slide);
                slideModel.build();
                SlideDataModelList.Add(slideModel);
            }
            
            Assert.AreEqual(5, SlideDataModelList[0].TotalTextCount);
            Assert.AreEqual(0, SlideDataModelList[1].TotalTextCount);
        }
        [Test]
        public void BYOBSmellTest()
        {
            SmellDetector detector = new SmellDetector();
            List<PresentationSmell> presentationSmells = detector.detectPresentationSmells(PPTObject.Slides);
            bool found = false;
            
            foreach(PresentationSmell smell in presentationSmells)
            {
               
                if (smell.SmellName.Equals(slideQ.Constants.BYOB) && smell.SlideNo == 4)
                    found = true;

            }
            Assert.AreEqual(true, found);
            
        }

        [Test]
        public void TextHellSmellTest()
        {
            SmellDetector detector = new SmellDetector();
            List<PresentationSmell> presentationSmells = detector.detectPresentationSmells(PPTObject.Slides);

            bool found = false;
            foreach (PresentationSmell smell in presentationSmells)
            {
                if (smell.SmellName.Equals(slideQ.Constants.TEXTHELL) && smell.SlideNo == 3)
                    found = true;
            }
            Assert.AreEqual(true, found);
        }

        [Test]
        public void ColormaniaSmellTest()
        {
            SmellDetector detector = new SmellDetector();
            List<PresentationSmell> presentationSmells = detector.detectPresentationSmells(PPTObject.Slides);

            bool found = false;
            foreach (PresentationSmell smell in presentationSmells)
            {
                if (smell.SmellName.Equals(slideQ.Constants.COLORMANIA) && smell.SlideNo == 4)
                    found = true;
            }
            Assert.AreEqual(true, found);
        }

        [Test]
        public void ItsLasVegasSmell()
        {
            SmellDetector detector = new SmellDetector();
            List<PresentationSmell> presentationSmells = detector.detectPresentationSmells(PPTObject.Slides);

            bool found = false;
            foreach (PresentationSmell smell in presentationSmells)
            {
                if (smell.SmellName.Equals(slideQ.Constants.ItsLasVegas) && smell.SlideNo == 2)
                    found = true;
            }
            Assert.AreEqual(true, found);
        }
    }
}
