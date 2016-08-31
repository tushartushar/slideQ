using System;
using NUnit.Framework;
using slideQ.Model;
using slideQ.SmellDetectors;
using slideQ;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;
using System.Collections.Generic;

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
        public void TextHellTest()
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
    }
}
