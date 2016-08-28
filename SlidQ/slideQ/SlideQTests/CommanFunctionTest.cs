using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using PowerPointAddInTest_CountSlides.Model;
using PowerPointAddInTest_CountSlides.SmellDetectors;
using slideQ;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using System.IO;

namespace SlideQTests
{
    [TestClass]
    public class CommanFunctionTest
    {

        public Microsoft.Office.Interop.PowerPoint._Presentation GetPPTObject()
        {
            string path = @"../../../SlideQTests/TestFile/SlideQTestCount.pptx";
            string absolute = Path.GetFullPath(path);
            Microsoft.Office.Interop.PowerPoint.Application ppApp = new Microsoft.Office.Interop.PowerPoint.Application();
            ppApp.Visible = Microsoft.Office.Core.MsoTriState.msoTrue;
            Microsoft.Office.Interop.PowerPoint.Presentations oPresSet = ppApp.Presentations;
            Microsoft.Office.Interop.PowerPoint._Presentation oPres = oPresSet.Open(absolute, Microsoft.Office.Core.MsoTriState.msoFalse, Microsoft.Office.Core.MsoTriState.msoFalse,
            Microsoft.Office.Core.MsoTriState.msoTrue);
            return oPres;
        }
        [TestMethod]
        public void SlideCountTest()
        {                      
            CommonFunction fun = new CommonFunction();
            Microsoft.Office.Interop.PowerPoint._Presentation PPTObject= GetPPTObject();
            int count =  fun.GetSlideCount(PPTObject.Slides);
            Assert.AreEqual(2, count);
        }
    }
}
