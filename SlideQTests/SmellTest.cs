using System;
using NUnit.Framework;
using slideQ.Model;
using slideQ.SmellDetectors;
using slideQ;
using Microsoft.Office.Core;
using Microsoft.Office.Interop;
using System.IO;
using Microsoft.Office.Interop.PowerPoint;

namespace SlideQTests
{
    [TestFixture]
    public class SmellTest
    {
        private _Presentation PPTObject;

        [SetUp]
        public void GetPPTObject()
        {
            string path = @"../../../SlideQTests/TestFile/SlideQTestCount.pptx";
            string absolute = Path.GetFullPath(path);
            Application ppApp = new Application();
            ppApp.Visible = MsoTriState.msoTrue;
            Presentations oPresSet = ppApp.Presentations;
            PPTObject = oPresSet.Open(absolute, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
        }
        [Test]
        public void TextHellTest()
        {                      

        }
    }
}
