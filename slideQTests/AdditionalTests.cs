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
    class AdditionalTests
    {
        private _Presentation pptObject;
        [TearDown]
        public void tearDown()
        {
            pptObject.Close();
        }
        [SetUp]
        public void GetPPTObject()
        {
            string solution_dir = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.WorkDirectory));
            string path = @solution_dir + @"\TestFile\DECAF.pptx";
            string absolute = Path.GetFullPath(path);
            Application ppApp = new Application();
            ppApp.Visible = MsoTriState.msoTrue;
            Presentations oPresSet = ppApp.Presentations;
            pptObject = oPresSet.Open(@path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
        }
        
    }
}
