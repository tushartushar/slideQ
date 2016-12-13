using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using NUnit.Framework;
using slideQ.Model;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace SlideQTests
{
    class SmellTestOnEmptySlide
    {
        private _Presentation PPTObject;

        [SetUp]
        public void GetPPTObject()
        {
            string solution_dir = Path.GetDirectoryName(Path.GetDirectoryName(TestContext.CurrentContext.WorkDirectory));
            string path = @solution_dir + @"\TestFile\EmptySlide.pptx";
            string absolute = Path.GetFullPath(path);
            Application ppApp = new Application();
            ppApp.Visible = MsoTriState.msoTrue;
            Presentations oPresSet = ppApp.Presentations;
            PPTObject = oPresSet.Open(@path, MsoTriState.msoFalse, MsoTriState.msoFalse, MsoTriState.msoTrue);
        }


        [TearDown]
        public void tearDown()
        {
            PPTObject.Close();
        }

        [Test]
        public void RunAnalysisProcess()
        {
            List<SlideDataModel> SlideDataModelList = new List<SlideDataModel>();
            Assert.AreEqual(0, PPTObject.Slides.Count);
            foreach (Slide slide in PPTObject.Slides)
            {
                try
                {
                    SlideDataModel slideModel = new SlideDataModel(slide);
                    slideModel.build();
                    SlideDataModelList.Add(slideModel);
                }
                catch
                {
                    Assert.AreEqual(1, 2);
                }
            }

            Assert.AreEqual(0, SlideDataModelList.Count);
        }
    }
}
