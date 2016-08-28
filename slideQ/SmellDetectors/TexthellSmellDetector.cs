using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using slideQ.Properties;
using slideQ.Model;
using slideQ;
using Microsoft.Office.Interop.PowerPoint;

namespace slideQ.SmellDetectors
{
    public class TexthellSmellDetector:ISmellDetector
    {
        private MasterDataModel dataModel;

        public TexthellSmellDetector(MasterDataModel dataModel)
        {
            this.dataModel = dataModel;
        }

        public List<PresentationSmell> detect()
        {
            List<PresentationSmell> smellList = new List<PresentationSmell>();
            foreach(SlideDataModel slide in dataModel.SlideDataModelList)
            {
                if(slide.TotalTextCount > Constants.TEXTHELL_THRESHOLD)
                {
                    PresentationSmell smell = new PresentationSmell();
                    smell.SmellName = Constants.TEXTHELL;
                    smell.Cause = "The tool detected the smell since the slide contains a lot of text (" + slide.TotalTextCount.ToString() + " characters).";
                    smell.SlideNo = slide.SlideNo;
                    smellList.Add(smell);
                }
            }
            return smellList;
        }
    }


}
