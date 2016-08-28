using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.PowerPoint;

namespace slideQ.Model
{
    public class MasterDataModel
    {
        private Slides Slides;

        public MasterDataModel(Slides Slides)
        {
            this.Slides = Slides;
            SlideDataModelList = new List<SlideDataModel>();
        }
        public void build()
        {
            foreach (Slide slide in Slides)
            {
                SlideDataModel slideModel = new SlideDataModel(slide);
                slideModel.build();
                SlideDataModelList.Add(slideModel);
            }
        }

        public List<SlideDataModel> SlideDataModelList { get; set; }
    }
}
