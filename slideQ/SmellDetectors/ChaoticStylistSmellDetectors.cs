using slideQ.Model;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;

namespace slideQ.SmellDetectors
{
    class ChaoticStylistSmellDetectors : ISmellDetector
    {
          private MasterDataModel dataModel;

          public ChaoticStylistSmellDetectors(MasterDataModel dataModel)
          {
              this.dataModel = dataModel;
          }

          public List<PresentationSmell> detect()
          {
              List<PresentationSmell> smellList = new List<PresentationSmell>();
              List<List<TextStyle>> CharAttrList=  dataModel.SlideDataModelList.Select(x => x.TextStlyeList).ToList();
              List<TextStyle> AllCharObject = new List<TextStyle>();
              foreach (List<TextStyle> item in CharAttrList)
              {
                  AllCharObject.AddRange(item);
              }
              
              List<CustumTextStyle> CustomList = AllCharObject.Select(x => new CustumTextStyle { Size = x.Size, Color = x.Color, FontNameofChar = x.FontName }).ToList();
              int Count = CustomList.GroupBy(x => new { x.Color, x.FontNameofChar, x.Size }).Select(x => x.FirstOrDefault()).Count();
         
              if (Count > Constants.CHAOTIC_STYLIST_THRESHOLD)
              {
                  PresentationSmell smell = new PresentationSmell();
                  smell.SmellName = Constants.CHAOTIC_STYLIST;
                  string Cause = "The tool detected the smell since the slides contains " + Count + " different styles.";
                  smell.Cause = Cause;
                  smell.SlideNo = 1;
                  smellList.Add(smell);
              }
              return smellList;
          }
    }

    class CustumTextStyle
    {
        public float Size { get; set; }
        public int Color { get; set; }
        public string FontNameofChar { get; set; }
    }
}
