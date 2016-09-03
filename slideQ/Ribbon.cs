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
using slideQ.SmellDetectors;
using slideQ.Model;
using PowerPointAddInTest_CountSlides;
using Microsoft.Office.Core;
using slideQ.View;

// TODO:  Follow these steps to enable the Ribbon (XML) item:

// 1: Copy the following code block into the ThisAddin, ThisWorkbook, or ThisDocument class.

//  protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
//  {
//      return new Ribbon();
//  }

// 2. Create callback methods in the "Ribbon Callbacks" region of this class to handle user
//    actions, such as clicking a button. Note: if you have exported this Ribbon from the Ribbon designer,
//    move your code from the event handlers to the callback methods and modify the code to work with the
//    Ribbon extensibility (RibbonX) programming model.

// 3. Assign attributes to the control tags in the Ribbon XML file to identify the appropriate callback methods in your code.  

// For more information, see the Ribbon XML documentation in the Visual Studio Tools for Office Help.


namespace slideQ
{
    [ComVisible(true)]
    public class Ribbon : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public Ribbon()
        {
        }

        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("PowerPointAddInTest_CountSlides.Ribbon.xml");
        }

        #endregion

        #region Ribbon Callbacks
        //Create callback methods here. For more information about adding callback methods, visit http://go.microsoft.com/fwlink/?LinkID=271226

        public void Ribbon_Load(Office.IRibbonUI ribbonUI)
        {
            this.ribbon = ribbonUI;
        }

        #endregion

        #region Helpers

        private static string GetResourceText(string resourceName)
        {
            Assembly asm = Assembly.GetExecutingAssembly();
            string[] resourceNames = asm.GetManifestResourceNames();
            for (int i = 0; i < resourceNames.Length; ++i)
            {
                if (string.Compare(resourceName, resourceNames[i], StringComparison.OrdinalIgnoreCase) == 0)
                {
                    using (StreamReader resourceReader = new StreamReader(asm.GetManifestResourceStream(resourceNames[i])))
                    {
                        if (resourceReader != null)
                        {
                            return resourceReader.ReadToEnd();
                        }
                    }
                }
            }
            return null;
        }

        public void AddAnimationButtonClick(Office.IRibbonControl control)
        {
            try
            {
                //MessageBox.Show("Count of Slides: " + Globals.slideQAddIn.Application.ActivePresentation.Slides.Count, Constants.AppName);
                
                SmellDetector detector = new SmellDetector();
                List<PresentationSmell> presentationSmells = detector.detectPresentationSmells(Globals.slideQAddIn.Application.ActivePresentation.Slides);
                SmellDisplayControl.PPTSmellList.ItemsSource = presentationSmells;
                SmellDisplayControl.PPTSmellList.UpdateLayout();
                Globals.slideQAddIn.TaskPane.Visible = true;

            }
            catch (Exception)
            {
               //Log the exception
            }
        }


        public static void Gotoslide(int index)
        {
            try
            {
                Globals.slideQAddIn.Application.ActivePresentation.Slides[index].Select();
            }
            catch(COMException)
            {
                Globals.slideQAddIn.Application.ActivePresentation.Slides[index - Globals.slideQAddIn.Application.ActivePresentation.Slides[1].SlideNumber+1].Select();
            }
          
        }
   

        #endregion
        #region GetIcon
        public Bitmap GetImage(Office.IRibbonControl control)
        {
            return new Bitmap(PowerPointAddInTest_CountSlides.Properties.Resources.icon);
        }
        #endregion
    }
}
