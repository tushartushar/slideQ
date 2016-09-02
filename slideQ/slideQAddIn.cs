using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using PowerPointAddInTest_CountSlides;

namespace slideQ
{
    using PowerPoint = Microsoft.Office.Interop.PowerPoint;
    public partial class slideQAddIn
    {
        private PaneBackWinControl WinControl;
        private static Microsoft.Office.Tools.CustomTaskPane TaskPaneObj;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            WinControl = new PaneBackWinControl();
            TaskPaneObj = this.CustomTaskPanes.Add(WinControl, "Presentation smells - slideQ");
            TaskPaneObj.Visible = false;
            TaskPaneObj.VisibleChanged += TaskPane_VisibleChanged; 
        }

        void TaskPane_VisibleChanged(object sender, EventArgs e)
        {
            //throw new NotImplementedException();
        }

        public Microsoft.Office.Tools.CustomTaskPane TaskPane
        {
            get
            {
                return TaskPaneObj;
            }
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon();
        }

    

        #region VSTO generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InternalStartup()
        {
            this.Startup += new System.EventHandler(ThisAddIn_Startup);
            this.Shutdown += new System.EventHandler(ThisAddIn_Shutdown);
        }
        
        #endregion
    }
}
