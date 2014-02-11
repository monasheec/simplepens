using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;
using Office = Microsoft.Office.Core;
using System.Windows.Forms;
using Microsoft.Office.Interop.PowerPoint;
using Microsoft.Office.Core;
using System.Windows.Input;

namespace PowerpointJabber
{
    public partial class ThisAddIn
    {
        public static ThisAddIn instance;
        public SimplePenWindow SSSW;
        public bool customPresenterIsEnabledForPresenterMode
        {
            get
            {
                Properties.Settings.Default.Reload();
                return Properties.Settings.Default.SimplePensEnabledForPresenterMode;
            }
            set
            {
                Properties.Settings.Default.SimplePensEnabledForPresenterMode = value;
                Properties.Settings.Default.Save();
            }
        }
        public bool customPresenterIsEnabledForDefaultMode
        {
            get
            {
                Properties.Settings.Default.Reload();
                return Properties.Settings.Default.SimplePensEnabledForDefaultMode;
            }
            set
            {
                Properties.Settings.Default.SimplePensEnabledForDefaultMode = value;
                Properties.Settings.Default.Save();
            }
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            instance = this;
            this.Application.SlideShowBegin += onSlideShowBegin;
            this.Application.SlideShowEnd += onSlideShowEnd;
        }
        private void onSlideShowBegin(object sender)
        {
            if (WindowsInteropFunctions.presenterActive)
            {
                if (customPresenterIsEnabledForPresenterMode)
                {
                    SSSW = new SimplePenWindow();
                    SSSW.Show();
                }
            }
            else
            {
                if (customPresenterIsEnabledForDefaultMode)
                {
                    SSSW = new SimplePenWindow();
                    SSSW.Show();
                }
            }
        }
        private void onSlideShowEnd(object sender)
        {
            if (SSSW != null)
            {
                SSSW.Close();
                SSSW = null;
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (SSSW != null)
            {
                SSSW.Close();
                SSSW = null;
            }
            ThisAddIn.instance = null;
        }
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new Ribbon1();
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
