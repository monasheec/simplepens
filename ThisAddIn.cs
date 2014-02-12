using System;
using System.Collections;
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
        private static string _version;
        public static string version
        {
            get
            {
                if (!String.IsNullOrEmpty(_version)) return _version;
                else
                {
                    var tempVersion = "SimplePens PowerPoint " + ThisAddIn.instance.Application.Version;
                    if (!String.IsNullOrEmpty(tempVersion))
                    {
                        _version = tempVersion;
                        return _version;
                    }
                    else
                        return "unknown";
                }
            }
        }
        public bool customPresenterIsEnabledForPresenterMode
        {
            get
            {
                Properties.Settings.Default.Reload();
                return Properties.Settings.Default.SimplePensEnabledForPresenterMode;
            }
            set
            {
                Logger.Info("Setting SimplePens enabled for presenter presentation mode");
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
                Logger.Info("Setting SimplePens enabled for default presentation mode");
                Properties.Settings.Default.SimplePensEnabledForDefaultMode = value;
                Properties.Settings.Default.Save();
            }
        }
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            instance = this;
            //Logger.StartLogger();
            this.Application.SlideShowBegin += onSlideShowBegin;
            this.Application.SlideShowEnd += onSlideShowEnd;
        }
        private void onSlideShowBegin(object sender)
        {
            if (WindowsInteropFunctions.presenterActive)
            {
                if (customPresenterIsEnabledForPresenterMode)
                {
                    Logger.Info("starting simplePens for presenter presentation mode");
                    SSSW = new SimplePenWindow();
                    SSSW.Show();
                }
            }
            else
            {
                if (customPresenterIsEnabledForDefaultMode)
                {
                    Logger.Info("starting simplePens for presenter presentation mode");
                    SSSW = new SimplePenWindow();
                    SSSW.Show();
                }
            }
        }
        private void onSlideShowEnd(object sender)
        {
            if (SSSW != null)
            {
                Logger.Info("Slideshow ended");
                SSSW.Close();
                SSSW = null;
            }
        }
        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            if (SSSW != null)
            {
                Logger.Info("Shutting down SimplePens");
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
