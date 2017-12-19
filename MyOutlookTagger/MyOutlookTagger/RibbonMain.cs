using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Runtime.InteropServices;
using System.Text;
using Office = Microsoft.Office.Core;

namespace MyOutlookTagger
{
    [ComVisible(true)]
    public class RibbonMain : Office.IRibbonExtensibility
    {
        private Office.IRibbonUI ribbon;

        public RibbonMain()
        {
        }

        public void onSearchButton(Office.IRibbonControl control)
        {
            // [TODO] open a new SearchBox
        }
        public void onSettingsButton(Office.IRibbonControl control)
        {
            TaggerMain.Instance.showSettings();
        }
        public void onTagButton(Office.IRibbonControl control)
        {
            Globals.ThisAddIn.Tagger();
        }
        
        #region IRibbonExtensibility Members

        public string GetCustomUI(string ribbonID)
        {
            return GetResourceText("MyOutlookTagger.RibbonMain.xml");
        }

        #endregion

        #region Ribbon Callbacks
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

        #endregion
    }
}
