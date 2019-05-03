using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;
using System.Diagnostics;

namespace VSTO_AddIn
{
    public partial class ThisAddIn
    {

        // TO DO //
        /* ADD WAPPERS FOR THE FOLLOWING TYPES:
         * Microsoft.Office.Interop.Outlook.AccountsEvents
         * Microsoft.Office.Interop.Outlook._DDocSiteControlEvents
         * Microsoft.Office.Interop.Outlook._DRecipientControlEvents
         * Microsoft.Office.Interop.Outlook.FormRegionEvents
         * Microsoft.Office.Interop.Outlook.NavigationGroupsEvents_12
         * Microsoft.Office.Interop.Outlook.NavigationPaneEvents_12
         * Microsoft.Office.Interop.Outlook.OlkBusinessCardControlEvents
         * Microsoft.Office.Interop.Outlook.OlkCategoryEvents
         * Microsoft.Office.Interop.Outlook.OlkCheckBoxEvents
         * Microsoft.Office.Interop.Outlook.OlkComboBoxEvents
         * Microsoft.Office.Interop.Outlook.OlkCommandButtonEvents
         * Microsoft.Office.Interop.Outlook.OlkContactPhotoEvents
         * Microsoft.Office.Interop.Outlook.OlkDateControlEvents
         * Microsoft.Office.Interop.Outlook.OlkFrameHeaderEvents
         * Microsoft.Office.Interop.Outlook.OlkInfoBarEvents
         * Microsoft.Office.Interop.Outlook.OlkLabelEvents
         * Microsoft.Office.Interop.Outlook.OlkListBoxEvents
         * Microsoft.Office.Interop.Outlook.OlkOptionButtonEvents
         * Microsoft.Office.Interop.Outlook.OlkPageControlEvents
         * Microsoft.Office.Interop.Outlook.OlkTextBoxEvents
         * Microsoft.Office.Interop.Outlook.OlkTimeControlEvents
         * Microsoft.Office.Interop.Outlook.OlkTimeZoneControlEvents
         * Microsoft.Office.Interop.Outlook.OutlookBarGroupsEvents
         * Microsoft.Office.Interop.Outlook.OutlookBarPaneEvents
         * Microsoft.Office.Interop.Outlook.OutlookBarShortcutsEvents
         * Microsoft.Office.Interop.Outlook.ReminderCollectionEvents
         * Microsoft.Office.Interop.Outlook.ResultsEvents
         * Microsoft.Office.Interop.Outlook.SyncObjectEvents
         * Microsoft.Office.Interop.Outlook._ViewsEvents
         */

        internal ControlPanel controlPanel1;
        private Microsoft.Office.Tools.CustomTaskPane controlPanelTaskPane;

        ApplicationWrapper applicationWrapper;
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            controlPanel1 = new ControlPanel();
            controlPanelTaskPane = this.CustomTaskPanes.Add(controlPanel1, "VSTO_AddIn Control Panel");
            controlPanelTaskPane.Width = 400;
            controlPanelTaskPane.Visible = true;
            applicationWrapper = new ApplicationWrapper(Application);
        }



        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
            applicationWrapper.Dispose();
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
