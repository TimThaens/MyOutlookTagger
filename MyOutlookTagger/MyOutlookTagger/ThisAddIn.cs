using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Xml.Linq;
using Outlook = Microsoft.Office.Interop.Outlook;
using Office = Microsoft.Office.Core;

namespace MyOutlookTagger
{
    public partial class ThisAddIn
    {
        public const string CATEGORY_SEPERATOR = ",";

        #region Variables
        public Outlook.Application app = null;
        public Outlook.NameSpace ns = null;
        public Outlook.MAPIFolder inbox = null;
        public Outlook.MAPIFolder sentbox = null;
        public Outlook.MAPIFolder archive = null;
        public Outlook.MAPIFolder followup = null;

        private Outlook.Items _inboxItems = null;
        private Outlook.Items _sentboxItems = null;
        #endregion

        #region Setup
        private void ThisAddIn_Startup(object sender, System.EventArgs e)
        {
            this.app = this.Application as Outlook.Application;
            this.ns = this.Application.GetNamespace("MAPI");

            this.app.ItemSend += new Outlook.ApplicationEvents_11_ItemSendEventHandler(_app_ItemSend);

            TaggerMain.Instance.init();
        }

        private void _app_ItemSend(object Item, ref bool Cancel)
        {
            // [TODO] Handle mail when being send
        }

        public void setInbox(string folderId)
        {
            inbox = ns.GetFolderFromID(folderId);
            _inboxItems = inbox.Items;
            _inboxItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(inboxItems_ItemAdd);
        }

        public void setSentbox(string folderId)
        {
            sentbox = ns.GetFolderFromID(folderId);
            _sentboxItems = sentbox.Items;
            _sentboxItems.ItemAdd += new Outlook.ItemsEvents_ItemAddEventHandler(sentboxItems_ItemAdd);
        }

        public void setArchive(string folderId)
        {
            archive = ns.GetFolderFromID(folderId);
        }

        public void setFollowup(string folderId)
        {
            followup = ns.GetFolderFromID(folderId);
        }

        public void inboxItems_ItemAdd(object Item)
        {
            if (Item is Outlook.MailItem)
                TaggerMain.Instance.newMailItemAdded((Outlook.MailItem)Item);
        }

        public void sentboxItems_ItemAdd(object Item)
        {
            // [TODO] Handle an item that enters the sentbox
        }

        private void ThisAddIn_Shutdown(object sender, System.EventArgs e)
        {
            // Note: Outlook no longer raises this event. If you have code that 
            //    must run when Outlook shuts down, see https://go.microsoft.com/fwlink/?LinkId=506785
        }
        #endregion

        #region Tagger
        public void Tagger()
        {
            // If not a single mail or meeting is selection, the Tagger box won't open
            if (this.Application.ActiveExplorer().Selection.Count != 1)
                return;
            object selObject = this.Application.ActiveExplorer().Selection[1];

            // If the selected Item is a MailItem, open the Tagger Box, and check 'Archive' by default.
            if(selObject is Outlook.MailItem)
            {
                Outlook.MailItem mail = selObject as Outlook.MailItem;
                System.Windows.Forms.Form f = new FormTag(mail, true);
                f.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
                f.ShowDialog();
            }
            // If the selected Item is a MeetingItem, open the Tagger Box, and don't check 'Archive' by default.
            else if(selObject is Outlook.MeetingItem)
            {
                Outlook.MeetingItem meeting = selObject as Outlook.MeetingItem;
                System.Windows.Forms.Form f = new FormTag(meeting, false);
                f.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
                f.ShowDialog();
            }

        }
        #endregion

        #region Ribbon Settings
        protected override Microsoft.Office.Core.IRibbonExtensibility CreateRibbonExtensibilityObject()
        {
            return new RibbonMain();
        }
        #endregion
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
