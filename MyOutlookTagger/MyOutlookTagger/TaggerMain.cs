using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using Outlook = Microsoft.Office.Interop.Outlook;
using System.Data.Sql;
using System.Data.SqlClient;

namespace MyOutlookTagger
{
    sealed class TaggerMain
    {
        private static TaggerMain instance = new TaggerMain();
        private SqlConnection cn = null;
        private List<string> allCategories = new List<string>();

        private TaggerMain()
        {
            cn = new SqlConnection(global::MyOutlookTagger.Properties.Settings.Default.MyOutlookTaggerCS);
        }

        public static TaggerMain Instance {
            get { return instance;  }
        }

        public void init()
        {
            loadFolders();
            getAllTagsFromDB();
        }

        private void getAllTagsFromDB()
        {
            SqlCommand cmd = getSqlCommand("SELECT tag FROM TAGS");
            SqlDataReader r = cmd.ExecuteReader();
            while(r.Read())
            {
                string tag = (string)r["tag"];
                allCategories.Add(tag);
            }
        }

        private void loadFolders() { 
            SqlCommand cmd = getSqlCommand("SELECT count(id) FROM Settings WHERE [key] LIKE '%FolderMAPI' AND [value] IS NOT NULL");
            int num = (int)cmd.ExecuteScalar();
            while (num != 4)
            {
                System.Windows.Forms.Form f = new FormSettings();
                f.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
                f.ShowDialog();
                num = (int)cmd.ExecuteScalar();
            }

            cmd = getSqlCommand("SELECT [key],[value] FROM Settings WHERE [key] LIKE '%FolderMAPI'");
            SqlDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                loadFolder((string)r["key"], (string)r["value"]);
            }
        }

        public string[] getAllExistingTags()
        {
            return allCategories.ToArray();
        }

        public void setInboxFolder(string folderId)
        {
            Globals.ThisAddIn.setInbox(folderId);
            updateSettings("InboxFolderMAPI", folderId);
        }
        public void setSentboxFolder(string folderId)
        {
            Globals.ThisAddIn.setSentbox(folderId);
            updateSettings("SentFolderMAPI", folderId);
        }
        public void setArchiveFolder(string folderId)
        {
            Globals.ThisAddIn.setArchive(folderId);
            updateSettings("ArchiveFolderMAPI", folderId);
        }
        public void setFollowupFolder(string folderId)
        {
            Globals.ThisAddIn.setFollowup(folderId);
            updateSettings("FollowupFolderMAPI", folderId);
        }

        public void addNewAutoTagSender(string sender, string tag)
        {
            SqlCommand cmd = getSqlCommand("INSERT INTO Settings ([key], [value]) VALUES (@key, @value)");
            cmd.Parameters.AddWithValue("@key", "__ATS__" + sender);
            cmd.Parameters.AddWithValue("@value", "!" + tag.Trim('@', '!', '#', '.'));
            cmd.ExecuteNonQuery();
        }
        public void addNewAutoTagAttachement(string extension, string tag)
        {
            SqlCommand cmd = getSqlCommand("INSERT INTO Settings ([key], [value]) VALUES (@key, @value)");
            cmd.Parameters.AddWithValue("@key", "__ATA__" + extension.Trim('.'));
            cmd.Parameters.AddWithValue("@value", "@" + tag.Trim('@', '!', '#', '.'));
            cmd.ExecuteNonQuery();
        }
        public void delAutoTagSender(string sender)
        {
            SqlCommand cmd = getSqlCommand("DELETE FROM Settings WHERE [key] LIKE @key");
            cmd.Parameters.AddWithValue("@key", "__ATS__" + sender);
            cmd.ExecuteNonQuery();
        }
        public void delAutoTagAttachement(string extension)
        {
            SqlCommand cmd = getSqlCommand("DELETE FROM Settings WHERE [key] LIKE @key");
            cmd.Parameters.AddWithValue("@key", "__ATA__" + extension.Trim('.'));
            cmd.ExecuteNonQuery();
        }
        public List<AutoTagSender> getAllAutoTagSenderList()
        {
            List<AutoTagSender> list = new List<AutoTagSender>();

            SqlCommand cmd = getSqlCommand("SELECT * FROM Settings WHERE [key] LIKE '__ATS__%'");
            SqlDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                string sender = ((string)r["key"]).Remove(0, 7);
                string tag = (string)r["value"];
                list.Add(new AutoTagSender(sender, tag));
            }
            return list;
        }
        public List<AutoTagAttachment> getAllAutoTagAttachmentList()
        {
            List<AutoTagAttachment> list = new List<AutoTagAttachment>();

            SqlCommand cmd = getSqlCommand("SELECT * FROM Settings WHERE [key] LIKE '__ATA__%'");
            SqlDataReader r = cmd.ExecuteReader();
            while (r.Read())
            {
                string ext = "." + ((string)r["key"]).Remove(0, 7);
                string tag = (string)r["value"];
                list.Add(new AutoTagAttachment(ext, tag));
            }
            return list;
        }

        public void setNewCategories(Outlook.MailItem mail, string newCats)
        {
            checkTagsExists(newCats);
            if(mail != null)
            {
                mail.Categories = newCats;
                mail.Save();
                updateEntireConversation(mail);
            }
        }
        public void setNewCategories(Outlook.MeetingItem meeting, string newCats)
        {
            checkTagsExists(newCats);
            if(meeting != null)
            {
                meeting.Categories = newCats;
                meeting.Save();
            }
        }

        public void checkTagsExists(string categories)
        {
            string[] cats = categories.Split(ThisAddIn.CATEGORY_SEPERATOR[0]);
            foreach(string cat in cats)
            {
                string tcat = cat.Trim();
                if (tcat.Equals("") || tcat == null)
                    continue;

                // If the category doesn't exist in Outlook, add it.
                if(Globals.ThisAddIn.app.Session.Categories[tcat] == null)
                {
                    Outlook.Category ocat = Globals.ThisAddIn.app.Session.Categories.Add(tcat);
                    ocat.Color = Outlook.OlCategoryColor.olCategoryColorNone;                    
                }

                // If the category doesn't exist in the DB, add it.
                if(!allCategories.Contains(tcat))
                {
                    SqlCommand cmd = getSqlCommand("INSERT INTO Tags(tag) VALUES (@tag)");
                    cmd.Parameters.AddWithValue("@tag", tcat);
                    cmd.ExecuteNonQuery();
                    allCategories.Add(tcat);
                }
            }
        }
        public void moveToArchive(Outlook.MailItem mail)
        {
            mail.UnRead = false;
            Outlook.MAPIFolder curFolder = mail.Parent as Outlook.MAPIFolder;
            if (!curFolder.EntryID.Equals(Globals.ThisAddIn.archive.EntryID))
                mail.Move(Globals.ThisAddIn.archive);
        }
        public void moveToArchive(Outlook.MeetingItem meeting) { }
        public void moveToFollowup(Outlook.MailItem mail) { }
        public void moveToFollowup(Outlook.MeetingItem meeting) { }
        public void followupMail(Outlook.MailItem mail, int selection, bool archive) { }
        public void followupMeeting(Outlook.MeetingItem meeting, int selection, bool archive) { }

        private SqlCommand getSqlCommand(string query)
        {
            if (cn.State != System.Data.ConnectionState.Open)
                cn.Open();
            return new SqlCommand(query, cn);
        }

        public void loadFolder(string key, string folder)
        {
            switch (key)
            {
                case "InboxFolderMAPI":
                    setInboxFolder(folder);
                    break;
                case "SentFolderMAPI":
                    setSentboxFolder(folder);
                    break;
                case "FollowupFolderMAPI":
                    setFollowupFolder(folder);
                    break;
                case "ArchiveFolderMAPI":
                    setArchiveFolder(folder);
                    break;
            }
        }

        private object getSelectedObject()
        {
            if (Globals.ThisAddIn.app.ActiveExplorer().Selection.Count != 1)
                return null;
            return Globals.ThisAddIn.app.ActiveExplorer().Selection[1];
        }

        public void showSettings()
        {
            object selObject = getSelectedObject();
            string sender = "";
            if (selObject != null)
            {
                if (selObject is Outlook.MailItem)
                    sender = ((Outlook.MailItem)selObject).SenderName;
                else if (selObject is Outlook.MeetingItem)
                    sender = ((Outlook.MeetingItem)selObject).SenderName;
            }
            System.Windows.Forms.Form f = new FormSettings(sender);
            f.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            f.ShowDialog();
        }

        private void updateSettings(string key, string value)
        {
            SqlCommand cmd = getSqlCommand("UPDATE Settings SET [value]=@val WHERE [key] LIKE @key");
            cmd.Parameters.AddWithValue("@val", value);
            cmd.Parameters.AddWithValue("@key", key);
            cmd.ExecuteNonQuery();
        }

        public string getPersistantCategories(string categories)
        {
            string persCats = "";
            if (categories == null || categories.Equals(""))
                return "";
            string[] cats = categories.Split(ThisAddIn.CATEGORY_SEPERATOR[0]);
            foreach (string cat in cats)
            {
                if (!cat.Trim().StartsWith("!") && !cat.Trim().StartsWith("@") && !cat.Trim().StartsWith("#"))
                    persCats += ThisAddIn.CATEGORY_SEPERATOR + cat.Trim();
            }

            return persCats;
        }

        private void updateEntireConversation(Outlook.MailItem mail)
        {
            string persTags = getPersistantCategories(mail.Categories);
            Outlook.Conversation conv = mail.GetConversation();
            if (conv == null)
                return;

            Outlook.SimpleItems simpleItems = conv.GetRootItems();
            foreach(object item in simpleItems)
            {
                if(item is Outlook.MailItem smail)
                {
                    string cats = persTags;
                    cats += getSenderCategory(smail.SenderName);
                    cats += getAttachmentCategories(mail.Attachments);
                    smail.Categories = getUniqueTags(cats);
                    smail.Save();
                }

                Outlook.SimpleItems OlChildren = conv.GetChildren(item);
                foreach (object rci in OlChildren)
                {
                    if (rci is Outlook.MailItem icr)
                    {
                        string cats = persTags;
                        cats += getSenderCategory(icr.SenderName);
                        cats += getAttachmentCategories(mail.Attachments);
                        icr.Categories = getUniqueTags(cats);
                        icr.Save();
                    }
                }
            }
        }

        public void newMailItemAdded(Outlook.MailItem mail)
        {
            string cats = mail.Categories;
            cats += getAttachmentCategories(mail.Attachments);
            cats += getSenderCategory(mail.SenderName);
            cats += getConversationCategories(mail);
            mail.Categories = getUniqueTags(cats);
            mail.Save();
        }

        private string getConversationCategories(Outlook.MailItem mail)
        {
            string convCats = "";
            Outlook.Conversation conv = mail.GetConversation();
            if (conv == null)
                return "";

            Outlook.SimpleItems simpleItems = conv.GetRootItems();
            foreach(object item in simpleItems)
            {
                if(item is Outlook.MailItem)
                {
                    Outlook.MailItem smail = item as Outlook.MailItem;
                    if (smail.Categories == null)
                        continue;
                    string[] scats = smail.Categories.Split(ThisAddIn.CATEGORY_SEPERATOR[0]);
                    foreach(string scat in scats)
                    {
                        string tscat = scat.Trim();
                        if (!tscat.StartsWith("!") && !tscat.StartsWith("@") && !tscat.StartsWith("#"))
                            convCats += ThisAddIn.CATEGORY_SEPERATOR + tscat;
                    }
                }
            }

            return convCats;
        }


        private string getAttachmentCategories(Outlook.Attachments atts)
        {
            List<string> cats = new List<string>();
            List<AutoTagAttachment> atas = getAllAutoTagAttachmentList();
            string attcats = "";

            foreach (Outlook.Attachment att in atts)
            {
                string pat = @"image\d\d\d\.[png|jpg]";
                Regex r = new Regex(pat, RegexOptions.IgnoreCase);
                try
                {
                    if (att.Type == Outlook.OlAttachmentType.olEmbeddeditem || att.Type == Outlook.OlAttachmentType.olOLE)
                        continue;
                    // Those annoying imbedded images are all named imageXXX.jpg/png
                    Match m = r.Match(att.FileName);
                    if (m.Success)
                        continue;
                    String ext = System.IO.Path.GetExtension(att.FileName);
                    AutoTagAttachment result = atas.Find(ata => ata.extension.Equals(ext));
                    if (result != null)
                        cats.Add(result.tag);
                }
                catch (System.Exception e) { System.Windows.Forms.MessageBox.Show(att.Type.ToString() + "\n" + e.ToString(), "Error - setAttachmentCategories"); }
            }

            var enumCats = cats.Distinct();
            foreach (var item in enumCats)
            {
                attcats += ThisAddIn.CATEGORY_SEPERATOR + (string)item;
            }
            return attcats;
        }

        private string getSenderCategory(string sender)
        {
            string key = "__ATS__" + sender;
            SqlCommand cmd = getSqlCommand("SELECT [value] FROM Settings WHERE [key] LIKE @sender");
            cmd.Parameters.AddWithValue("@sender", key);
            object retVal = cmd.ExecuteScalar();
            if (retVal != null)
                return ThisAddIn.CATEGORY_SEPERATOR + (string)retVal;
            return "";
        }

        private string getUniqueTags(string categories)
        {
            if (categories == null || categories.Equals(""))
                return "";

            string uniqueCats = "";

            string[] cats = categories.Replace(", ", ",").Split(ThisAddIn.CATEGORY_SEPERATOR[0]);

            var enumCats = cats.Distinct();
            foreach (var item in enumCats)
            {
                uniqueCats += ThisAddIn.CATEGORY_SEPERATOR + (string)item;
            }

            return uniqueCats;
        }

    }
}
