using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Office = Microsoft.Office.Core;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MyOutlookTagger
{
    public partial class FormSettings : Form
    {
        #region Variables
        private Outlook.MAPIFolder _tmpInbox = null;
        private Outlook.MAPIFolder _tmpSent = null;
        private Outlook.MAPIFolder _tmpFollowUp = null;
        private Outlook.MAPIFolder _tmpArchive = null;
        #endregion

        #region Creator/Loader
        public FormSettings(string sender = "")
        {
            InitializeComponent();
            tabControl1.DrawItem += new DrawItemEventHandler(tabControl1_DrawItem);

            ScrollBar vScrollBar3 = new VScrollBar();
            vScrollBar3.Dock = DockStyle.Right;
            vScrollBar3.Scroll += (sender1, e) => { pnlATA.VerticalScroll.Value = vScrollBar3.Value; };
            pnlATA.Controls.Add(vScrollBar3);

            var source = new AutoCompleteStringCollection();
            source.AddRange(TaggerMain.Instance.getAllExistingTags());
            txtNewTag.AutoCompleteCustomSource = source;
            txtNewTag.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            txtNewTag.AutoCompleteSource = AutoCompleteSource.CustomSource;
            txtNewSender.Text = sender;
            getConfig();
            getConfig2();

        }

        private void SettingsForm_Load(object sender, EventArgs e)
        {
            textBox1.Text = Globals.ThisAddIn.inbox == null ? "" : Globals.ThisAddIn.inbox.Name;
            textBox2.Text = Globals.ThisAddIn.sentbox == null ? "" : Globals.ThisAddIn.sentbox.Name;
            textBox3.Text = Globals.ThisAddIn.followup == null ? "" : Globals.ThisAddIn.followup.Name;
            textBox4.Text = Globals.ThisAddIn.archive == null ? "" : Globals.ThisAddIn.archive.Name;
        }

        private void tabControl1_DrawItem(object sender, DrawItemEventArgs e)
        {
            Graphics g = e.Graphics;
            Brush _textBrush;

            // Get the item from the collection.
            TabPage _tabPage = tabControl1.TabPages[e.Index];

            // Get the real bounds for the tab rectangle
            Rectangle _tabBounds = tabControl1.GetTabRect(e.Index);

            if (e.State == DrawItemState.Selected)
            {
                // Draw a different background color, and don't paint a focus rectangle
                _textBrush = new SolidBrush(Color.Red);
                g.FillRectangle(Brushes.Gray, e.Bounds);
            }
            else
            {
                _textBrush = new System.Drawing.SolidBrush(e.ForeColor);
                e.DrawBackground();
            }

            // Use our own font
            Font _tabFont = new Font("Arial", (float)10.0, FontStyle.Bold, GraphicsUnit.Pixel);

            // Draw string. Center the text
            StringFormat _stringFlags = new StringFormat();
            _stringFlags.Alignment = StringAlignment.Center;
            _stringFlags.LineAlignment = StringAlignment.Center;
            g.DrawString(_tabPage.Text, _tabFont, _textBrush, _tabBounds, new StringFormat(_stringFlags));
        }
        #endregion

        #region Folder Selects
        private void button1_Click(object sender, EventArgs e)
        {
            Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.PickFolder();
            if (folder != null)
            {
                textBox1.Text = folder.Name;
                _tmpInbox = folder;
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.PickFolder();
            if (folder != null)
            {
                textBox4.Text = folder.Name;
                _tmpArchive = folder;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.PickFolder();
            if (folder != null)
            {
                textBox3.Text = folder.Name;
                _tmpFollowUp = folder;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Outlook.MAPIFolder folder = Globals.ThisAddIn.Application.Session.PickFolder();
            if (folder != null)
            {
                textBox2.Text = folder.Name;
                _tmpSent = folder;
            }
        }
        #endregion

        #region OK - Close buttons
        private void button6_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (_tmpInbox != null)
                TaggerMain.Instance.setInboxFolder(_tmpInbox.EntryID);
            if (_tmpSent != null)
                TaggerMain.Instance.setSentboxFolder(_tmpSent.EntryID);
            if (_tmpFollowUp != null)
                TaggerMain.Instance.setFollowupFolder(_tmpFollowUp.EntryID);
            if (_tmpArchive != null)
                TaggerMain.Instance.setArchiveFolder(_tmpArchive.EntryID);
            Close();
        }
        #endregion

        #region AutoTagSender
        private void btnAdd_Click(object sender, EventArgs e)
        {
            if (txtNewSender.Text.Equals("") || txtNewTag.Text.Equals(""))
            {
                MessageBox.Show("Please fill in a Sender and Tag", "New Auto Tag Sender");
                return;
            }

            TaggerMain.Instance.addNewAutoTagSender(txtNewSender.Text, txtNewTag.Text);
            getConfig();

        }

        private void getConfig()
        {
            List<AutoTagSender> list = TaggerMain.Instance.getAllAutoTagSenderList();
            this.pnlATS.Controls.Clear();
            int i = 0;
            foreach (AutoTagSender ats in list)
            {
                Label lblSender = new Label();
                Label lblTag = new Label();
                Button btnDelete = new Button();

                lblSender.Text = ats.sender;
                lblTag.Text = ats.tag;
                btnDelete.Text = "X";

                lblSender.Top = 3 + i * (3 + btnDelete.Height);
                lblTag.Top = 3 + i * (3 + btnDelete.Height);
                btnDelete.Top = 3 + i * (3 + btnDelete.Height);

                lblSender.Left = 3;
                lblTag.Left = 3 + (pnlATS.Width - 9) / 2;
                btnDelete.Left = pnlATS.Width - 3 - btnDelete.Width;

                lblSender.Width = lblTag.Left - 6;
                lblTag.Width = btnDelete.Left - lblTag.Left - 6;

                btnDelete.Tag = ats.sender;
                btnDelete.Click += new EventHandler(atsDeleteBtn_Click);

                pnlATS.Controls.Add(lblSender);
                pnlATS.Controls.Add(lblTag);
                pnlATS.Controls.Add(btnDelete);


                i++;
            }
        }
        void atsDeleteBtn_Click(Object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            string s = (string)btn.Tag;
            TaggerMain.Instance.delAutoTagSender(s);
            getConfig();
        }
        #endregion

        #region AutoTagAttachment
        private void btnAdd2_Click(object sender, EventArgs e)
        {
            if (txtNewAttach.Text.Equals("") || txtNewTag2.Text.Equals(""))
            {
                MessageBox.Show("Please fill in an Extension and Tag", "New Auto Tag Attachment");
                return;
            }

            TaggerMain.Instance.addNewAutoTagAttachement(txtNewAttach.Text, txtNewTag2.Text);
            getConfig2();

        }

        private void getConfig2()
        {
            List<AutoTagAttachment> list = TaggerMain.Instance.getAllAutoTagAttachmentList();
            this.pnlATA.Controls.Clear();
            int i = 0;
            foreach (AutoTagAttachment ata in list)
            {
                Label lblExtension = new Label();
                Label lblTag = new Label();
                Button btnDelete = new Button();

                lblExtension.Text = ata.extension;
                lblTag.Text = ata.tag;
                btnDelete.Text = "X";

                lblExtension.Top = 3 + i * (3 + btnDelete.Height);
                lblTag.Top = 3 + i * (3 + btnDelete.Height);
                btnDelete.Top = 3 + i * (3 + btnDelete.Height);

                lblExtension.Left = 3;
                lblTag.Left = 3 + (pnlATA.Width - 9) / 2;
                btnDelete.Left = pnlATA.Width - 3 - btnDelete.Width;

                lblExtension.Width = lblTag.Left - 6;
                lblTag.Width = btnDelete.Left - lblTag.Left - 6;

                btnDelete.Tag = ata.extension;
                btnDelete.Click += new EventHandler(ataDeleteBtn_Click);

                pnlATA.Controls.Add(lblExtension);
                pnlATA.Controls.Add(lblTag);
                pnlATA.Controls.Add(btnDelete);


                i++;
            }
        }

        void ataDeleteBtn_Click(Object sender, EventArgs e)
        {
            Button btn = (Button)sender;
            string ext = (string)btn.Tag;
            TaggerMain.Instance.delAutoTagAttachement(ext);
            getConfig2();
        }
        #endregion



    }

    public class AutoTagSender
    {
        public string sender;
        public string tag;

        public AutoTagSender(string sender, string tag)
        {
            this.sender = sender;
            this.tag = tag;
        }
    }

    public class AutoTagAttachment
    {
        public string extension;
        public string tag;

        public AutoTagAttachment(string extension, string tag)
        {
            this.extension = extension;
            this.tag = tag;
        }
    }
}
