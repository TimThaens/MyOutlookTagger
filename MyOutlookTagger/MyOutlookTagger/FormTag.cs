using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Outlook = Microsoft.Office.Interop.Outlook;

namespace MyOutlookTagger
{
    public partial class FormTag : Form
    {
        #region Variables
        Outlook.MailItem _mail = null;
        Outlook.MeetingItem _meeting = null;
        string _origCats = "";
        #endregion

        #region Constructors
        public FormTag(Outlook.MailItem mail, bool defaultArchive = true)
        {
            InitializeComponent();
            cbArchive.Checked = defaultArchive;
            _mail = mail;
            setup(mail.Categories);
        }

        public FormTag(Outlook.MeetingItem meeting, bool defaultArchive = true)
        {
            InitializeComponent();
            cbArchive.Checked = defaultArchive;
            _meeting = meeting;
            setup(meeting.Categories);
        }
        #endregion

        #region Actions
        private void btnOK_Click(object sender, EventArgs e)
        {
            if (_mail != null)
                performOKMail();
            else if (_meeting != null)
                performOKMeeting();

            Close();
        }

        private void cbFollowUp_CheckedChanged(object sender, EventArgs e)
        {
            ddFollowUp.Visible = cbFollowUp.Checked;
        }

        private void tbCategories_KeyUp(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter || e.KeyCode == Keys.OemSemicolon || e.KeyCode == Keys.Oemcomma)
            {
                TagInputContainer.Controls.Add(enterNewTag(tbCategories.Text));
                tbCategories.Text = "";
            }
        }

        private void TagButtonX_Click(object sender, EventArgs e)
        {
            TextBox tag = (TextBox)((Button)sender).Parent;
            TagInputContainer.Controls.Remove(tag);
        }
        #endregion

        private void setup(string categories)
        {
            ddFollowUp.SelectedIndex = 0;
            _origCats = categories;

            var source = new AutoCompleteStringCollection();
            source.AddRange(TaggerMain.Instance.getAllExistingTags());
            tbCategories.AutoCompleteCustomSource = source;
            tbCategories.AutoCompleteMode = AutoCompleteMode.SuggestAppend;
            tbCategories.AutoCompleteSource = AutoCompleteSource.CustomSource;

            if (categories != null)
            {
                string[] cats = categories.Split(ThisAddIn.CATEGORY_SEPERATOR[0]);
                foreach (string cat in cats)
                {
                    TagInputContainer.Controls.Add(enterNewTag(cat.Trim()));
                }
            }
        }

        private TextBox enterNewTag(string text)
        {
            TextBox tag = new TextBox()
            {
                Width = 100,
                Height = 30,
                Font = new Font("Sergoe UI Light", 12),
                BorderStyle = BorderStyle.None,
                BackColor = Color.Khaki,
                Location = new Point(4, 0),
                Dock = DockStyle.Left,
                Margin = new Padding(30, 0, 30, 0),
                ReadOnly = true,
                Cursor = Cursors.Default
            };
            tag.Text = text.Replace(",", "");
            Size s = TextRenderer.MeasureText(tag.Text, tag.Font);
            tag.Width = s.Width + 35;

            Button btn = new Button()
            {
                Width = 30,
                Height = 30,
                Font = new Font("Sergoe UI Light", 12),
                BackColor = Color.LightGray,
                Dock = DockStyle.Right,
                Margin = new Padding(2, 0, 2, 0),
                Cursor = Cursors.Hand
            };
            btn.Parent = tag;
            btn.BringToFront();
            btn.Text = "X";
            btn.Click += new System.EventHandler(this.TagButtonX_Click);
            tag.Controls.Add(btn);

            return tag;
        }

        private string calculateNewCategories()
        {
            string newCats = "";
            foreach (object o in TagInputContainer.Controls)
            {
                if (o is TextBox)
                {
                    newCats += ThisAddIn.CATEGORY_SEPERATOR + ((TextBox)o).Text;
                }
            }
            return newCats;
        }



        private void performOKMail()
        {
            string newCats = calculateNewCategories();
            if (!newCats.Equals(_origCats))
                TaggerMain.Instance.setNewCategories(_mail, newCats);
            else
                TaggerMain.Instance.checkTagsExists(_origCats);

            if (cbFollowUp.Checked)
                TaggerMain.Instance.followupMail(_mail, ddFollowUp.SelectedIndex, cbArchive.Checked);
            else if (cbArchive.Checked)
                TaggerMain.Instance.moveToArchive(_mail);


        }

        private void performOKMeeting()
        {
            string newCats = calculateNewCategories();
            if (!newCats.Equals(_origCats))
                TaggerMain.Instance.setNewCategories(_meeting, newCats);

            if (cbFollowUp.Checked)
                TaggerMain.Instance.followupMeeting(_meeting, ddFollowUp.SelectedIndex, cbArchive.Checked);
            else if (cbArchive.Checked)
                TaggerMain.Instance.moveToArchive(_meeting);
        }
    }
}
