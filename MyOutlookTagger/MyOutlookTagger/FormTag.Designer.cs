namespace MyOutlookTagger
{
    partial class FormTag
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.tbCategories = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.cbArchive = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.cbFollowUp = new System.Windows.Forms.CheckBox();
            this.ddFollowUp = new System.Windows.Forms.ComboBox();
            this.TagInputContainer = new System.Windows.Forms.Panel();
            this.SuspendLayout();
            // 
            // tbCategories
            // 
            this.tbCategories.Location = new System.Drawing.Point(76, 6);
            this.tbCategories.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.tbCategories.Name = "tbCategories";
            this.tbCategories.Size = new System.Drawing.Size(148, 20);
            this.tbCategories.TabIndex = 0;
            this.tbCategories.KeyUp += new System.Windows.Forms.KeyEventHandler(this.tbCategories_KeyUp);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(10, 10);
            this.label1.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(31, 13);
            this.label1.TabIndex = 1;
            this.label1.Text = "Tags";
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(278, 172);
            this.btnOK.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(56, 19);
            this.btnOK.TabIndex = 2;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(12, 84);
            this.label2.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(49, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Archive?";
            // 
            // cbArchive
            // 
            this.cbArchive.AutoSize = true;
            this.cbArchive.Checked = true;
            this.cbArchive.CheckState = System.Windows.Forms.CheckState.Checked;
            this.cbArchive.Location = new System.Drawing.Point(76, 84);
            this.cbArchive.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cbArchive.Name = "cbArchive";
            this.cbArchive.Size = new System.Drawing.Size(15, 14);
            this.cbArchive.TabIndex = 4;
            this.cbArchive.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(12, 109);
            this.label3.Margin = new System.Windows.Forms.Padding(2, 0, 2, 0);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(58, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Follow up?";
            // 
            // cbFollowUp
            // 
            this.cbFollowUp.AutoSize = true;
            this.cbFollowUp.Location = new System.Drawing.Point(76, 109);
            this.cbFollowUp.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.cbFollowUp.Name = "cbFollowUp";
            this.cbFollowUp.Size = new System.Drawing.Size(15, 14);
            this.cbFollowUp.TabIndex = 6;
            this.cbFollowUp.UseVisualStyleBackColor = true;
            this.cbFollowUp.CheckedChanged += new System.EventHandler(this.cbFollowUp_CheckedChanged);
            // 
            // ddFollowUp
            // 
            this.ddFollowUp.FormattingEnabled = true;
            this.ddFollowUp.Items.AddRange(new object[] {
            "Today",
            "Tomorrow",
            "This Week",
            "Next Week",
            "This Month",
            "Some day"});
            this.ddFollowUp.Location = new System.Drawing.Point(93, 105);
            this.ddFollowUp.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.ddFollowUp.Name = "ddFollowUp";
            this.ddFollowUp.Size = new System.Drawing.Size(92, 21);
            this.ddFollowUp.TabIndex = 7;
            this.ddFollowUp.Visible = false;
            // 
            // TagInputContainer
            // 
            this.TagInputContainer.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.TagInputContainer.Cursor = System.Windows.Forms.Cursors.Default;
            this.TagInputContainer.Location = new System.Drawing.Point(14, 36);
            this.TagInputContainer.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.TagInputContainer.Name = "TagInputContainer";
            this.TagInputContainer.Size = new System.Drawing.Size(586, 43);
            this.TagInputContainer.TabIndex = 8;
            // 
            // FormTag
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(610, 201);
            this.Controls.Add(this.TagInputContainer);
            this.Controls.Add(this.ddFollowUp);
            this.Controls.Add(this.cbFollowUp);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.cbArchive);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.tbCategories);
            this.Margin = new System.Windows.Forms.Padding(2, 2, 2, 2);
            this.Name = "FormTag";
            this.Text = "Tagger";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox tbCategories;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOK;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.CheckBox cbArchive;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.CheckBox cbFollowUp;
        private System.Windows.Forms.ComboBox ddFollowUp;
        private System.Windows.Forms.Panel TagInputContainer;
    }
}