using System;
using System.Windows.Forms;

namespace OutlookAI
{
    /// <summary>
    /// Form for editing email category settings
    /// </summary>
    public partial class CategoryEditorForm : Form
    {
        private EmailCategory _category;
        private bool _isNewCategory;

        public EmailCategory Category
        {
            get { return _category; }
            set
            {
                _category = value;
                LoadCategoryData();
            }
        }

        public CategoryEditorForm(EmailCategory category = null)
        {
            InitializeComponent();

            if (category == null)
            {
                _category = new EmailCategory();
                _isNewCategory = true;
                this.Text = "New Email Category";
            }
            else
            {
                _category = category.Clone();
                _isNewCategory = false;
                this.Text = "Edit Email Category";
            }

            LoadCategoryData();
        }

        private void InitializeComponent()
        {
            this.labelCategoryName = new System.Windows.Forms.Label();
            this.textBoxCategoryName = new System.Windows.Forms.TextBox();
            this.labelDescription = new System.Windows.Forms.Label();
            this.textBoxDescription = new System.Windows.Forms.TextBox();
            this.labelClassificationPrompt = new System.Windows.Forms.Label();
            this.textBoxClassificationPrompt = new System.Windows.Forms.TextBox();
            this.checkBoxGenerateReply = new System.Windows.Forms.CheckBox();
            this.labelReplyPrompt = new System.Windows.Forms.Label();
            this.textBoxReplyPrompt = new System.Windows.Forms.TextBox();
            this.checkBoxIsEnabled = new System.Windows.Forms.CheckBox();
            this.buttonOK = new System.Windows.Forms.Button();
            this.buttonCancel = new System.Windows.Forms.Button();
            this.SuspendLayout();
            //
            // labelCategoryName
            //
            this.labelCategoryName.AutoSize = true;
            this.labelCategoryName.Location = new System.Drawing.Point(12, 15);
            this.labelCategoryName.Name = "labelCategoryName";
            this.labelCategoryName.Size = new System.Drawing.Size(82, 13);
            this.labelCategoryName.TabIndex = 0;
            this.labelCategoryName.Text = "Category Name:";
            //
            // textBoxCategoryName
            //
            this.textBoxCategoryName.Location = new System.Drawing.Point(15, 31);
            this.textBoxCategoryName.Name = "textBoxCategoryName";
            this.textBoxCategoryName.Size = new System.Drawing.Size(557, 20);
            this.textBoxCategoryName.TabIndex = 1;
            //
            // labelDescription
            //
            this.labelDescription.AutoSize = true;
            this.labelDescription.Location = new System.Drawing.Point(12, 60);
            this.labelDescription.Name = "labelDescription";
            this.labelDescription.Size = new System.Drawing.Size(63, 13);
            this.labelDescription.TabIndex = 2;
            this.labelDescription.Text = "Description:";
            //
            // textBoxDescription
            //
            this.textBoxDescription.Location = new System.Drawing.Point(15, 76);
            this.textBoxDescription.Multiline = true;
            this.textBoxDescription.Name = "textBoxDescription";
            this.textBoxDescription.Size = new System.Drawing.Size(557, 40);
            this.textBoxDescription.TabIndex = 3;
            //
            // labelClassificationPrompt
            //
            this.labelClassificationPrompt.AutoSize = true;
            this.labelClassificationPrompt.Location = new System.Drawing.Point(12, 125);
            this.labelClassificationPrompt.Name = "labelClassificationPrompt";
            this.labelClassificationPrompt.Size = new System.Drawing.Size(319, 13);
            this.labelClassificationPrompt.TabIndex = 4;
            this.labelClassificationPrompt.Text = "Classification Prompt (Describe when an email belongs to this category):";
            //
            // textBoxClassificationPrompt
            //
            this.textBoxClassificationPrompt.Location = new System.Drawing.Point(15, 141);
            this.textBoxClassificationPrompt.Multiline = true;
            this.textBoxClassificationPrompt.Name = "textBoxClassificationPrompt";
            this.textBoxClassificationPrompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxClassificationPrompt.Size = new System.Drawing.Size(557, 100);
            this.textBoxClassificationPrompt.TabIndex = 5;
            //
            // checkBoxGenerateReply
            //
            this.checkBoxGenerateReply.AutoSize = true;
            this.checkBoxGenerateReply.Location = new System.Drawing.Point(15, 255);
            this.checkBoxGenerateReply.Name = "checkBoxGenerateReply";
            this.checkBoxGenerateReply.Size = new System.Drawing.Size(227, 17);
            this.checkBoxGenerateReply.TabIndex = 6;
            this.checkBoxGenerateReply.Text = "Automatically generate draft reply for this category";
            this.checkBoxGenerateReply.UseVisualStyleBackColor = true;
            this.checkBoxGenerateReply.CheckedChanged += new System.EventHandler(this.CheckBoxGenerateReply_CheckedChanged);
            //
            // labelReplyPrompt
            //
            this.labelReplyPrompt.AutoSize = true;
            this.labelReplyPrompt.Location = new System.Drawing.Point(12, 280);
            this.labelReplyPrompt.Name = "labelReplyPrompt";
            this.labelReplyPrompt.Size = new System.Drawing.Size(276, 13);
            this.labelReplyPrompt.TabIndex = 7;
            this.labelReplyPrompt.Text = "Reply Prompt (Instructions for generating the reply):";
            //
            // textBoxReplyPrompt
            //
            this.textBoxReplyPrompt.Enabled = false;
            this.textBoxReplyPrompt.Location = new System.Drawing.Point(15, 296);
            this.textBoxReplyPrompt.Multiline = true;
            this.textBoxReplyPrompt.Name = "textBoxReplyPrompt";
            this.textBoxReplyPrompt.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
            this.textBoxReplyPrompt.Size = new System.Drawing.Size(557, 100);
            this.textBoxReplyPrompt.TabIndex = 8;
            //
            // checkBoxIsEnabled
            //
            this.checkBoxIsEnabled.AutoSize = true;
            this.checkBoxIsEnabled.Checked = true;
            this.checkBoxIsEnabled.CheckState = System.Windows.Forms.CheckState.Checked;
            this.checkBoxIsEnabled.Location = new System.Drawing.Point(15, 410);
            this.checkBoxIsEnabled.Name = "checkBoxIsEnabled";
            this.checkBoxIsEnabled.Size = new System.Drawing.Size(137, 17);
            this.checkBoxIsEnabled.TabIndex = 9;
            this.checkBoxIsEnabled.Text = "Category Enabled";
            this.checkBoxIsEnabled.UseVisualStyleBackColor = true;
            //
            // buttonOK
            //
            this.buttonOK.Location = new System.Drawing.Point(416, 445);
            this.buttonOK.Name = "buttonOK";
            this.buttonOK.Size = new System.Drawing.Size(75, 23);
            this.buttonOK.TabIndex = 10;
            this.buttonOK.Text = "OK";
            this.buttonOK.UseVisualStyleBackColor = true;
            this.buttonOK.Click += new System.EventHandler(this.ButtonOK_Click);
            //
            // buttonCancel
            //
            this.buttonCancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.buttonCancel.Location = new System.Drawing.Point(497, 445);
            this.buttonCancel.Name = "buttonCancel";
            this.buttonCancel.Size = new System.Drawing.Size(75, 23);
            this.buttonCancel.TabIndex = 11;
            this.buttonCancel.Text = "Cancel";
            this.buttonCancel.UseVisualStyleBackColor = true;
            //
            // CategoryEditorForm
            //
            this.AcceptButton = this.buttonOK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.CancelButton = this.buttonCancel;
            this.ClientSize = new System.Drawing.Size(584, 481);
            this.Controls.Add(this.buttonCancel);
            this.Controls.Add(this.buttonOK);
            this.Controls.Add(this.checkBoxIsEnabled);
            this.Controls.Add(this.textBoxReplyPrompt);
            this.Controls.Add(this.labelReplyPrompt);
            this.Controls.Add(this.checkBoxGenerateReply);
            this.Controls.Add(this.textBoxClassificationPrompt);
            this.Controls.Add(this.labelClassificationPrompt);
            this.Controls.Add(this.textBoxDescription);
            this.Controls.Add(this.labelDescription);
            this.Controls.Add(this.textBoxCategoryName);
            this.Controls.Add(this.labelCategoryName);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "CategoryEditorForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterParent;
            this.Text = "Email Category Editor";
            this.ResumeLayout(false);
            this.PerformLayout();
        }

        private System.Windows.Forms.Label labelCategoryName;
        private System.Windows.Forms.TextBox textBoxCategoryName;
        private System.Windows.Forms.Label labelDescription;
        private System.Windows.Forms.TextBox textBoxDescription;
        private System.Windows.Forms.Label labelClassificationPrompt;
        private System.Windows.Forms.TextBox textBoxClassificationPrompt;
        private System.Windows.Forms.CheckBox checkBoxGenerateReply;
        private System.Windows.Forms.Label labelReplyPrompt;
        private System.Windows.Forms.TextBox textBoxReplyPrompt;
        private System.Windows.Forms.CheckBox checkBoxIsEnabled;
        private System.Windows.Forms.Button buttonOK;
        private System.Windows.Forms.Button buttonCancel;

        private void LoadCategoryData()
        {
            if (_category != null)
            {
                textBoxCategoryName.Text = _category.CategoryName;
                textBoxDescription.Text = _category.Description;
                textBoxClassificationPrompt.Text = _category.ClassificationPrompt;
                checkBoxGenerateReply.Checked = _category.GenerateReplyDraft;
                textBoxReplyPrompt.Text = _category.ReplyPrompt;
                textBoxReplyPrompt.Enabled = _category.GenerateReplyDraft;
                checkBoxIsEnabled.Checked = _category.IsEnabled;
            }
        }

        private void CheckBoxGenerateReply_CheckedChanged(object sender, EventArgs e)
        {
            textBoxReplyPrompt.Enabled = checkBoxGenerateReply.Checked;
            labelReplyPrompt.Enabled = checkBoxGenerateReply.Checked;
        }

        private void ButtonOK_Click(object sender, EventArgs e)
        {
            // Validate input
            if (string.IsNullOrWhiteSpace(textBoxCategoryName.Text))
            {
                MessageBox.Show("Please enter a category name.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBoxCategoryName.Focus();
                return;
            }

            if (string.IsNullOrWhiteSpace(textBoxClassificationPrompt.Text))
            {
                MessageBox.Show("Please enter a classification prompt.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBoxClassificationPrompt.Focus();
                return;
            }

            if (checkBoxGenerateReply.Checked && string.IsNullOrWhiteSpace(textBoxReplyPrompt.Text))
            {
                MessageBox.Show("Please enter a reply prompt or disable auto-reply generation.", "Validation Error",
                    MessageBoxButtons.OK, MessageBoxIcon.Warning);
                textBoxReplyPrompt.Focus();
                return;
            }

            // Save data
            _category.CategoryName = textBoxCategoryName.Text.Trim();
            _category.Description = textBoxDescription.Text.Trim();
            _category.ClassificationPrompt = textBoxClassificationPrompt.Text.Trim();
            _category.GenerateReplyDraft = checkBoxGenerateReply.Checked;
            _category.ReplyPrompt = textBoxReplyPrompt.Text.Trim();
            _category.IsEnabled = checkBoxIsEnabled.Checked;

            this.DialogResult = DialogResult.OK;
            this.Close();
        }
    }
}
