namespace TspdCfg.FastTrack.PlugIn
{
    partial class frmTVMapper
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmTVMapper));
            this.label1 = new System.Windows.Forms.Label();
            this.cmbSOA = new System.Windows.Forms.ComboBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.rtf = new System.Windows.Forms.RichTextBox();
            this.lstTask = new System.Windows.Forms.ListView();
            this.columnHeader1 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.restoretext = new System.Windows.Forms.LinkLabel();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.btnAppend = new System.Windows.Forms.Button();
            this.rtfPreview = new System.Windows.Forms.RichTextBox();
            this.lstTvDesc = new System.Windows.Forms.ListView();
            this.columnHeader2 = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.btnClose = new System.Windows.Forms.Button();
            this.btnSave = new System.Windows.Forms.Button();
            this.groupBox1.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(6, 8);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select schedule:";
            // 
            // cmbSOA
            // 
            this.cmbSOA.FormattingEnabled = true;
            this.cmbSOA.Location = new System.Drawing.Point(98, 5);
            this.cmbSOA.Name = "cmbSOA";
            this.cmbSOA.Size = new System.Drawing.Size(423, 21);
            this.cmbSOA.TabIndex = 1;
            this.cmbSOA.SelectedIndexChanged += new System.EventHandler(this.cmbSOA_SelectedIndexChanged);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label6);
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.rtf);
            this.groupBox1.Controls.Add(this.lstTask);
            this.groupBox1.Controls.Add(this.restoretext);
            this.groupBox1.Location = new System.Drawing.Point(9, 28);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(723, 188);
            this.groupBox1.TabIndex = 2;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Task-Events to be Mapped";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label6.Location = new System.Drawing.Point(281, 17);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(74, 13);
            this.label6.TabIndex = 4;
            this.label6.Text = "Current details";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label5.Location = new System.Drawing.Point(6, 19);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(70, 13);
            this.label5.TabIndex = 3;
            this.label5.Text = "Task-Events:";
            this.label5.Visible = false;
            // 
            // rtf
            // 
            this.rtf.Location = new System.Drawing.Point(285, 33);
            this.rtf.Name = "rtf";
            this.rtf.Size = new System.Drawing.Size(431, 149);
            this.rtf.TabIndex = 1;
            this.rtf.Text = "";
            // 
            // lstTask
            // 
            this.lstTask.Alignment = System.Windows.Forms.ListViewAlignment.Left;
            this.lstTask.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader1});
            this.lstTask.FullRowSelect = true;
            this.lstTask.HideSelection = false;
            this.lstTask.Location = new System.Drawing.Point(6, 32);
            this.lstTask.MultiSelect = false;
            this.lstTask.Name = "lstTask";
            this.lstTask.Size = new System.Drawing.Size(271, 150);
            this.lstTask.Sorting = System.Windows.Forms.SortOrder.Ascending;
            this.lstTask.TabIndex = 0;
            this.lstTask.UseCompatibleStateImageBehavior = false;
            this.lstTask.View = System.Windows.Forms.View.Details;
            this.lstTask.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.lstTask_ItemSelectionChanged);
            // 
            // columnHeader1
            // 
            this.columnHeader1.Text = "Task-Events:";
            this.columnHeader1.Width = 265;
            // 
            // restoretext
            // 
            this.restoretext.AutoSize = true;
            this.restoretext.Location = new System.Drawing.Point(597, 17);
            this.restoretext.Name = "restoretext";
            this.restoretext.Size = new System.Drawing.Size(102, 13);
            this.restoretext.TabIndex = 2;
            this.restoretext.TabStop = true;
            this.restoretext.Text = "Restore Original text";
            this.restoretext.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.restoretext_LinkClicked);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label4);
            this.groupBox2.Controls.Add(this.label3);
            this.groupBox2.Controls.Add(this.btnAppend);
            this.groupBox2.Controls.Add(this.rtfPreview);
            this.groupBox2.Controls.Add(this.lstTvDesc);
            this.groupBox2.Location = new System.Drawing.Point(9, 222);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(723, 205);
            this.groupBox2.TabIndex = 3;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Standard Text Options";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label4.Location = new System.Drawing.Point(6, 14);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(0, 13);
            this.label4.TabIndex = 6;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label3.Location = new System.Drawing.Point(282, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(45, 13);
            this.label3.TabIndex = 5;
            this.label3.Text = "Preview";
            // 
            // btnAppend
            // 
            this.btnAppend.Location = new System.Drawing.Point(134, 177);
            this.btnAppend.Name = "btnAppend";
            this.btnAppend.Size = new System.Drawing.Size(143, 27);
            this.btnAppend.TabIndex = 4;
            this.btnAppend.Text = "Append to Current Details";
            this.btnAppend.UseVisualStyleBackColor = true;
            this.btnAppend.Click += new System.EventHandler(this.btnAppend_Click);
            // 
            // rtfPreview
            // 
            this.rtfPreview.Location = new System.Drawing.Point(285, 29);
            this.rtfPreview.Name = "rtfPreview";
            this.rtfPreview.ReadOnly = true;
            this.rtfPreview.Size = new System.Drawing.Size(432, 144);
            this.rtfPreview.TabIndex = 3;
            this.rtfPreview.Text = "";
            // 
            // lstTvDesc
            // 
            this.lstTvDesc.Alignment = System.Windows.Forms.ListViewAlignment.Left;
            this.lstTvDesc.CheckBoxes = true;
            this.lstTvDesc.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.columnHeader2});
            this.lstTvDesc.FullRowSelect = true;
            this.lstTvDesc.Location = new System.Drawing.Point(6, 30);
            this.lstTvDesc.Name = "lstTvDesc";
            this.lstTvDesc.Size = new System.Drawing.Size(271, 145);
            this.lstTvDesc.TabIndex = 2;
            this.lstTvDesc.UseCompatibleStateImageBehavior = false;
            this.lstTvDesc.View = System.Windows.Forms.View.Details;
            this.lstTvDesc.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.lstTvDesc_ItemChecked);
            this.lstTvDesc.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.lstTvDesc_ItemSelectionChanged);
            // 
            // columnHeader2
            // 
            this.columnHeader2.Text = "";
            this.columnHeader2.Width = 264;
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(660, 432);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(72, 26);
            this.btnClose.TabIndex = 4;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(582, 432);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(72, 26);
            this.btnSave.TabIndex = 5;
            this.btnSave.Text = "&Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // frmTVMapper
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(737, 461);
            this.Controls.Add(this.cmbSOA);
            this.Controls.Add(this.btnSave);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmTVMapper";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Task-Event Mapping";
            this.Load += new System.EventHandler(this.frmTVMapper_Load);
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ComboBox cmbSOA;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.RichTextBox rtf;
        private System.Windows.Forms.ListView lstTask;
        private System.Windows.Forms.LinkLabel restoretext;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button btnAppend;
        private System.Windows.Forms.RichTextBox rtfPreview;
        private System.Windows.Forms.ListView lstTvDesc;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ColumnHeader columnHeader1;
        private System.Windows.Forms.ColumnHeader columnHeader2;
    }
}