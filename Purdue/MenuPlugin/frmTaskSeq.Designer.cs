namespace TspdCfg.FastTrack.PlugIn
{
    partial class frmTaskSeq
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
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmTaskSeq));
            this.cmbSOA = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.btnSave = new System.Windows.Forms.Button();
            this.btnClose = new System.Windows.Forms.Button();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.btnReset = new System.Windows.Forms.Button();
            this.btnDown = new System.Windows.Forms.Button();
            this.btnUp = new System.Windows.Forms.Button();
            this.lstTask = new System.Windows.Forms.ListView();
            this.colIdx = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.colTask = ((System.Windows.Forms.ColumnHeader)(new System.Windows.Forms.ColumnHeader()));
            this.cmbVisit = new System.Windows.Forms.ComboBox();
            this.label2 = new System.Windows.Forms.Label();
            this.imageList1 = new System.Windows.Forms.ImageList(this.components);
            this.groupBox2.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.SuspendLayout();
            // 
            // cmbSOA
            // 
            this.cmbSOA.FormattingEnabled = true;
            this.cmbSOA.Location = new System.Drawing.Point(104, 10);
            this.cmbSOA.Name = "cmbSOA";
            this.cmbSOA.Size = new System.Drawing.Size(375, 21);
            this.cmbSOA.TabIndex = 3;
            this.cmbSOA.SelectedIndexChanged += new System.EventHandler(this.cmbSOA_SelectedIndexChanged_1);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(86, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select schedule:";
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.btnSave);
            this.groupBox2.Controls.Add(this.btnClose);
            this.groupBox2.Controls.Add(this.groupBox3);
            this.groupBox2.Controls.Add(this.cmbVisit);
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.groupBox2.Location = new System.Drawing.Point(0, 30);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(531, 399);
            this.groupBox2.TabIndex = 5;
            this.groupBox2.TabStop = false;
            // 
            // btnSave
            // 
            this.btnSave.Location = new System.Drawing.Point(389, 368);
            this.btnSave.Name = "btnSave";
            this.btnSave.Size = new System.Drawing.Size(66, 23);
            this.btnSave.TabIndex = 8;
            this.btnSave.Text = "&Save";
            this.btnSave.UseVisualStyleBackColor = true;
            this.btnSave.Click += new System.EventHandler(this.btnSave_Click);
            // 
            // btnClose
            // 
            this.btnClose.Location = new System.Drawing.Point(461, 369);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(66, 23);
            this.btnClose.TabIndex = 7;
            this.btnClose.Text = "&Close";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.btnReset);
            this.groupBox3.Controls.Add(this.btnDown);
            this.groupBox3.Controls.Add(this.btnUp);
            this.groupBox3.Controls.Add(this.lstTask);
            this.groupBox3.Location = new System.Drawing.Point(18, 41);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(507, 321);
            this.groupBox3.TabIndex = 6;
            this.groupBox3.TabStop = false;
            // 
            // btnReset
            // 
            this.btnReset.Location = new System.Drawing.Point(352, 13);
            this.btnReset.Name = "btnReset";
            this.btnReset.Size = new System.Drawing.Size(109, 22);
            this.btnReset.TabIndex = 8;
            this.btnReset.Text = "&Reset Numbering";
            this.btnReset.UseVisualStyleBackColor = true;
            this.btnReset.Click += new System.EventHandler(this.btnReset_Click);
            // 
            // btnDown
            // 
            this.btnDown.ImageIndex = 1;
            this.btnDown.ImageList = this.imageList1;
            this.btnDown.Location = new System.Drawing.Point(468, 171);
            this.btnDown.Name = "btnDown";
            this.btnDown.Size = new System.Drawing.Size(32, 32);
            this.btnDown.TabIndex = 7;
            this.btnDown.UseVisualStyleBackColor = true;
            this.btnDown.Click += new System.EventHandler(this.btnDown_Click);
            // 
            // btnUp
            // 
            this.btnUp.ImageIndex = 0;
            this.btnUp.ImageList = this.imageList1;
            this.btnUp.Location = new System.Drawing.Point(468, 139);
            this.btnUp.Name = "btnUp";
            this.btnUp.Size = new System.Drawing.Size(32, 32);
            this.btnUp.TabIndex = 6;
            this.btnUp.UseVisualStyleBackColor = true;
            this.btnUp.Click += new System.EventHandler(this.btnUp_Click);
            // 
            // lstTask
            // 
            this.lstTask.AllowColumnReorder = true;
            this.lstTask.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
            this.colIdx,
            this.colTask});
            this.lstTask.FullRowSelect = true;
            this.lstTask.GridLines = true;
            this.lstTask.Location = new System.Drawing.Point(6, 39);
            this.lstTask.MultiSelect = false;
            this.lstTask.Name = "lstTask";
            this.lstTask.Size = new System.Drawing.Size(455, 277);
            this.lstTask.TabIndex = 5;
            this.lstTask.UseCompatibleStateImageBehavior = false;
            this.lstTask.View = System.Windows.Forms.View.Details;
            // 
            // colIdx
            // 
            this.colIdx.Text = "Seq";
            this.colIdx.Width = 44;
            // 
            // colTask
            // 
            this.colTask.Text = "Task";
            this.colTask.Width = 399;
            // 
            // cmbVisit
            // 
            this.cmbVisit.FormattingEnabled = true;
            this.cmbVisit.Location = new System.Drawing.Point(104, 13);
            this.cmbVisit.Name = "cmbVisit";
            this.cmbVisit.Size = new System.Drawing.Size(375, 21);
            this.cmbVisit.TabIndex = 4;
            this.cmbVisit.SelectedIndexChanged += new System.EventHandler(this.cmbVisit_SelectedIndexChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(15, 17);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(62, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Select Visit:";
            // 
            // imageList1
            // 
            this.imageList1.ImageStream = ((System.Windows.Forms.ImageListStreamer)(resources.GetObject("imageList1.ImageStream")));
            this.imageList1.TransparentColor = System.Drawing.Color.Transparent;
            this.imageList1.Images.SetKeyName(0, "up_arrow.png");
            this.imageList1.Images.SetKeyName(1, "down_arrow.png");
            this.imageList1.Images.SetKeyName(2, "Down_Arrow_Icon.png");
            // 
            // frmTaskSeq
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(531, 429);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.cmbSOA);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "frmTaskSeq";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Task Sequencing";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.frmTaskSeq_Load);
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.groupBox3.ResumeLayout(false);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.ComboBox cmbSOA;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.ComboBox cmbVisit;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ListView lstTask;
        private System.Windows.Forms.Button btnSave;
        private System.Windows.Forms.Button btnClose;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Button btnDown;
        private System.Windows.Forms.Button btnUp;
        private System.Windows.Forms.ColumnHeader colIdx;
        private System.Windows.Forms.ColumnHeader colTask;
        private System.Windows.Forms.Button btnReset;
        private System.Windows.Forms.ImageList imageList1;
    }
}