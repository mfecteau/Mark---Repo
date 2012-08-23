using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using Tspd.Businessobject;

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for TaskSelect.
	/// </summary>
	public class TaskSelect : System.Windows.Forms.Form
	{
		public int SelectedTask = -1;
		public bool AddHeader = false;
		public bool UseInSynopsis = false;
		public bool DisplayWindow = false;
		public bool AddStudyVariables = false;

		private System.Windows.Forms.ComboBox cboTasks;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.CheckBox chkAddHeader;
		private System.Windows.Forms.CheckBox chkUseInSynopsis;
		private System.Windows.Forms.CheckBox chkDisplayWindow;
		private System.Windows.Forms.CheckBox chkAddStudyVariables;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public TaskSelect()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if(components != null)
				{
					components.Dispose();
				}
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code
		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            this.btnOK = new System.Windows.Forms.Button();
            this.cboTasks = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.chkAddHeader = new System.Windows.Forms.CheckBox();
            this.chkUseInSynopsis = new System.Windows.Forms.CheckBox();
            this.chkDisplayWindow = new System.Windows.Forms.CheckBox();
            this.chkAddStudyVariables = new System.Windows.Forms.CheckBox();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnOK.Location = new System.Drawing.Point(159, 108);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 6;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // cboTasks
            // 
            this.cboTasks.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cboTasks.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboTasks.Location = new System.Drawing.Point(8, 34);
            this.cboTasks.Name = "cboTasks";
            this.cboTasks.Size = new System.Drawing.Size(358, 21);
            this.cboTasks.TabIndex = 1;
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.Location = new System.Drawing.Point(5, 11);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(84, 20);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select a Task";
            // 
            // chkAddHeader
            // 
            this.chkAddHeader.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.chkAddHeader.Checked = true;
            this.chkAddHeader.CheckState = System.Windows.Forms.CheckState.Checked;
            this.chkAddHeader.Location = new System.Drawing.Point(203, 87);
            this.chkAddHeader.Name = "chkAddHeader";
            this.chkAddHeader.Size = new System.Drawing.Size(46, 24);
            this.chkAddHeader.TabIndex = 2;
            this.chkAddHeader.Text = "Add Task Header";
            this.chkAddHeader.Visible = false;
            // 
            // chkUseInSynopsis
            // 
            this.chkUseInSynopsis.Location = new System.Drawing.Point(182, 67);
            this.chkUseInSynopsis.Name = "chkUseInSynopsis";
            this.chkUseInSynopsis.Size = new System.Drawing.Size(136, 24);
            this.chkUseInSynopsis.TabIndex = 4;
            this.chkUseInSynopsis.Text = "Use in Synopsis";
            this.chkUseInSynopsis.Visible = false;
            // 
            // chkDisplayWindow
            // 
            this.chkDisplayWindow.Location = new System.Drawing.Point(8, 60);
            this.chkDisplayWindow.Name = "chkDisplayWindow";
            this.chkDisplayWindow.Size = new System.Drawing.Size(128, 24);
            this.chkDisplayWindow.TabIndex = 3;
            this.chkDisplayWindow.Text = "Display Window";
            // 
            // chkAddStudyVariables
            // 
            this.chkAddStudyVariables.Location = new System.Drawing.Point(8, 82);
            this.chkAddStudyVariables.Name = "chkAddStudyVariables";
            this.chkAddStudyVariables.Size = new System.Drawing.Size(168, 24);
            this.chkAddStudyVariables.TabIndex = 5;
            this.chkAddStudyVariables.Text = "List Study Variables";
            // 
            // TaskSelect
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(376, 136);
            this.Controls.Add(this.chkAddStudyVariables);
            this.Controls.Add(this.chkDisplayWindow);
            this.Controls.Add(this.chkUseInSynopsis);
            this.Controls.Add(this.chkAddHeader);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboTasks);
            this.Controls.Add(this.btnOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "TaskSelect";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Task Selection";
            this.TopMost = true;
            this.ResumeLayout(false);

		}
		#endregion

		public void loadTasks(ArrayList taskList) 
		{
			cboTasks.Items.Clear();

			foreach (Task t in taskList) 
			{
				cboTasks.Items.Add(t.getBriefDescription());
			}

			cboTasks.SelectedIndex = -1;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			if (cboTasks.SelectedIndex == -1) 
			{
				MessageBox.Show("Select a Task", "Invalid Selection");
				return;
			}

			// Set return values
			SelectedTask = cboTasks.SelectedIndex;
			AddHeader = chkAddHeader.Checked;
			UseInSynopsis = chkUseInSynopsis.Checked;
			DisplayWindow = chkDisplayWindow.Checked;
			AddStudyVariables = chkAddStudyVariables.Checked;

			DialogResult = DialogResult.OK;
			this.Close();
		}
	}
}
