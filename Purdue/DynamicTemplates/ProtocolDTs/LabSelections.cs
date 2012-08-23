using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for LabSelections.
	/// </summary>
	public class LabSelections : System.Windows.Forms.Form
	{
		public bool ShowVariableAbbreviation;
		public bool IncludeScheduledTimes;

		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.CheckBox chkVariableAbbreviation;
		private System.Windows.Forms.CheckBox chkIncludeTimes;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public LabSelections()
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
			this.button1 = new System.Windows.Forms.Button();
			this.chkVariableAbbreviation = new System.Windows.Forms.CheckBox();
			this.chkIncludeTimes = new System.Windows.Forms.CheckBox();
			this.SuspendLayout();
			// 
			// button1
			// 
			this.button1.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.button1.Location = new System.Drawing.Point(132, 104);
			this.button1.Name = "button1";
			this.button1.TabIndex = 0;
			this.button1.Text = "OK";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// chkVariableAbbreviation
			// 
			this.chkVariableAbbreviation.Location = new System.Drawing.Point(48, 24);
			this.chkVariableAbbreviation.Name = "chkVariableAbbreviation";
			this.chkVariableAbbreviation.Size = new System.Drawing.Size(256, 24);
			this.chkVariableAbbreviation.TabIndex = 1;
			this.chkVariableAbbreviation.Text = "Show variable abbreviation";
			// 
			// chkIncludeTimes
			// 
			this.chkIncludeTimes.Location = new System.Drawing.Point(48, 56);
			this.chkIncludeTimes.Name = "chkIncludeTimes";
			this.chkIncludeTimes.Size = new System.Drawing.Size(256, 24);
			this.chkIncludeTimes.TabIndex = 2;
			this.chkIncludeTimes.Text = "Include scheduled times";
			// 
			// LabSelections
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(338, 136);
			this.Controls.Add(this.chkIncludeTimes);
			this.Controls.Add(this.chkVariableAbbreviation);
			this.Controls.Add(this.button1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "LabSelections";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Lab Assessment Selections";
			this.TopMost = true;
			this.ResumeLayout(false);

		}
		#endregion

		private void button1_Click(object sender, System.EventArgs e)
		{
			ShowVariableAbbreviation = chkVariableAbbreviation.Checked;
			IncludeScheduledTimes = chkIncludeTimes.Checked;

			DialogResult = DialogResult.OK;
			this.Close();
		}
	}
}
