using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using Tspd.Businessobject;

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for DurationSelect.
	/// </summary>
	public class DurationSelect : System.Windows.Forms.Form
	{
		public ArrayList EnumPairs = null;
		public int SelectedDuration = -1;

		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.Label lbl1;
		private ComboBox comboBox1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public DurationSelect()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
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
			this.lbl1 = new System.Windows.Forms.Label();
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.SuspendLayout();
			// 
			// btnOK
			// 
			this.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.btnOK.Location = new System.Drawing.Point(152, 88);
			this.btnOK.Name = "btnOK";
			this.btnOK.TabIndex = 0;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// lbl1
			// 
			this.lbl1.Location = new System.Drawing.Point(16, 16);
			this.lbl1.Name = "lbl1";
			this.lbl1.Size = new System.Drawing.Size(120, 16);
			this.lbl1.TabIndex = 1;
			this.lbl1.Text = "Select a duration:";
			// 
			// comboBox1
			// 
			this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBox1.Location = new System.Drawing.Point(16, 40);
			this.comboBox1.Name = "comboBox1";
			this.comboBox1.Size = new System.Drawing.Size(360, 21);
			this.comboBox1.TabIndex = 2;
			// 
			// DurationSelect
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(386, 120);
			this.Controls.Add(this.comboBox1);
			this.Controls.Add(this.lbl1);
			this.Controls.Add(this.btnOK);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "DurationSelect";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Duration Selection";
			this.TopMost = true;
			this.Load += new System.EventHandler(this.DurationSelect_Load);
			this.ResumeLayout(false);

		}
		#endregion

		public void loadDurations(ArrayList durations) 
		{
			comboBox1.Items.Clear();

			foreach (EnumPair ep in durations) 
			{
				comboBox1.Items.Add(ep.getUserLabel());
			}

			comboBox1.SelectedIndex = -1;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			if (comboBox1.SelectedIndex == -1) 
			{
				MessageBox.Show("Select a Duration", "Invalid Selection");
				return;
			}

			// Set return values
			SelectedDuration = comboBox1.SelectedIndex;

			DialogResult = DialogResult.OK;
			this.Close();
		}

		private void DurationSelect_Load(object sender, System.EventArgs e)
		{
		
		}
	}
}
