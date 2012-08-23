using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

namespace TspdCfg.SalesDemo.DynTmplts
{
	/// <summary>
	/// Summary description for testForm.
	/// </summary>
	public class testForm : System.Windows.Forms.Form
	{
		private System.Windows.Forms.ListView listView1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public testForm()
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
			this.listView1 = new System.Windows.Forms.ListView();
			this.SuspendLayout();
			// 
			// listView1
			// 
			this.listView1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(177)));
			this.listView1.ForeColor = System.Drawing.SystemColors.HotTrack;
			this.listView1.Location = new System.Drawing.Point(112, 184);
			this.listView1.Name = "listView1";
			this.listView1.Size = new System.Drawing.Size(296, 368);
			this.listView1.TabIndex = 0;
			// 
			// testForm
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(904, 692);
			this.Controls.Add(this.listView1);
			this.Name = "testForm";
			this.Text = "testForm";
			this.ResumeLayout(false);

		}
		#endregion
	}
}
