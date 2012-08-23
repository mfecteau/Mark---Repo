using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

using TspdCfg.FastTrack.DynTmplts;
using TspdCfg.SalesDemo.DynTmplts;

namespace Tests
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.TextBox txtStart;
		private System.Windows.Forms.TextBox txtEnd;
		private System.Windows.Forms.TextBox txtUnit;
		private System.Windows.Forms.TextBox txtErr;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.TextBox txtInput;
		private System.Windows.Forms.TextBox txtFailResults;
		private System.Windows.Forms.TextBox txtSuccessResults;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Label label2;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
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
				if (components != null) 
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
			this.txtInput = new System.Windows.Forms.TextBox();
			this.button1 = new System.Windows.Forms.Button();
			this.txtErr = new System.Windows.Forms.TextBox();
			this.txtStart = new System.Windows.Forms.TextBox();
			this.txtEnd = new System.Windows.Forms.TextBox();
			this.txtUnit = new System.Windows.Forms.TextBox();
			this.txtFailResults = new System.Windows.Forms.TextBox();
			this.button2 = new System.Windows.Forms.Button();
			this.txtSuccessResults = new System.Windows.Forms.TextBox();
			this.label1 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// txtInput
			// 
			this.txtInput.Location = new System.Drawing.Point(8, 8);
			this.txtInput.Name = "txtInput";
			this.txtInput.Size = new System.Drawing.Size(408, 20);
			this.txtInput.TabIndex = 0;
			this.txtInput.Text = "1.5 h";
			// 
			// button1
			// 
			this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button1.Location = new System.Drawing.Point(440, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(67, 23);
			this.button1.TabIndex = 1;
			this.button1.Text = "Test";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// txtErr
			// 
			this.txtErr.Location = new System.Drawing.Point(8, 64);
			this.txtErr.Name = "txtErr";
			this.txtErr.ReadOnly = true;
			this.txtErr.Size = new System.Drawing.Size(408, 20);
			this.txtErr.TabIndex = 7;
			this.txtErr.Text = "";
			// 
			// txtStart
			// 
			this.txtStart.Location = new System.Drawing.Point(40, 40);
			this.txtStart.Name = "txtStart";
			this.txtStart.ReadOnly = true;
			this.txtStart.Size = new System.Drawing.Size(72, 20);
			this.txtStart.TabIndex = 3;
			this.txtStart.Text = "";
			// 
			// txtEnd
			// 
			this.txtEnd.Location = new System.Drawing.Point(168, 40);
			this.txtEnd.Name = "txtEnd";
			this.txtEnd.ReadOnly = true;
			this.txtEnd.Size = new System.Drawing.Size(80, 20);
			this.txtEnd.TabIndex = 5;
			this.txtEnd.Text = "";
			// 
			// txtUnit
			// 
			this.txtUnit.Location = new System.Drawing.Point(264, 40);
			this.txtUnit.Name = "txtUnit";
			this.txtUnit.ReadOnly = true;
			this.txtUnit.Size = new System.Drawing.Size(80, 20);
			this.txtUnit.TabIndex = 6;
			this.txtUnit.Text = "";
			// 
			// txtFailResults
			// 
			this.txtFailResults.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
				| System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtFailResults.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtFailResults.Location = new System.Drawing.Point(8, 104);
			this.txtFailResults.Multiline = true;
			this.txtFailResults.Name = "txtFailResults";
			this.txtFailResults.ReadOnly = true;
			this.txtFailResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtFailResults.Size = new System.Drawing.Size(424, 208);
			this.txtFailResults.TabIndex = 8;
			this.txtFailResults.Text = "";
			// 
			// button2
			// 
			this.button2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
			this.button2.Location = new System.Drawing.Point(440, 104);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(67, 23);
			this.button2.TabIndex = 9;
			this.button2.Text = "Test";
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// txtSuccessResults
			// 
			this.txtSuccessResults.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.txtSuccessResults.Font = new System.Drawing.Font("Lucida Console", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.txtSuccessResults.Location = new System.Drawing.Point(8, 320);
			this.txtSuccessResults.Multiline = true;
			this.txtSuccessResults.Name = "txtSuccessResults";
			this.txtSuccessResults.ReadOnly = true;
			this.txtSuccessResults.ScrollBars = System.Windows.Forms.ScrollBars.Vertical;
			this.txtSuccessResults.Size = new System.Drawing.Size(424, 152);
			this.txtSuccessResults.TabIndex = 10;
			this.txtSuccessResults.Text = "";
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 40);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(32, 23);
			this.label1.TabIndex = 2;
			this.label1.Text = "start";
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(128, 40);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(32, 23);
			this.label2.TabIndex = 4;
			this.label2.Text = "end";
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(520, 488);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.txtSuccessResults);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.txtFailResults);
			this.Controls.Add(this.txtUnit);
			this.Controls.Add(this.txtEnd);
			this.Controls.Add(this.txtStart);
			this.Controls.Add(this.txtErr);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.txtInput);
			this.Name = "Form1";
			this.Text = "Form1";
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			string startTime;
			string endTime;
			string unit;
			string serr;

			txtSuccessResults.Text = "";
			txtFailResults.Text = "";

			bool success = PfizerUtil.parseTimePoint(txtInput.Text, out startTime, out endTime, out unit, out serr);

			if (success) 
			{
				txtErr.Text = "";
				txtStart.Text = startTime;
				txtEnd.Text = endTime;
				txtUnit.Text = unit;

			
				PfizerUtil.TimeUnit tu1 = PfizerUtil.TimeUnit.find(unit);

				double smin = double.Parse(startTime) * tu1.getMultiplier();
				string stxt = "Start in min: " + smin + "\r\n";

				if (endTime.Length != 0) 
				{
					double emin = double.Parse(endTime) * tu1.getMultiplier();
					stxt += "End in min: " + emin + "\r\n";
				}

				txtFailResults.Text = stxt;
			}
			else
			{
				txtErr.Text = serr;
				txtStart.Text = "";
				txtEnd.Text = "";
				txtUnit.Text = "";
			}
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			string startTime;
			string endTime;
			string unit;
			string serr;
			string stxt = "";

			ArrayList tests = new ArrayList();

			tests.Add("0.5 to 1.5 h");
			tests.Add("1.5 to 0.5 h");

			tests.Add("-5 to -6 min");
			tests.Add("-6 to -6 min");
			tests.Add("-7 to -6 min");

			tests.Add("-7 to -6 p");
			tests.Add("-7 to -6 mn");
			tests.Add("-7 to -6 minnits");
			tests.Add("-7 to -6 m");
			tests.Add("-7 to -6 min");
			tests.Add("-7 to -6 minutes");

			tests.Add("-7 to -6 h");
			tests.Add("-7 to -6 hour");
			tests.Add("-7 to -6 hours");
			tests.Add("-7 to -6 hrs");
			tests.Add("-7 to -6 his");
			tests.Add("-7 to -6 hers");
			tests.Add("-7 to -6 plp");
			

			tests.Add("-5to-6 min");
			tests.Add("-6to-6 min");
			tests.Add("-7to-6 min");

			tests.Add("-5 - -6 min");
			tests.Add("-6 - -6 min");
			tests.Add("-7 - -6 min");

			tests.Add("-5- -6 min");
			tests.Add("-6- -6 min");
			tests.Add("-7- -6 min");

			tests.Add("-5 --6 min");
			tests.Add("-6 --6 min");
			tests.Add("-7 --6 min");

			tests.Add("-5--6 min");
			tests.Add("-6--6 min");
			tests.Add("-7--6 min");

			tests.Add("5 to 6 min");
			tests.Add("6 to 6 min");
			tests.Add("7 to 6 min");

			tests.Add("5to6 min");
			tests.Add("6to6 min");
			tests.Add("7to6 min");

			tests.Add("5 - 6 min");
			tests.Add("6 - 6 min");
			tests.Add("7 - 6 min");

			tests.Add("5- 6 min");
			tests.Add("6- 6 min");
			tests.Add("7- 6 min");

			tests.Add("5 -6 min");
			tests.Add("6 -6 min");
			tests.Add("7 -6 min");

			tests.Add("5-6 min");
			tests.Add("6-6 min");
			tests.Add("7-6 min");

			tests.Add("5 6 min");
			tests.Add("-7 6 min");

			txtSuccessResults.Text = "";
			txtFailResults.Text = "";

			foreach (string test1 in tests) 
			{
				bool success = PfizerUtil.parseTimePoint(test1, out startTime, out endTime, out unit, out serr);
				if (success) 
				{
					stxt = "Success\t\"" + test1 + "\"; start: " + startTime;

					if (endTime.Length != 0) 
					{
						stxt += ", end: " + endTime;
					}

					stxt += " " + unit;

					txtSuccessResults.Text += stxt + "\r\n";
				}
				else
				{
					txtFailResults.Text += "Failure\t\"" + test1 + "\": " + serr + "\r\n";
				}
			}
		}
	}
}
