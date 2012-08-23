using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Tspd.Tspddoc;
using Tspd.Businessobject;
using Tspd.Utilities;

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for LibraryItem.
	/// </summary>
	public class LibraryItem : System.Windows.Forms.Form
	{
		public object SelectedItem;

		ArrayList buckets = new ArrayList();

		private System.Windows.Forms.ComboBox comboBox1;
		private System.Windows.Forms.Label lbl1;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.ComboBox comboBox2;
		private System.Windows.Forms.Label label1;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public LibraryItem()
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
            this.comboBox1 = new System.Windows.Forms.ComboBox();
            this.lbl1 = new System.Windows.Forms.Label();
            this.btnOK = new System.Windows.Forms.Button();
            this.comboBox2 = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // comboBox1
            // 
            this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox1.Location = new System.Drawing.Point(12, 33);
            this.comboBox1.Name = "comboBox1";
            this.comboBox1.Size = new System.Drawing.Size(286, 21);
            this.comboBox1.TabIndex = 5;
            // 
            // lbl1
            // 
            this.lbl1.Location = new System.Drawing.Point(12, 9);
            this.lbl1.Name = "lbl1";
            this.lbl1.Size = new System.Drawing.Size(120, 16);
            this.lbl1.TabIndex = 4;
            this.lbl1.Text = "Select a Library Item:";
            // 
            // btnOK
            // 
            this.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.btnOK.Location = new System.Drawing.Point(118, 59);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 3;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // comboBox2
            // 
            this.comboBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.comboBox2.Location = new System.Drawing.Point(75, 70);
            this.comboBox2.Name = "comboBox2";
            this.comboBox2.Size = new System.Drawing.Size(286, 21);
            this.comboBox2.TabIndex = 7;
            this.comboBox2.Visible = false;
            this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(-43, 70);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(152, 16);
            this.label1.TabIndex = 6;
            this.label1.Text = "Select a Library Category:";
            this.label1.Visible = false;
            // 
            // LibraryItem
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(318, 86);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.comboBox1);
            this.Controls.Add(this.lbl1);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.comboBox2);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "LibraryItem";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Library Item";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.LibraryItem_Load);
            this.ResumeLayout(false);

		}
		#endregion

		public void loadLibraryItems() 
		{
			buckets.Clear();
			comboBox1.Items.Clear();
             IEnumerator bucketEnum =null;
            try
            {
                LibraryManager lm = LibraryManager.getInstance();
                 bucketEnum = lm.getLibraryBuckets();
            }
            catch (Exception e)
            {
                Log.exception(e,"Error getting library Buckets ");
            }

			while (bucketEnum.MoveNext()) 
			{			 
				
				LibraryBucket bucket = (LibraryBucket )bucketEnum.Current;
				if (bucket.getBucketName().StartsWith("__readonly"))
				{
					buckets.Add(bucket);
					comboBox2.Items.Add(bucket.getBucketName());
				}
				
			}
            if (comboBox2.Items.Count > 0)
            {
                comboBox2.SelectedIndex = 0;
            }
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			if (comboBox1.SelectedIndex == -1) 
			{
				MessageBox.Show("Select a Library Item", "Invalid Selection");
				return;
			}

			// Set return values
			SelectedItem = comboBox1.SelectedItem;

			DialogResult = DialogResult.OK;
			this.Close();
		
		}

		private void LibraryItem_Load(object sender, System.EventArgs e)
		{
		
		}

		private void comboBox2_SelectedIndexChanged(object sender, System.EventArgs e)
		{
		//	comboBox2.Items.Clear();
			comboBox1.Items.Clear();

			if (comboBox2.SelectedIndex == -1) 
			{
				return;
			}

			// load the items combo based on the selected bucket
			LibraryBucket b = buckets[comboBox2.SelectedIndex] as LibraryBucket;

			IEnumerator elementEnum = b.getElements().iterator();
			while (elementEnum.MoveNext()) 
			{
				LibraryElement libElement = (LibraryElement )elementEnum.Current;
				comboBox1.Items.Add(libElement.getElementName());
			}
		}

		

	}
	

	
}
