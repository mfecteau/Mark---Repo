using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using Tspd.Tspddoc;
using Tspd.MacroBase;
using Tspd.Macros;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ICPXSLTSelect.
	/// </summary>
	public class ICPXSLTSelect : System.Windows.Forms.Form
	{
		// Used for returns
		public string SelectedBucketName = "";
		public string SelectedLibraryItemName = "";
		public LibraryElement SelectedLibraryElement = null;

		ArrayList buckets = new ArrayList();
		ArrayList elements = new ArrayList();

		bool _loading = true;

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox cboCategory;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.ComboBox cboItem;
		private System.Windows.Forms.Button btnOK;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ICPXSLTSelect()
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
			this.label1 = new System.Windows.Forms.Label();
			this.cboCategory = new System.Windows.Forms.ComboBox();
			this.label2 = new System.Windows.Forms.Label();
			this.cboItem = new System.Windows.Forms.ComboBox();
			this.btnOK = new System.Windows.Forms.Button();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(8, 16);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(224, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Select Category:";
			// 
			// cboCategory
			// 
			this.cboCategory.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cboCategory.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboCategory.Location = new System.Drawing.Point(16, 40);
			this.cboCategory.Name = "cboCategory";
			this.cboCategory.Size = new System.Drawing.Size(410, 21);
			this.cboCategory.TabIndex = 1;
			this.cboCategory.SelectedIndexChanged += new System.EventHandler(this.cboCategory_SelectedIndexChanged);
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(100, 16);
			this.label2.TabIndex = 2;
			this.label2.Text = "Select Item";
			// 
			// cboItem
			// 
			this.cboItem.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.cboItem.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.cboItem.Location = new System.Drawing.Point(16, 96);
			this.cboItem.Name = "cboItem";
			this.cboItem.Size = new System.Drawing.Size(410, 21);
			this.cboItem.TabIndex = 3;
			// 
			// btnOK
			// 
			this.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.btnOK.Location = new System.Drawing.Point(185, 128);
			this.btnOK.Name = "btnOK";
			this.btnOK.TabIndex = 4;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// ICPXSLTSelect
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(434, 160);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.cboItem);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.cboCategory);
			this.Controls.Add(this.label1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "ICPXSLTSelect";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Select Parameters";
			this.TopMost = true;
			this.Load += new System.EventHandler(this.ICPXSLTSelect_Load);
			this.ResumeLayout(false);

		}
		#endregion

		private void ICPXSLTSelect_Load(object sender, System.EventArgs e)
		{
			string filter = "__XSL";

			// Load the buckets matching filter
			buckets.Clear();
			elements.Clear();

			_loading = true;

			cboCategory.Items.Clear();
			cboItem.Items.Clear();

			LibraryManager lm = LibraryManager.getInstance();
			IEnumerator bucketEnum = lm.getLibraryBuckets();
			while (bucketEnum.MoveNext()) 
			{
				LibraryBucket bucket = (LibraryBucket )bucketEnum.Current;
				string bucketName = bucket.getBucketName();

				if (bucketName.Length > filter.Length && bucketName.StartsWith(filter))
				{
					buckets.Add(bucket);
					cboCategory.Items.Add(bucketName.Substring(filter.Length));
				}
			}

			_loading = false;
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{
			// Validation checks
			if (cboCategory.SelectedIndex == -1) 
			{
				MessageBox.Show("Select a Category", "Invalid Selection");
				return;
			}

			if (cboItem.SelectedIndex == -1) 
			{
				MessageBox.Show("Select an Item", "Invalid Selection");
				return;
			}
		
			// Set up returns
			LibraryBucket b = buckets[cboCategory.SelectedIndex] as LibraryBucket;
			SelectedLibraryElement = elements[cboItem.SelectedIndex] as LibraryElement;
			SelectedBucketName = b.getBucketName();
			SelectedLibraryItemName = SelectedLibraryElement.getElementName();

			// Clean up
			buckets.Clear();
			elements.Clear();

			// Done
			DialogResult = DialogResult.OK;
			this.Close();
		}

		private void cboCategory_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			// Ignore if we are loading the combo
			if (_loading) return;

			// Clear
			cboItem.Items.Clear();
			elements.Clear();

			if (cboCategory.SelectedIndex == -1) 
			{
				return;
			}

			// load the items combo based on the selected bucket
			LibraryBucket b = buckets[cboCategory.SelectedIndex] as LibraryBucket;

			IEnumerator elementEnum = b.getElements().iterator();
			while (elementEnum.MoveNext()) 
			{
				LibraryElement libElement = (LibraryElement )elementEnum.Current;
				if (libElement.getContentType() == LibraryContentType.TEXT) 
				{
					elements.Add(libElement);
					cboItem.Items.Add(libElement.getElementName());
				}
			}
		}
	}
}
