using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Context;

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for StudyConductSel.
	/// </summary>
	public class StudyConductSel : System.Windows.Forms.Form
	{
		private System.Windows.Forms.ComboBox comboBox2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox comboBox1;
		private System.Windows.Forms.Label lbl1;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.CheckedListBox ChkLsb;
		public ArrayList SelectedItems;
		public string bucket_name,chooser_name;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public StudyConductSel()
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
			this.comboBox2 = new System.Windows.Forms.ComboBox();
			this.label1 = new System.Windows.Forms.Label();
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.lbl1 = new System.Windows.Forms.Label();
			this.btnOK = new System.Windows.Forms.Button();
			this.ChkLsb = new System.Windows.Forms.CheckedListBox();
			this.SuspendLayout();
			// 
			// comboBox2
			// 
			this.comboBox2.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.comboBox2.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBox2.Location = new System.Drawing.Point(15, 23);
			this.comboBox2.Name = "comboBox2";
			this.comboBox2.Size = new System.Drawing.Size(372, 21);
			this.comboBox2.TabIndex = 12;
			this.comboBox2.SelectedIndexChanged += new System.EventHandler(this.comboBox2_SelectedIndexChanged);
			// 
			// label1
			// 
			this.label1.Location = new System.Drawing.Point(15, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(152, 16);
			this.label1.TabIndex = 11;
			this.label1.Text = "Select a Study conduct:";
			// 
			// comboBox1
			// 
			this.comboBox1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
				| System.Windows.Forms.AnchorStyles.Right)));
			this.comboBox1.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBox1.Location = new System.Drawing.Point(16, 80);
			this.comboBox1.Name = "comboBox1";
			this.comboBox1.Size = new System.Drawing.Size(372, 21);
			this.comboBox1.TabIndex = 10;
			this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
			// 
			// lbl1
			// 
			this.lbl1.Location = new System.Drawing.Point(16, 64);
			this.lbl1.Name = "lbl1";
			this.lbl1.Size = new System.Drawing.Size(120, 16);
			this.lbl1.TabIndex = 9;
			this.lbl1.Text = "Select an Item:";
			// 
			// btnOK
			// 
			this.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
			this.btnOK.DialogResult = System.Windows.Forms.DialogResult.Cancel;
			this.btnOK.Location = new System.Drawing.Point(164, 176);
			this.btnOK.Name = "btnOK";
			this.btnOK.TabIndex = 8;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// ChkLsb
			// 
			this.ChkLsb.Location = new System.Drawing.Point(16, 120);
			this.ChkLsb.Name = "ChkLsb";
			this.ChkLsb.Size = new System.Drawing.Size(376, 49);
			this.ChkLsb.TabIndex = 13;
			this.ChkLsb.SelectedIndexChanged += new System.EventHandler(this.ChkLsb_SelectedIndexChanged);
			// 
			// StudyConductSel
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(402, 204);
			this.Controls.Add(this.ChkLsb);
			this.Controls.Add(this.comboBox2);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.comboBox1);
			this.Controls.Add(this.lbl1);
			this.Controls.Add(this.btnOK);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "StudyConductSel";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Study Conduct";
			this.TopMost = true;
			this.ResumeLayout(false);

		}
		#endregion

		BusinessObjectMgr bom = null;
		ChooserEntry Sel_Entry = null;
		ArrayList bucketname = new ArrayList();
		ArrayList sel_List = new ArrayList();

		public void loadStudyconducts() 
		{
			btnOK.Enabled = false;
			bom = ContextManager.getInstance().getActiveDocument().getBom();
			BucketEntry bucketEntry = bom.getIcd().getBucketBySystemName(ElementType.STUDY_CONDUCT_COMPLIANCE);
			BucketEntry bucketEntry1 = bom.getIcd().getBucketBySystemName(ElementType.STUDY_CONDUCT_TERMINATION);

			comboBox2.Items.Add(bucketEntry.getBucketLabel());
			comboBox2.Items.Add(bucketEntry1.getBucketLabel());


			comboBox2.SelectedIndex =-1;

			
			IEnumerator termination = bom.getIcd().getChooserEntriesForBucketEntry(bucketEntry1);

			bucketname.Add(bucketEntry);
			bucketname.Add(bucketEntry1);
		
		}

		private void StudyConductItem(BusinessObjectMgr thisBom_)
		{
			comboBox2.Items.Clear();			
			comboBox2.SelectedIndex = -1;
		}

		private void comboBox2_SelectedIndexChanged(object sender, System.EventArgs e)
		{

			if (comboBox2.SelectedIndex != -1 )
				//DocType docType = ;

			{
				ChkLsb.Items.Clear();
				comboBox1.Items.Clear(); //Clearing items in combo.
				sel_List.Clear();  //Clearing ArrayList

				BucketEntry entry1 = (BucketEntry)bucketname[comboBox2.SelectedIndex];
				IEnumerator compliance = bom.getIcd().getChooserEntriesForBucketEntry(entry1);
				while (compliance.MoveNext())
				{
					ChooserEntry ce = (ChooserEntry)compliance.Current;
					sel_List.Add(ce);
					comboBox1.Items.Add(ce.getActualDisplayValue());
				}
			}
		}

		private void comboBox1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (comboBox1.SelectedIndex != -1)
			{
				ChkLsb.Items.Clear();
				ChooserEntry sel_entry = (ChooserEntry)sel_List[comboBox1.SelectedIndex];
				StringListHelper sh = bom.getIcp().getStringList(sel_entry.getElementPath(),ContextManager.getInstance().getActiveDocument().getDocType());
				ArrayList sh1 =  new ArrayList();
				sh1 = bom.getIcp().getStringListValues(sel_entry.getElementPath());

				if (sh1.Count == 0)
				{
					btnOK.Enabled = false;
					return;
				}
				int i = 0;
				for (i =0; i<sh1.Count;i++)
				{
					ChkLsb.Items.Add(sh1[i]);
				}	
//				btnOK.Enabled = true;			
			}		
		}

		private void btnOK_Click(object sender, System.EventArgs e)
		{	
			if((comboBox1.SelectedIndex == -1) || (comboBox2.SelectedIndex == -1))
			{
				MessageBox.Show("Please Select Study Conduct!","Invalid Selection");
				return;				
			}
			else
			{						 
				SelectedItems = new ArrayList();
				if (ChkLsb.Items.Count == 0)
				{
					SelectedItems.Add("No Study Conduct Items are defined for " + comboBox1.SelectedItem.ToString() + "\\" + comboBox2.SelectedItem.ToString() + ".");
					this.DialogResult = DialogResult.OK;				
				}
				int i =0;
	
				System.Windows.Forms.CheckedListBox.CheckedItemCollection	LstSel;
				LstSel = ChkLsb.CheckedItems;
			


				for (i =0; i <LstSel.Count;i++)
				{
					SelectedItems.Add(LstSel[i].ToString());				
				}

				bucket_name = comboBox2.SelectedItem.ToString();
				chooser_name = comboBox1.SelectedItem.ToString();



				this.DialogResult = DialogResult.OK;
			}
				}

		private void ChkLsb_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			if (ChkLsb.CheckedItems.Count == 0)
			{
				btnOK.Enabled =false;
			}
			else
			{
				btnOK.Enabled= true;
			}
		}
		}
}
