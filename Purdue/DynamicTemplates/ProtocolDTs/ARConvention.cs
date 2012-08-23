using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.IO;
using System.Data;

using Tspd.Tspddoc;
using Tspd.Businessobject;
using Tspd.Bridge;
using Tspd.Utilities;

using Word = Microsoft.Office.Interop.Word;

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for ARConvention.
	/// </summary>
	public class ARConvention : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.ComboBox comboBox1;
		private System.Windows.Forms.ComboBox comboBox2;
		/// <summary>
		/// Required designer variable.
		/// 

		ArrayList buckets = new ArrayList();
		public ArrayList selList = new ArrayList();
		public ArrayList specialList = new ArrayList();
		public DataRow[] sortedRows;
		public Word.Bookmarks Sel_Bookmark;
		public bool col_flag = false;
		public string myText ="";

		LibraryManager lm = null;
		private System.Windows.Forms.RichTextBox richTextBox1;
		private System.Windows.Forms.TextBox txtToolTip;
		private System.Windows.Forms.Button btnOK;
		private System.Windows.Forms.PictureBox picImageFile;
		private System.Windows.Forms.ListView lstLibItems;
		private System.Windows.Forms.ColumnHeader columnHeader1;
		private System.Windows.Forms.Button btnCancel;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.RichTextBox richTextBox2;
		private System.Windows.Forms.Label label4;

		
		/// </summary>
		private System.ComponentModel.Container components = null;

		public ARConvention()
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
			this.txtToolTip = new System.Windows.Forms.TextBox();
			this.richTextBox1 = new System.Windows.Forms.RichTextBox();
			this.comboBox1 = new System.Windows.Forms.ComboBox();
			this.comboBox2 = new System.Windows.Forms.ComboBox();
			this.btnOK = new System.Windows.Forms.Button();
			this.picImageFile = new System.Windows.Forms.PictureBox();
			this.lstLibItems = new System.Windows.Forms.ListView();
			this.columnHeader1 = new System.Windows.Forms.ColumnHeader();
			this.btnCancel = new System.Windows.Forms.Button();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label4 = new System.Windows.Forms.Label();
			this.richTextBox2 = new System.Windows.Forms.RichTextBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label2 = new System.Windows.Forms.Label();
			this.groupBox1.SuspendLayout();
			this.groupBox2.SuspendLayout();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 12F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(16, 8);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(736, 16);
			this.label1.TabIndex = 0;
			this.label1.Text = "Library Items: ";
			// 
			// txtToolTip
			// 
			this.txtToolTip.ForeColor = System.Drawing.Color.Red;
			this.txtToolTip.Location = new System.Drawing.Point(8, 40);
			this.txtToolTip.Multiline = true;
			this.txtToolTip.Name = "txtToolTip";
			this.txtToolTip.Size = new System.Drawing.Size(648, 32);
			this.txtToolTip.TabIndex = 0;
			this.txtToolTip.Text = "";
			// 
			// richTextBox1
			// 
			this.richTextBox1.Location = new System.Drawing.Point(8, 88);
			this.richTextBox1.Name = "richTextBox1";
			this.richTextBox1.Size = new System.Drawing.Size(648, 192);
			this.richTextBox1.TabIndex = 2;
			this.richTextBox1.Text = "";
			// 
			// comboBox1
			// 
			this.comboBox1.Location = new System.Drawing.Point(16, 32);
			this.comboBox1.Name = "comboBox1";
			this.comboBox1.Size = new System.Drawing.Size(336, 21);
			this.comboBox1.TabIndex = 3;
			this.comboBox1.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged_1);
			// 
			// comboBox2
			// 
			this.comboBox2.Location = new System.Drawing.Point(664, 200);
			this.comboBox2.Name = "comboBox2";
			this.comboBox2.Size = new System.Drawing.Size(328, 21);
			this.comboBox2.TabIndex = 4;
			this.comboBox2.Visible = false;
			// 
			// btnOK
			// 
			this.btnOK.Location = new System.Drawing.Point(112, 360);
			this.btnOK.Name = "btnOK";
			this.btnOK.Size = new System.Drawing.Size(75, 24);
			this.btnOK.TabIndex = 6;
			this.btnOK.Text = "OK";
			this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
			// 
			// picImageFile
			// 
			this.picImageFile.Location = new System.Drawing.Point(128, 88);
			this.picImageFile.Name = "picImageFile";
			this.picImageFile.Size = new System.Drawing.Size(368, 192);
			this.picImageFile.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.picImageFile.TabIndex = 8;
			this.picImageFile.TabStop = false;
			this.picImageFile.Visible = false;
			// 
			// lstLibItems
			// 
			this.lstLibItems.CheckBoxes = true;
			this.lstLibItems.Columns.AddRange(new System.Windows.Forms.ColumnHeader[] {
																						  this.columnHeader1});
			this.lstLibItems.Dock = System.Windows.Forms.DockStyle.Fill;
			this.lstLibItems.FullRowSelect = true;
			this.lstLibItems.Location = new System.Drawing.Point(3, 16);
			this.lstLibItems.MultiSelect = false;
			this.lstLibItems.Name = "lstLibItems";
			this.lstLibItems.Size = new System.Drawing.Size(322, 269);
			this.lstLibItems.TabIndex = 9;
			this.lstLibItems.View = System.Windows.Forms.View.Details;
			this.lstLibItems.ColumnClick += new System.Windows.Forms.ColumnClickEventHandler(this.lstLibItems_ColumnClick);
			this.lstLibItems.SelectedIndexChanged += new System.EventHandler(this.lstLibItems_SelectedIndexChanged);
			this.lstLibItems.ItemCheck += new System.Windows.Forms.ItemCheckEventHandler(this.lstLibItems_ItemCheck);
			// 
			// columnHeader1
			// 
			this.columnHeader1.Text = "Check All/Uncheck All";
			this.columnHeader1.Width = 318;
			// 
			// btnCancel
			// 
			this.btnCancel.Location = new System.Drawing.Point(200, 360);
			this.btnCancel.Name = "btnCancel";
			this.btnCancel.TabIndex = 10;
			this.btnCancel.Text = "Cancel";
			this.btnCancel.Click += new System.EventHandler(this.btnCancel_Click);
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.lstLibItems);
			this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox1.Location = new System.Drawing.Point(16, 64);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(328, 288);
			this.groupBox1.TabIndex = 11;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Available Prose";
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label4);
			this.groupBox2.Controls.Add(this.richTextBox2);
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.label2);
			this.groupBox2.Controls.Add(this.txtToolTip);
			this.groupBox2.Controls.Add(this.richTextBox1);
			this.groupBox2.Controls.Add(this.picImageFile);
			this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Italic, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.groupBox2.Location = new System.Drawing.Point(352, 64);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(664, 288);
			this.groupBox2.TabIndex = 12;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Preview Window";
			// 
			// label4
			// 
			this.label4.Location = new System.Drawing.Point(8, 296);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(112, 16);
			this.label4.TabIndex = 12;
			this.label4.Text = "Selected Prose:";
			this.label4.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// richTextBox2
			// 
			this.richTextBox2.Location = new System.Drawing.Point(16, 312);
			this.richTextBox2.Name = "richTextBox2";
			this.richTextBox2.Size = new System.Drawing.Size(648, 192);
			this.richTextBox2.TabIndex = 11;
			this.richTextBox2.Text = @"{\\rtf1\\ansi\\ansicpg1252\\deff0\\deflang1033{\\fonttbl{\\f0\\froman\\fprq2\\fcharset0 Times New Roman;}{\\f1\\fswiss\\fcharset0 Arial;}}\r\n{\\colortbl ;\\red255\\green0\\blue0;}\r\n{\\*\\generator Msftedit 5.41.15.1507;}\\viewkind4\\uc1\\pard\\sa120\\tx252\\f0\\fs24 Treatment differences in \\cf1 mean change from baseline \\cf0 to selected post-baseline \\cf1 visits\\cf0  for vital signs and physical characteristics will be assessed using a maximum likelihood-based mixed-effects repeated measures analysis using all the longitudinal observations at each post-baseline \\cf1 visit\\cf0 .  The model will include the fixed categorical effects of \\cf1 treatment, visit, and treatment-by-visit interaction and the continuous covariate of baseline result, where subject is treated as a random effect.\\cf0   The Kenward-Roger method will be used to determine denominator degrees of freedom.  SAS (V8.2) PROC MIXED will be used to perform the analysis.  \\par\r\n\\pard\\sa120\\b\\par\r\n\\b0  \\par\r\n\\pard\\f1\\fs20\\par\r\n}\r\n\\viewkind4\\uc1\\pard\\sa120\\f0\\fs24 The covariance structure to model the within-subject errors will be unstructured.  If the unstructured covariance structure leads to lack of convergence, Akaike\\rquote s Information Criteria will be used to select the best fitting covariancestructure.\\par\r\n\\pard\\sa120\\tx252\\par\r\n\\pard\\sa120\\b\\par\r\n\\b0  \\par\r\n\\pard\\f1\\fs20\\par\r\n}\r\n\0";
			this.richTextBox2.TextChanged += new System.EventHandler(this.richTextBox2_TextChanged);
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(8, 24);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(112, 16);
			this.label3.TabIndex = 10;
			this.label3.Text = "Guide Text:";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label2
			// 
			this.label2.Location = new System.Drawing.Point(8, 72);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(112, 16);
			this.label2.TabIndex = 9;
			this.label2.Text = "Preview Prose:";
			this.label2.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// ARConvention
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.AutoScroll = true;
			this.ClientSize = new System.Drawing.Size(1018, 388);
			this.Controls.Add(this.groupBox2);
			this.Controls.Add(this.groupBox1);
			this.Controls.Add(this.btnCancel);
			this.Controls.Add(this.btnOK);
			this.Controls.Add(this.comboBox1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.comboBox2);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
			this.Name = "ARConvention";
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "Library Item Chooser";
			this.TopMost = true;
			this.Load += new System.EventHandler(this.ARConvention_Load);
			this.groupBox1.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion
		
		private void ARConvention_Load(object sender, System.EventArgs e)
		{
		
		}

		public void Load_Form(Word.Bookmarks bkMark_List,string LibItemCode)
		{
			lm = LibraryManager.getInstance();
			IEnumerator bucketEnum = lm.getLibraryBuckets();

//			if (LibItemCode == "AE")
//			{
//				label1.Text = label1.Text + "Adverse Events";
//			}
//			else
//			{
//				if (LibItemCode != "Library Items")
//				{
//					label1.Text = label1.Text + LibItemCode;
//				}
//			}
//			
			load_ToolTip();   //Setting up tooltip for controls.
			Sel_Bookmark = bkMark_List;

			comboBox1.Items.Clear();
			string bucketName = "";
			while (bucketEnum.MoveNext()) 
			{	
				LibraryBucket bucket = (LibraryBucket)bucketEnum.Current;
				bucketName =bucket.getBucketName();
				if (bucketName.StartsWith("__"))
				{
					if (bucketName != "__guide")
					{

//						
						if ((bucketName.StartsWith("__chooser/")) &&  LibItemCode == "Library Items")
						{
							buckets.Add(bucket);
							string[] strname = bucketName.Split('/');
							string strFilter = strname[1];  //// 1 becoz 
							comboBox1.Items.Add(strFilter);
						}
						else if (bucketName.StartsWith("__A&R/" + LibItemCode))
						{
							buckets.Add(bucket);
							string[] strname = bucketName.Split('/');
							string strFilter = strname[2];
							comboBox1.Items.Add(strFilter);
						}
					}					
				}
			}
		}


		private void comboBox1_SelectedIndexChanged_1(object sender, System.EventArgs e)
		{
			comboBox2.Items.Clear();
			ArrayList arr_SortedData = new ArrayList();
			
			Font myFont = new Font("Tahoma",8,System.Drawing.FontStyle.Italic);
			Font myFont1 = new Font("Tahoma",8,System.Drawing.FontStyle.Regular);


			// load the items combo based on the selected 
			DataSet dsSorted = new DataSet();
			DataTable dtSorted = new DataTable();

			DataColumn dtCol1 = new DataColumn("Seq");
			DataColumn dtCol2 = new DataColumn("Obj");
			dtCol1.DataType = System.Type.GetType("System.Decimal");
			dtCol2.DataType = System.Type.GetType("System.Object");

			dtSorted.Columns.Add(dtCol1);
			dtSorted.Columns.Add(dtCol2);

			LibraryBucket b = buckets[comboBox1.SelectedIndex] as LibraryBucket;
			IEnumerator elementEnum = b.getElements().iterator();
			while (elementEnum.MoveNext()) 
			{
				LibraryElement libElement = (LibraryElement )elementEnum.Current;
				DataRow dtRow = dtSorted.NewRow();
				string strLibItem = libElement.getElementName();
				int start = strLibItem.IndexOf("[");
				int end = strLibItem.IndexOf("]");
				if ((start < 0) || (end < 0))
				{
					//IF NO UID   -- Sequence# fixed to 100.

					dtRow[0] = 100;    
					dtRow[1] = libElement;
				}
				else
				{				
					try
					{
						strLibItem = strLibItem.Substring(start,end);
						string[] arr_seq = strLibItem.Split('-');
						
						if (arr_seq.Length >1)
						{
							Decimal seq = Convert.ToDecimal(arr_seq[1]);
							dtRow[0] = seq;  //Assigning sequence# to Lib Item.
						}
						else
						{
							dtRow[0] = 99;	// Sequence# fixed to 99 for those without seq#.
						}

					
						dtRow[1] = libElement;

					}
					catch(Exception ex)
					{
						MessageBox.Show(ex.ToString());
					}
				}

				dtSorted.Rows.Add(dtRow);		
			}

			

			// Sort Ascending  by column named Seq
			string sortOrder = "Seq ASC";
			// Use the Select method to find all rows matching the filter.
			

			sortedRows = dtSorted.Select("", sortOrder);
			comboBox2.Items.Clear();
			
			lstLibItems.Items.Clear();
			richTextBox1.Visible = true;
			picImageFile.Visible = false;
			richTextBox1.Clear();
			txtToolTip.Clear();

			string[] filename;
			string strName;

			for (int i =0; i < sortedRows.Length;i++)
			{
				try
				{				
					LibraryElement libElement = (LibraryElement) sortedRows[i][1];
					comboBox2.Items.Add(libElement.getElementName());
					ListViewItem lvItem;		


					//if there is no UID defined.
					if (libElement.getElementName().IndexOf("]") < 0)
					{
						strName = libElement.getElementName();
						lvItem = new ListViewItem(new string[] {""},0,System.Drawing.Color.Black,System.Drawing.Color.White,myFont1);
						lvItem.Text = strName;	

						lstLibItems.Items.Add(lvItem);

					}
					else
					{
						filename = libElement.getElementName().Split(']');
						strName = filename[1];

						//Check if Bookmark exists
						if (BookMark_exist(libElement.getElementName()))
						{
							//CODE to make it italics.

							lvItem = new ListViewItem(new string[] {""},0,System.Drawing.Color.DarkSlateBlue,System.Drawing.Color.White,myFont);
							lvItem.Text = strName; 
						}
						else
						{						
							lvItem = new ListViewItem(new string[] {""},0,System.Drawing.Color.Black,System.Drawing.Color.White,myFont1);
							lvItem.Text = strName;						
						}
		
						lstLibItems.Items.Add(lvItem);
					}
				}
				catch(Exception ex)
				{
					MessageBox.Show(ex.ToString());
				}
			}

		}
		

		private bool BookMark_exist(string strCode)
		{
			string[] bkname = strCode.Split('-');
			string strName = bkname[0];
			string[] bkname2 = strName.Split(':');
			string bookMark = bkname2[1];

			for (int i=1; i<= Sel_Bookmark.Count; i++)
			{
				object oI = i;
				Word.Bookmark bk = Sel_Bookmark.get_Item(ref oI);
				string bmName = bk.Name;

				if (bmName.StartsWith(bookMark))
				{
					return true;
				}
			}
			return false;
		}

		

	
		private void btnOK_Click(object sender, System.EventArgs e)
		{
			

			if (lstLibItems.CheckedItems.Count <= 0)
			{
				//Message if there is nothing checked
				MessageBox.Show("Please select atleast one Library Item!");
				this.DialogResult = DialogResult.Cancel;
			}
			else
			{
				
				foreach(int indexChecked in lstLibItems.CheckedIndices) 
				{
					LibraryElement libElement = (LibraryElement)sortedRows[indexChecked][1];
					selList.Add(libElement);
					
				}
				this.DialogResult = DialogResult.OK;
			}
		}

	
		
		private void lstLibItems_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			try
			{
				ListView.SelectedListViewItemCollection selItems = lstLibItems.SelectedItems;

				if (selItems.Count > 0)
				{	
					//Select the file.			
					int cnt =  selItems[0].Index;
					FillData_RichTXT(cnt);
					
				}
			}	
			catch(Exception ex)
			{
				MessageBox.Show("Lib Item Macro -- " + ex.ToString());
			}
		}

		private void lstLibItems_ItemCheck(object sender, System.Windows.Forms.ItemCheckEventArgs e)
		{
			try
			{
				ListView.CheckedListViewItemCollection chkItems = lstLibItems.CheckedItems;

				richTextBox2 .Rtf = "";
				myText = "";
				for(int i =0;i< chkItems.Count; i++)
				{
					int cnt =  chkItems[i].Index;
					PreviewData_RichTXT(cnt);
				}
				string cleanedText = myText.Replace("\0", "");
				richTextBox2.Rtf = cleanedText;;
			}	
			catch(Exception ex)
			{
				MessageBox.Show("Lib Item Macro -- " + ex.ToString());
			}
			
		}
		private void PreviewData_RichTXT(int index)
		{
			try
			{

				LibraryElement libElement = (LibraryElement) sortedRows[index][1];
				
				//Tooltip box.
				try
				{
					txtToolTip.Text = "";
					txtToolTip.Text = libElement.getTooltip();
				}
				catch(Exception ex)
				{
					//Ignore the Casting error - If any.
				}
			
				//Download file to local host.				
				if (libElement.getContentType() == LibraryContentType.IMAGE)
				{
					richTextBox1.Visible=false;
					picImageFile.Visible= true;					
					string path = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
					if (File.Exists(path))
					{
						Log.log(System.Diagnostics.TraceLevel.Verbose, "inserting file: "+ libElement.getFileName());
						object theRange = System.Reflection.Missing.Value;
						object confirm = false;
						object link = false;
						object attachment = System.Reflection.Missing.Value;
						//range.InsertFile(path, ref theRange, ref confirm, ref link, ref attachment);
					}
					Bitmap MyImage = new Bitmap(path);
					picImageFile.Image= MyImage;
				}
				else if (libElement.getContentType() == LibraryContentType.INLINE)
				{
					picImageFile.Visible = false;
					richTextBox1.Visible = true;
					
					//richTextBox1.Text= "";
					richTextBox2.Text= libElement.getInlineData();
				}
				else if ((libElement.getContentType() == LibraryContentType.MSWORD) || (libElement.getContentType() == LibraryContentType.TEXT))
				{
					picImageFile.Visible = false;
					richTextBox1.Visible = true;

					string path = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
					if (File.Exists(path))
					{
						Log.log(System.Diagnostics.TraceLevel.Verbose, "inserting file: "+ libElement.getFileName());
						object theRange = System.Reflection.Missing.Value;
						object confirm = false;
						object link = false;
						object attachment = System.Reflection.Missing.Value;
						//range.InsertFile(path, ref theRange, ref confirm, ref link, ref attachment);
					}
					//Read it in to Rich Text Box Control.

				//	richTextBox1.Clear();
					if (System.IO.Path.GetExtension(path) == ".rtf")
					{
						FileStream fs = File.OpenRead(path);
						StreamReader sr = new StreamReader(fs); 									
						 myText += sr.ReadToEnd();
						//richTextBox2.Rtf = myText;
						fs.Close();
						//richTextBox2.Rtf = myText;
					//	richTextBox2.Rtf = "\n";
					}
					else
					{
//						richTextBox2.AppendText("sds");						
//						richTextBox2.LoadFile(path,RichTextBoxStreamType.RichText);
					}
				}
			}
			catch(Exception ex)
			{
				MessageBox.Show("Lib Item Macro -- " + ex.ToString());
			}

		}


		private void FillData_RichTXT(int index)
		{
			try
			{
			LibraryElement libElement = (LibraryElement) sortedRows[index][1];
				
			//Tooltip box.
			try
			{
				txtToolTip.Text = "";
				txtToolTip.Text = libElement.getTooltip();
			}
			catch(Exception ex)
			{
				//Ignore the Casting error - If any.
			}
			
			//Download file to local host.				
			if (libElement.getContentType() == LibraryContentType.IMAGE)
			{
				richTextBox1.Visible=false;
				picImageFile.Visible= true;					
				string path = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
				if (File.Exists(path))
				{
					Log.log(System.Diagnostics.TraceLevel.Verbose, "inserting file: "+ libElement.getFileName());
					object theRange = System.Reflection.Missing.Value;
					object confirm = false;
					object link = false;
					object attachment = System.Reflection.Missing.Value;
					//range.InsertFile(path, ref theRange, ref confirm, ref link, ref attachment);
				}
				Bitmap MyImage = new Bitmap(path);
				picImageFile.Image= MyImage;
			}
			else if (libElement.getContentType() == LibraryContentType.INLINE)
			{
				picImageFile.Visible = false;
				richTextBox1.Visible = true;
					
				richTextBox1.Text= "";
				richTextBox1.Text= libElement.getInlineData();
			}
			else if ((libElement.getContentType() == LibraryContentType.MSWORD) || (libElement.getContentType() == LibraryContentType.TEXT))
			{
				picImageFile.Visible = false;
				richTextBox1.Visible = true;

				string path = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
				if (File.Exists(path))
				{
					Log.log(System.Diagnostics.TraceLevel.Verbose, "inserting file: "+ libElement.getFileName());
					object theRange = System.Reflection.Missing.Value;
					object confirm = false;
					object link = false;
					object attachment = System.Reflection.Missing.Value;
					//range.InsertFile(path, ref theRange, ref confirm, ref link, ref attachment);
				}
				//Read it in to Rich Text Box Control.

				richTextBox1.Clear();
				if (System.IO.Path.GetExtension(path) == ".doc")
				{
					FileStream fs = File.OpenRead(path);
					StreamReader sr = new StreamReader(fs); 									
					string myText = sr.ReadToEnd();
					richTextBox1.Text = myText;
					fs.Close();
					richTextBox1.Text= myText;
				}
				else
				{
					richTextBox1.LoadFile(path,RichTextBoxStreamType.RichText);
				}
			}
		}
		catch(Exception ex)
			{
				MessageBox.Show("Lib Item Macro -- " + ex.ToString());
			}

		}

		private void lstLibItems_ColumnClick(object sender, System.Windows.Forms.ColumnClickEventArgs e)
		{
			if (lstLibItems.Items.Count > 0)
			{
				if (col_flag == false)
				{
					col_flag =true;
				}
				else{
					col_flag = false;

				}
					for(int i=0; i <= lstLibItems.Items.Count; i++)
				{
					lstLibItems.Items[i].Checked = col_flag;
				}

			}
			txtToolTip.Clear();
			richTextBox1.Clear();
			//picImageFile.
		
		}

		private void load_ToolTip()
		{
			ToolTip tp = new ToolTip();
			tp.SetToolTip(comboBox1,"Select a category to view the available suggested language for the corresponding convention.");
			tp.SetToolTip(lstLibItems,"Highlight an item to preview the prose in the preview window. Check each item to be inserted into the document.");
			tp.SetToolTip(txtToolTip,"Instructional text which applies to the below prose.  Guide text will not be inserted into the document.");
			tp.SetToolTip(richTextBox1,"Suggested language to be inserted into the document.");

			
		}

		private void btnCancel_Click(object sender, System.EventArgs e)
		{
			this.Close();
			this.DialogResult = DialogResult.Cancel;
		}

		private void richTextBox2_TextChanged(object sender, System.EventArgs e)
		{
		
		}

		
	}
}
