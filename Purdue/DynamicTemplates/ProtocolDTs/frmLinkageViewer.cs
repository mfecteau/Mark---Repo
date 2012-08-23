using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;

using Tspd.Tspddoc;
using Tspd.Businessobject;
using Tspd.Icp;
using MSXML2;
namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for frmLinkageViewer.
	/// </summary>
	public class frmLinkageViewer : System.Windows.Forms.Form
	{
		private System.Windows.Forms.Panel panel1;
		private System.Windows.Forms.Button cmdClose;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public frmLinkageViewer()
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
			this.panel1 = new System.Windows.Forms.Panel();
			this.cmdClose = new System.Windows.Forms.Button();
			this.panel1.SuspendLayout();
			this.SuspendLayout();
			// 
			// panel1
			// 
			this.panel1.Controls.Add(this.cmdClose);
			this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
			this.panel1.Location = new System.Drawing.Point(0, 548);
			this.panel1.Name = "panel1";
			this.panel1.Size = new System.Drawing.Size(1248, 56);
			this.panel1.TabIndex = 0;
			// 
			// cmdClose
			// 
			this.cmdClose.Location = new System.Drawing.Point(552, 24);
			this.cmdClose.Name = "cmdClose";
			this.cmdClose.TabIndex = 1;
			this.cmdClose.Text = "Close";
			this.cmdClose.Click += new System.EventHandler(this.cmdClose_Click);
			// 
			// frmLinkageViewer
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(1248, 604);
			this.Controls.Add(this.panel1);
			this.Name = "frmLinkageViewer";
			this.Text = "Link Viewer";
			this.TopMost = true;
			this.panel1.ResumeLayout(false);
			this.ResumeLayout(false);

		}
		#endregion


		BusinessObjectMgr bom_= null;
		double outcomeTotal =0;
		double perObjectiveTotal =0;
		double ObjectiveTotal =0;

		public TreeListView tlv = null;

		public void Load_Data(TspdDocument tspdDoc_)
		{
			bom_ = tspdDoc_.getBom();
			
			//createMainNode();

			tlv = new TreeListView();

			tlv.Location = new System.Drawing.Point(8, 0);
			tlv.Name = "treeListView1";
			tlv.Size = new System.Drawing.Size(540, 250);
			tlv.TabIndex = 1;
			tlv.Anchor = AnchorStyles.Top;
			tlv.FullRowSelect = true;
			tlv.ShowPlusMinus = true;
			tlv.ShowLines = true;
			tlv.ShowRootLines = true;
			//tlv.Font = new Font("Microsoft Sans Serif",11.25F);
			tlv.Font = new Font("Microsoft Sans Serif",11.25F);
			tlv.ForeColor = Color.Blue;
			

			ToggleColumnHeader colHeader = new ToggleColumnHeader();
			colHeader.Text = "Element";
			colHeader.Width = 400;
			tlv.Columns.Add(colHeader);

			colHeader = new ToggleColumnHeader();
			colHeader.Text = "Type";
			colHeader.Width = 300;
			tlv.Columns.Add(colHeader);

			colHeader = new ToggleColumnHeader();
			colHeader.Text = "Details";
			colHeader.Width = 500;
			tlv.Columns.Add(colHeader);

			colHeader = new ToggleColumnHeader();
			colHeader.Text = "Cost";
			colHeader.Width = 75;
			tlv.Columns.Add(colHeader);
            
			this.Controls.Add(tlv);

			this.cmdClose.Click += new System.EventHandler(this.cmdClose_Click);
			

			tlv.Dock = System.Windows.Forms.DockStyle.Fill;

			addNodes();

			this.ShowDialog();
		}

		private void addNodes()
		{
			foreach (Control c in this.Controls)
			{
				if (c.Name == "treeListView1")
				{
					
					//	c.KeyPress = new System.EventHandler(this.mycontrol_keypress);

					TreeListView tlv = (TreeListView)c;

					//					tlv.KeyPress = new System.EventHandler(this.mycontrol_keypress);

					TreeListNode tln = new TreeListNode();

					objectives(tlv);

					OutcomesTask(tlv);

					OrphanTask(tlv);

					////					tln = new TreeListNode();
					////					tln.Text = "Orphan Tasks";
					////					tln.SubItems.Add("Objective");
					////					tln.SubItems.Add("objective details");
					////					tln.SubItems.Add("$50"); 
					////					tlv.Nodes.Add(tln);

					
				}//end if
			}//end foreach
		}//end function


		public void mycontrol_keypress(object sender, System.EventArgs e)
		{
			MessageBox.Show("Success");
			MessageBox.Show(tlv.Nodes.Count.ToString());

		}
		
		private void createMainNode()
		{
			//	tv_Links.Nodes.Add("Objectives");
			
			//	tv_Links.Nodes.Add("Outcomes");
			//outcomes();
			//	tv_Links.Nodes.Add("Orphan Tasks");

			//objectives();
		}

		
		private void OutcomesTask(TreeListView tlv)
		{
			TreeListNode tln = new TreeListNode();
			tln.Text = "Outcomes without an Objective";
			tln.SubItems.Add("");
			tln.SubItems.Add("Collection of protocol outcomes with no associated objective");
			tln.SubItems.Add(""); 
			tln.ForeColor = Color.Firebrick;
			tlv.Nodes.Add(tln);

			OutcomeEnumerator oe = bom_.getOutcomes();
			int count = bom_.getOutcomes().getList().Count;
			int j=0;
			while(oe.MoveNext()) 
			{	
				Outcome outcome1 = (Outcome)oe.Current;
				ObjectiveEnumerator objEnum =  bom_.getAssociatedObjectives(outcome1);	
				if (objEnum.getList().Count <= 0)
				{
					TreeListNode tln2 = new TreeListNode();
					tln2.Text = outcome1.getActualDisplayValue();
					tln2.SubItems.Add("Outcome - " + outcome1.getOutcomeType().ToString());
					tln2.SubItems.Add(outcome1.getFullDescription());
					tln2.SubItems.Add("");
					tln.Nodes.Add(tln2);
					
					outcomeTotal =0;
					get_TaskVisits(outcome1,-1,j,tln);
					//Update the COST of each Outcome - {Summation of all TASK}
						tln.Nodes[j].SubItems[2].Text = "$" + outcomeTotal.ToString();
					
					j++;
				}			
			}
		}

		private void objectives(TreeListView tlv)
		{
			IEnumerator ie = bom_.getObjectives();
			int cnt =0; 
			int count = bom_.getObjectives().getList().Count;


			TreeListNode tln = new TreeListNode();
			tln.Text = "Objectives";
			tln.SubItems.Add("");
			tln.SubItems.Add("Collection of protocol objectives");
			tln.SubItems.Add("$5725");
			tln.ForeColor = Color.Blue;
			//tln.nodeFont = new Font("Verdana",8F,System.Drawing.FontStyle.Italic);
			

			if (count > 0)
			{
				while(ie.MoveNext()) 
				{
					Objective obj = (Objective)ie.Current;
					//			tv_Links.Nodes[0].Nodes.Add();

					TreeListNode tln2 = new TreeListNode();
					tln2.Text = obj.getActualDisplayValue();
					tln2.SubItems.Add("Objective - " + obj.getObjectiveType().ToString());
					tln2.SubItems.Add(obj.getFullDescription());
					tln2.SubItems.Add("$2629");
					tln.Nodes.Add(tln2);
					outcomes(obj,cnt,tln);

					ObjectiveTotal += perObjectiveTotal;
					

					cnt++;
				}

				tlv.Nodes.Add(tln);		
				//	tln.SubItems[2].Text = "$ " +  ObjectiveTotal.ToString();

				
			}			
		}

		private void outcomes(Objective currObj,int i,TreeListNode tln)
		{
			OutcomeEnumerator oe = bom_.getOutcomes();
			int count = bom_.getOutcomes().getList().Count;
			int j=0;

			TreeListNode tln2 = new TreeListNode();
			tln2.Text = "Outcomes";
			tln2.SubItems.Add("");
			tln2.SubItems.Add("Collection of protocol outcomes");
			tln2.SubItems.Add("$2527");
			tln2.ForeColor = Color.Firebrick;			
		//	tln2.nodeFont = new System.Drawing.Font("Calibri",10F,System.Drawing.FontStyle.Italic);
			tln.Nodes[i].Nodes.Add(tln2);


			while(oe.MoveNext()) 
			{	
				Outcome outcome1 = (Outcome)oe.Current;
				ObjectiveEnumerator objEnum =  bom_.getAssociatedObjectives(outcome1);			

				while(objEnum.MoveNext())
				{
					Objective PriObj = (Objective)objEnum.Current;
					
					if (PriObj.getObjectiveType()== currObj.getObjectiveType())
					{
						//outcomes.Add(outcome1);
						//				tv_Links.Nodes[0].Nodes[i].Nodes.Add(outcome1.getActualDisplayValue());
						tln2 = new TreeListNode();
						tln2.Text = outcome1.getActualDisplayValue();
						tln2.SubItems.Add("Outcome - " + outcome1.getOutcomeType().ToString());
						tln2.SubItems.Add(outcome1.getFullDescription());
						tln2.SubItems.Add("");
					
						tln.Nodes[i].Nodes[0].Nodes.Add(tln2);
					
						outcomeTotal =0;  //Reset
						get_TaskVisits(outcome1,i,j,tln);

						//Update the COST of each Outcome - {Summation of all TASK}
						tln.Nodes[i].Nodes[0].Nodes[j].SubItems[2].Text = "$" + outcomeTotal.ToString();
						j++;
					}
				}
				//	tlv.Nodes.Add(tln);
			}			
		}

		private void get_TaskVisits(Outcome currOut,int i, int j,TreeListNode tln)
		{

			//Task Visits
			double groupTaskCost = 0;  //Summation of all TASK COST PER OUTCOME
			TaskVisitEnumerator tve  =  bom_.getAssociatedTaskVists(currOut);
			 
			if (tve.getList().Count <= 0)
			{
				return;
			}
			
			int k=0;
			TreeListNode tln2 = new TreeListNode();
			tln2.Text = "Tasks";
			tln2.SubItems.Add("");
			tln2.SubItems.Add("Collection of tasks");
			tln2.SubItems.Add("");
			tln2.ForeColor = Color.CadetBlue;					

			try
			{
				if(i == -1)  ///Outcome with no objectives
				{
					tln.Nodes[j].Nodes.Add(tln2);
				}
				else
				{
					tln.Nodes[i].Nodes[0].Nodes[j].Nodes.Add(tln2);
				}
		//		tln2.nodeFont = new Font("Times New Roman",10F,FontStyle.Italic);

				ArrayList taskName = new ArrayList();  /// Contains all TaskNames to avoid Duplication
				double totalTaskCost = 0;

				while(tve.MoveNext())
				{
					TaskVisit tv = tve.Current as TaskVisit; 			
					Task task = bom_.getTaskByPath(tv.pathToAssociatedTask());

					if (taskName.Count > 0)
					{
						if (taskName.Contains(task.getActualDisplayValue()) == true)
						{
							//Exisiting Task - Just Update Total
							totalTaskCost += task.getCost();
							if(i == -1)
							{
								tln.Nodes[j].Nodes[0].SubItems[2].Text = "$" + totalTaskCost.ToString();
							}
							else
							{
								tln.Nodes[i].Nodes[0].Nodes[j].Nodes[0].Nodes[k-1].SubItems[2].Text = "$" + totalTaskCost.ToString();
							}
						}
						else
						{
							//Add New Task Node
							tln2 = new TreeListNode();
							tln2.Text = task.getActualDisplayValue();
							tln2.SubItems.Add("Task");
							tln2.SubItems.Add(task.getFullDescription());
							totalTaskCost = task.getCost();
							tln2.SubItems.Add("$" + totalTaskCost.ToString());		
							if(i ==-1)
							{
								tln.Nodes[j].Nodes[0].Nodes.Add(tln2);  //TASK	
							}
							else
							{
								tln.Nodes[i].Nodes[0].Nodes[j].Nodes[0].Nodes.Add(tln2);  //TASK	
							}
							taskName.Add(task.getActualDisplayValue());  
							k++;  //Sub Node incrementor
						}
					}
					else
					{
						// ONLY First time this code will be executed. -- Add New Task Node
						tln2 = new TreeListNode();
						tln2.Text = task.getActualDisplayValue();
						tln2.SubItems.Add("Task");
						tln2.SubItems.Add(task.getFullDescription());
						totalTaskCost = task.getCost();
						tln2.SubItems.Add("$" + totalTaskCost.ToString());	
						if(i ==-1)
						{
							tln.Nodes[j].Nodes[0].Nodes.Add(tln2); 
						}
						else
						{
							tln.Nodes[i].Nodes[0].Nodes[j].Nodes[0].Nodes.Add(tln2);  //TASK	
						}
						taskName.Add(task.getActualDisplayValue());  //Adding taskname to Array Litst.
						k++;   //Sub node incrementor
					}

				
					//Adding Subnode in either case. {New or Update}			
					tln2 = new TreeListNode();
					tln2.Text = tv.getActualDisplayValue();
					tln2.SubItems.Add("Task-Event");
					tln2.SubItems.Add(tv.getFullDescription());			
					tln2.SubItems.Add("$" + task.getCost().ToString());			


					if(i ==-1)
					{
						tln.Nodes[j].Nodes[0].Nodes[k-1].Nodes.Add(tln2); //Task Event
					}
					else
					{
						tln.Nodes[i].Nodes[0].Nodes[j].Nodes[0].Nodes[k-1].Nodes.Add(tln2); //Task Event
					}

					groupTaskCost += task.getCost();  //Adding on all task cost.
				}
				
				////Total Outcome Cost.
				outcomeTotal = groupTaskCost;
				//tln.Nodes[i].Nodes[0].Nodes[j].Nodes[0].SubItems[2].Text = "$" + groupTaskCost.ToString();			
			}
			catch(Exception ex)
			{
				MessageBox.Show(ex.ToString());
			}
		}

		
		
		private void OrphanTask(TreeListView tlv)
		{
			//TaskVisitEnumerator tve  = bom_.
			ArrayList taskvisitID = new ArrayList();
			ArrayList taskID = new ArrayList();
			SOAEnumerator soaEnum = bom_.getAllSchedules();

			///Adding Root node
			TreeListNode tln = new TreeListNode();
			tln.Text = "Orphan Tasks";
			tln.SubItems.Add("");
			tln.SubItems.Add("objective details");
			tln.SubItems.Add("$50"); 
			tlv.Nodes.Add(tln);
			tln.ForeColor = Color.CadetBlue;

			double totalTaskCost=0;

			bool flag = false;

			while(soaEnum.MoveNext())
			{
				SOA soa = (SOA )soaEnum.Current;		
				TaskEnumerator tskEnum =  soa.getTaskEnumerator();   //Getting Task
				while(tskEnum.MoveNext())
				{
					Task tsk = (Task) tskEnum.Current; 
					if(taskID.Contains(tsk.getObjID()) == false)
					{
						IEnumerator tvEnum = soa.getTaskVisitForTaskID(tsk.getObjID());  //Getting Task Event
						while(tvEnum.MoveNext())
						{
							IXMLDOMNode node = (IXMLDOMNode)tvEnum.Current;
							TaskVisit tv = new TaskVisit(node);												
							IEnumerator tvpEnum =    soa.getTaskVisitPurposes(tv);
							while (tvpEnum.MoveNext())
							{								
								TaskVisitPurpose tvp = (TaskVisitPurpose)tvpEnum.Current;
								if (tvp.getAssociatedOutcomeID() != 0)
								{
									flag = true;  //set it to be true.
								}
							}
						} //End Task Visit Event						
						if (flag == false)
						{
							taskID.Add(tsk.getObjID());
							TreeListNode tln2 = new TreeListNode();
							tln2.Text = tsk.getActualDisplayValue();
							tln2.SubItems.Add("Task ");
							tln2.SubItems.Add(tsk.getFullDescription());
							tln2.SubItems.Add("$ " + tsk.getCost().ToString());
							tln.Nodes.Add(tln2);
							totalTaskCost += tsk.getCost();
						}
						flag = false;
					}// End TASK
				}//End SOAEnum				
			}
			
			tln.SubItems[2].Text = "$" + totalTaskCost.ToString();
		}		
		
		private void cmdClose_Click(object sender, System.EventArgs e)
		{
			this.Close();		
		}




 
	}
}
