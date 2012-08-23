using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Xml;
using System.IO;
using Tspd.Businessobject;
using Tspd.Icp;
using Tspd.Utilities;
using Tspd.Bridge;
using Tspd.Tspddoc;

namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for TaskSelect.
	/// </summary>
	public class ProcedureList : System.Windows.Forms.Form
	{
		public int SelectedTask = -1;
        public string perORsubPer = "";
        public object sel_ObjID = 0;
        public Tspd.Icp.SOA _currentSOA = null;
        public ArrayList listObjID = new ArrayList();
        public ArrayList customTaskDec;  //For ReadXML method.
		private System.Windows.Forms.ComboBox cboPeriod;
		public System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button btnOK;
        public RadioButton rdPeriod;
        public RadioButton rdSubperiod;
        private ListView lstTaskEvent;
        private ListView lstTaskDesc;
        public XmlNode _selNode = null;
        public LibraryManager lm = null;
        ArrayList buckets = new ArrayList();

        Hashtable htBuckets = new Hashtable();
        Hashtable htCheckedIndices = new Hashtable();
        Hashtable htFinalTE = new Hashtable();
        
        RichTextBox rtBox = new RichTextBox();
        private RichTextBox richTextBox1;
        public string strTaskEvent = "";
        public string strTaskDesc = "";
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

        public ProcedureList()
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
            this.cboPeriod = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.rdPeriod = new System.Windows.Forms.RadioButton();
            this.rdSubperiod = new System.Windows.Forms.RadioButton();
            this.lstTaskEvent = new System.Windows.Forms.ListView();
            this.lstTaskDesc = new System.Windows.Forms.ListView();
            this.richTextBox1 = new System.Windows.Forms.RichTextBox();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.btnOK.DialogResult = System.Windows.Forms.DialogResult.OK;
            this.btnOK.Location = new System.Drawing.Point(268, 404);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(75, 23);
            this.btnOK.TabIndex = 6;
            this.btnOK.Text = "OK";
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // cboPeriod
            // 
            this.cboPeriod.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.cboPeriod.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.cboPeriod.Location = new System.Drawing.Point(12, 36);
            this.cboPeriod.Name = "cboPeriod";
            this.cboPeriod.Size = new System.Drawing.Size(606, 21);
            this.cboPeriod.TabIndex = 1;
            this.cboPeriod.SelectedIndexChanged += new System.EventHandler(this.cboPeriod_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left)
                        | System.Windows.Forms.AnchorStyles.Right)));
            this.label1.Location = new System.Drawing.Point(13, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(600, 15);
            this.label1.TabIndex = 0;
            this.label1.Text = "Select Period/Sub-Period:";
            // 
            // rdPeriod
            // 
            this.rdPeriod.AutoSize = true;
            this.rdPeriod.Location = new System.Drawing.Point(182, 12);
            this.rdPeriod.Name = "rdPeriod";
            this.rdPeriod.Size = new System.Drawing.Size(55, 17);
            this.rdPeriod.TabIndex = 7;
            this.rdPeriod.TabStop = true;
            this.rdPeriod.Text = "Period";
            this.rdPeriod.UseVisualStyleBackColor = true;
            this.rdPeriod.CheckedChanged += new System.EventHandler(this.rdTask_CheckedChanged);
            // 
            // rdSubperiod
            // 
            this.rdSubperiod.AutoSize = true;
            this.rdSubperiod.Location = new System.Drawing.Point(244, 12);
            this.rdSubperiod.Name = "rdSubperiod";
            this.rdSubperiod.Size = new System.Drawing.Size(77, 17);
            this.rdSubperiod.TabIndex = 8;
            this.rdSubperiod.TabStop = true;
            this.rdSubperiod.Text = "Sub-Period";
            this.rdSubperiod.UseVisualStyleBackColor = true;
            this.rdSubperiod.CheckedChanged += new System.EventHandler(this.rdVisit_CheckedChanged);
            // 
            // lstTaskEvent
            // 
            this.lstTaskEvent.Alignment = System.Windows.Forms.ListViewAlignment.Left;
            this.lstTaskEvent.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.lstTaskEvent.Cursor = System.Windows.Forms.Cursors.Default;
            this.lstTaskEvent.FullRowSelect = true;
            this.lstTaskEvent.Location = new System.Drawing.Point(12, 63);
            this.lstTaskEvent.MultiSelect = false;
            this.lstTaskEvent.Name = "lstTaskEvent";
            this.lstTaskEvent.Size = new System.Drawing.Size(263, 221);
            this.lstTaskEvent.TabIndex = 10;
            this.lstTaskEvent.UseCompatibleStateImageBehavior = false;
            this.lstTaskEvent.View = System.Windows.Forms.View.List;
            this.lstTaskEvent.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.lstTaskEvent_ItemSelectionChanged);
            // 
            // lstTaskDesc
            // 
            this.lstTaskDesc.CheckBoxes = true;
            this.lstTaskDesc.FullRowSelect = true;
            this.lstTaskDesc.Location = new System.Drawing.Point(281, 62);
            this.lstTaskDesc.MultiSelect = false;
            this.lstTaskDesc.Name = "lstTaskDesc";
            this.lstTaskDesc.Scrollable = false;
            this.lstTaskDesc.Size = new System.Drawing.Size(336, 222);
            this.lstTaskDesc.TabIndex = 11;
            this.lstTaskDesc.UseCompatibleStateImageBehavior = false;
            this.lstTaskDesc.View = System.Windows.Forms.View.List;
            this.lstTaskDesc.ItemChecked += new System.Windows.Forms.ItemCheckedEventHandler(this.lstTaskDesc_ItemChecked);

            this.lstTaskDesc.ItemSelectionChanged += new System.Windows.Forms.ListViewItemSelectionChangedEventHandler(this.lstTaskDesc_ItemSelectionChanged);
            // 
            // richTextBox1
            // 
            this.richTextBox1.Location = new System.Drawing.Point(12, 290);
            this.richTextBox1.Name = "richTextBox1";
            this.richTextBox1.Size = new System.Drawing.Size(606, 108);
            this.richTextBox1.TabIndex = 12;
            this.richTextBox1.Text = "";
            // 
            // ProcedureList
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(624, 434);
            this.Controls.Add(this.richTextBox1);
            this.Controls.Add(this.lstTaskDesc);
            this.Controls.Add(this.lstTaskEvent);
            this.Controls.Add(this.rdSubperiod);
            this.Controls.Add(this.rdPeriod);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.cboPeriod);
            this.Controls.Add(this.btnOK);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "ProcedureList";
            this.ShowInTaskbar = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Task or Visit Selection";
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		public bool loadItems(Tspd.Icp.SOA _SOA,string path,Period per,bool _usePeriod,bool _isDummyCall) 
		{

            //Parm 4: _usePeriod =  indicates if its Period (true) or SubPeriod (false)
            //Para 4" _isDummyCall = true; do not initialize radio buttons as it will run all the code for its events.
		
            _currentSOA = _SOA;
            cboPeriod.Items.Clear();
            sel_ObjID = -1;
            htBuckets.Clear();
            lm = LibraryManager.getInstance();  //global LibraryMgr Object
            Initialize_LibItems();  // Fill the hash table.

            ///  Get an Array List of Nodes (task) from the Config XML.
            customTaskDec = ReadXML(path);

            if (per != null)  //If period is NOT null, we do not need to check, just have default settings
            {
                if (!LoadMeorNot(per))
                {
                    //If Func returns "false" - We do not need to load form. 
                    return false;
                }
                else
                {
                    //If Func returns "true" - We HAVE  to load form. 
                    if (!_isDummyCall)
                    {
                        if (_usePeriod)
                        {
                            rdPeriod.Checked = true;  //Default task is checked.
                        }
                        else
                        {
                            rdSubperiod.Checked = true;
                        }

                        int idx = cboPeriod.Items.IndexOf(per.getBriefDescription());
                        cboPeriod.SelectedIndex = idx;
                    }
                }
            }
            else
            {
                //If PERIOD = NULL;
                rdPeriod.Checked = true;
                cboPeriod.SelectedIndex = -1;
            }
            return true;
           
		}

        private bool LoadMeorNot(Period per)
        {
            //This method will run sort of pre-check to see, if there are any visit specific task with empty description.
            //"Default" task should be updated, with out notification.


            bool flag4formload = false;   //Flag is set to true for loading form.
           // Period per = null;

            ArrayList tskList = GetTaskListforPeriod(per); //Get all the task with TaskEvents for Period
            ArrayList taskEvents = GetTaskEvents(per, tskList);  // Get all task events for selected Period and Task's.
               Task tsk = null;
               foreach (long tskID in tskList)
               {
                   tsk = _currentSOA.getTaskByID(tskID);
                   if (tsk != null)
                   {
                       if (TaskExistinXML(tsk.getBriefDescription()))
                       {
                           //Check for each task, in the Xml, if there is any visit specific 
                           foreach (XmlNode cnode in _selNode.ChildNodes)
                           {
                               if (cnode.Attributes.GetNamedItem("type").Value.ToLower().Equals("visit-specific"))
                               {
                                   foreach (TaskVisit pe in taskEvents)
                                   {
                                       if (pe.getFullDescription().Trim().Length <= 0 && pe.getAssociatedTaskID() == tsk.getObjID())  //If EMPTY
                                       {
                                           //Set the flag, if there is any Task Events has empty full description.
                                           flag4formload = true;
                                       }
                                   }

                               }
                               else if (cnode.Attributes.GetNamedItem("type").Value.ToLower().Equals("default"))
                               {
                                   //Call Method for handling default types
                                   //Get Bucket from Hash Table, based on 

                                   if (htBuckets.ContainsKey(cnode.Attributes.GetNamedItem("libraryItem").Value))
                                   {
                                       LibraryElement libElement = (LibraryElement)htBuckets[cnode.Attributes.GetNamedItem("libraryItem").Value];

                                       string filePath = null;
                                       try
                                       {
                                           filePath = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
                                       }
                                       catch (Exception e)
                                       {
                                           Log.exception(e, cnode.Attributes.GetNamedItem("libraryItem").Value + " - library item not found!");
                                       }
                                       _HandleDefault(filePath, tsk);
                                   }
                                   else
                                   {
                                       Log.trace(cnode.Attributes.GetNamedItem("libraryItem").Value + " - not found.");
                                   }
                               } //end elseif
                           }
                       }
                   } //Endif Tsk != null
               }

            

            return flag4formload;
        }

		private void btnOK_Click(object sender, System.EventArgs e)
		{
            if (cboPeriod.SelectedIndex >= 0)
            {
                sel_ObjID = listObjID[cboPeriod.SelectedIndex];

                Save_TaskEvents();

                DialogResult = DialogResult.OK;
                this.Close();
            }
            else 
            {
                sel_ObjID = -1;
                DialogResult = DialogResult.OK;
                this.Close();
            }
		}

        private void Save_TaskEvents()
        {
            //This method will save description for selected TaskEvents
            //HashTable consists of
            //(1) Key = taskEventID ">" TaskID
            //(2) Value = "LibItem1" | "libItem2"

            string[] strID = null;
            string val = null;
            Task t1 = null;
            string TE_Desc = "";
            TaskVisit pe = null;
            bool isBad = false;


            //Update the Taskevents which are not being seleceted by user with Standard Mesg text. So they do not popup again.
            //User has to clear that message from Design guide --> Task Event
            foreach (ListViewItem lvtItem in lstTaskEvent.Items)
            {
                if (!htFinalTE.ContainsKey(lvtItem.Name))
                {
                    strID = null;
                    strID = lvtItem.Name.Split('>');
                    if (strID.Length > 1)
                    {
                        t1 = _currentSOA.getTaskByID(PurdueUtil.getNumber(strID[1], out isBad));
                        pe = _currentSOA.getTaskVisitById(PurdueUtil.getNumber(strID[0], out isBad));
                    }
                    //update the taskevent description.
               //    MessageBox.Show(pe.getActualDisplayValue() + " (TaskEvent) has been updated. Msg3 ");
                  pe.setFullDescription("<No Description is selected.>"); 
                }
 
            }


            foreach (string key in htFinalTE.Keys)
            {
                strID = null;
                TE_Desc = null;
                strID = key.Split('>');
                if (strID.Length > 1)
                {
                    t1 = _currentSOA.getTaskByID(PurdueUtil.getNumber(strID[1], out isBad));
                    pe = _currentSOA.getTaskVisitById(PurdueUtil.getNumber(strID[0], out isBad));
                }

                if (pe != null)
                {
                    val = htFinalTE[key].ToString();
                    strID = val.Split('|');
                    foreach (string itemName in strID)
                    {
                        if (itemName.Trim().Length > 0)
                        {
                            TE_Desc += ReadfromLibrary(itemName);
                            TE_Desc += "\r\n";
                        }
                    }

                    //update the taskevent description.
                   // MessageBox.Show(pe.getActualDisplayValue() + " (TaskEvent) has been updated. Msg2 ");
                    pe.setFullDescription(TE_Desc);
                }
            }
        }


        private string ReadfromLibrary(string key_itemName)
        {
            LibraryElement libElement = (LibraryElement)htBuckets[key_itemName];
            string filePath = null;
            try
            {
                filePath = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
            }
            catch (Exception e)
            {
                Log.exception(e, key_itemName + " - library item not found!");
            }

            string plainText = "";
            if (System.IO.Path.GetExtension(filePath) == ".rtf")
            {
                try
                {
                    string myText = File.ReadAllText(filePath);
                    rtBox.Rtf = myText;
                     plainText = rtBox.Text;
                     return plainText;  
                    ////richTextBox1.Clear();
                    ////richTextBox1.LoadFile(filePath, RichTextBoxStreamType.RichText);
                }
                catch (Exception ex)
                {
                    Log.exception(ex, filePath + " - This file has some format exception (not RTF)." + ex.Message);
                }
            }
            return "";
        }


        private void fillPeriod()
        {
            cboPeriod.Items.Clear();
            listObjID.Clear();
            PeriodEnumerator perEnum = _currentSOA.getPeriodEnumerator();
            
            while (perEnum.MoveNext())
            {
                Period per = (Period)perEnum.Current;
                if (!per.isSubPeriod())
                {
                    cboPeriod.Items.Add(per.getBriefDescription());
                    listObjID.Add(per.getObjID());
                }
            }
            perORsubPer = "period";
        }

        private void fillSubPeriod()
        {
            cboPeriod.Items.Clear();
            listObjID.Clear();
            PeriodEnumerator perEnum = _currentSOA.getPeriodEnumerator();

            while (perEnum.MoveNext())
            {
                Period per = (Period)perEnum.Current;
                EventScheduleEnumerator subPerChildren = _currentSOA.getPeriodChildren(per);
                while (subPerChildren.MoveNext())
                {
                    EventScheduleBase subPrd = subPerChildren.getCurrent();
                    try
                    {
                        Period p1 = (Period)subPrd;   /// Just making sure, Visits are not included.
                        if (p1.isSubPeriod())
                        {
                            cboPeriod.Items.Add(subPrd.getBriefDescription());
                            listObjID.Add(subPrd.getObjID());
                        }
                    }
                    catch (Exception ex)
                    {
                        //Ignore the exception, as we know Visits cannot be type casted as protocol.
                    }
                }

            }
            perORsubPer = "subperiod";
        }

        private void rdTask_CheckedChanged(object sender, EventArgs e)
        {
            //loadTasks
            if (rdPeriod.Checked)
            {
                lstTaskDesc.Items.Clear();
                lstTaskEvent.Items.Clear();
                htCheckedIndices.Clear();
                htFinalTE.Clear();
                fillPeriod();
            }
        }

        private void rdVisit_CheckedChanged(object sender, EventArgs e)
        {
            //loadVisits
            if (rdSubperiod.Checked)
            {
                lstTaskDesc.Items.Clear();
                lstTaskEvent.Items.Clear();
                htCheckedIndices.Clear();
                htFinalTE.Clear();
                fillSubPeriod();
            }
        }

        private void cboPeriod_SelectedIndexChanged(object sender, EventArgs e)
        {
            //When a selection is change for a Period/Sub-Period. Task List is being filled in Listbox.
            //Certain Rules are applied see individual methods for descripion of rules.

            //Clear the list view items.
            lstTaskDesc.Items.Clear();
            lstTaskEvent.Items.Clear();
            htCheckedIndices.Clear();
            htFinalTE.Clear();

            try
            {
                if (cboPeriod.SelectedIndex >= 0)
                {
                    sel_ObjID = listObjID[cboPeriod.SelectedIndex];
                    long pID = (long)sel_ObjID;
                    Period per = null;

                    if (perORsubPer == "period")
                    {
                        per = _currentSOA.getPeriodByID(pID);
                    }
                    else if (perORsubPer == "subperiod")
                    {
                        IList perEnum = _currentSOA.getPeriodEnumerator().getList();
                        foreach (Period pr in perEnum)
                        {
                            IList sp_List = _currentSOA.getPeriodChildren(pr).getList();
                            {
                                foreach (EventScheduleBase subprd in sp_List)
                                {
                                    if (subprd.getObjID().Equals(pID))
                                    {
                                        per = (Period)subprd;
                                        break;
                                    }
                                }
                            }
                        }
                    }

                    ArrayList tskList = GetTaskListforPeriod(per);
                    ArrayList taskEvents = GetTaskEvents(per,tskList);
                    FillListView_TaskVisit(tskList, taskEvents);  //FILL IN TASK ACCORDING TO CONFIGURATION FILE


                }//ENDIF 

            }//End TRy
            catch (Exception e1)
            {
                MessageBox.Show(e1.ToString());
            }
        }

        private ArrayList GetTaskEvents(Period p, ArrayList tskList)
        {
            ProtocolEventEnumerator peEnum = _currentSOA.getProtocolEventEnumerator(p);
            IList tskEvts = peEnum.getList();
            ArrayList TaskEvents_ = new ArrayList();
           
            foreach (long tskID in tskList)
            {
                Task tsk = _currentSOA.getTaskByID(tskID);
                if (_currentSOA.getTaskUsageState(tsk, p) != SOA.UsageTriState.None)
                {//Check ONLY if the Task has any events.
                    //TaskEvents_ = getTaskVisitsOrderedByVisit(tsk, tskEvts);
                    ArrayList tvl = new ArrayList(_currentSOA.getTaskVisitsForTask(tsk).getList());
                    
                    foreach (ProtocolEvent pe in tskEvts)
                    {
                        for (int i = 0; i < tvl.Count; i++)
                        {
                            TaskVisit tv = tvl[i] as TaskVisit;
                            if (tv != null && tv.getAssociatedVisitID() == pe.getObjID())
                            {
                                tv.setViewAngle(TaskVisit.ViewAngle.Task);
                                //tv.setScheduleID(getObjID());
                                TaskEvents_.Add(tv);
                                tvl[i] = null;
                            }
                        }
                    }
                }
            }

            return TaskEvents_;
        }

        private ArrayList ReadXML(string xmlPath)
        {
            //This method will read xml, and store NODES as objects in an ArrayList.
            ArrayList ListTasks = new ArrayList();         
            try
            {
                XmlDocument doc = new XmlDocument();
                doc.Load(xmlPath);

                // Select and display all Tasks.
                XmlNodeList nodeList;
                XmlElement root = doc.DocumentElement;
                nodeList = root.SelectNodes("/Tasks/Task");
                foreach (XmlNode taskEntry in nodeList)
                {
                    ListTasks.Add(taskEntry);  //Adding node to ArrayList.
                }
            }
            catch (Exception e)
            {
                Log.exception(e, e.Message + " - Configuration file is missing. Please, contact your configuration administrator");
                return ListTasks;
            }
            return ListTasks;
        }
        
        private ArrayList  GetTaskListforPeriod(Period p)
        {
            //This method returns array of taskID's with TaskEvents for the Selected PERIOD.

            TaskEnumerator tenum = _currentSOA.getTaskEnumerator();
            ArrayList taskList = new ArrayList();

            if (_currentSOA.getProtocolEventCount(p) > 0)
            {
                IList peEnum = _currentSOA.getProtocolEventEnumerator(p).getList();
                foreach (ProtocolEvent ev in peEnum)
                {
                    IList tvEnum = _currentSOA.getTaskVisitsForVisit(ev).getList();
                    foreach (TaskVisit tv in tvEnum)
                    {
                        if (taskList.IndexOf(tv.getAssociatedTaskID()) < 0)
                        {
                            taskList.Add(tv.getAssociatedTaskID());
                        }
                    }
                }
            }
            return taskList;
        }

        private void Initialize_LibItems()
        {
            //This method will be initializing Library Items from Bucket = "__procdesc".

            lm = LibraryManager.getInstance();
            IEnumerator bucketEnum = lm.getLibraryBuckets();
            while (bucketEnum.MoveNext())
            {
                LibraryBucket bucket = (LibraryBucket)bucketEnum.Current;
                if (bucket.getBucketName().StartsWith("__procdesc"))
                {
                    IEnumerator elementEnum = bucket.getElements().iterator();
                    while (elementEnum.MoveNext())
                    {
                        LibraryElement libElement = (LibraryElement)elementEnum.Current;
                        htBuckets.Add(libElement.getElementName(), libElement);
                    }
                }
            }
        }

        private void FillListView_TaskVisit(ArrayList ListofTasks,ArrayList taskEvents)
        { 
            /*/This Method will go thru each task and follow the steps below:
             *  [1] Check if TASK exists in the Config XML.
             *  [2] 
            */
            Task tsk = null;
            foreach (long tskID in ListofTasks)
            {
                tsk = _currentSOA.getTaskByID(tskID);
                if (tsk != null)
                {
                    if (TaskExistinXML(tsk.getBriefDescription()))
                    {
                        if (Fill_TaskEvents(taskEvents, tsk))
                        {
                           // GetChildNodes(_selNode, tsk, taskEvents);
                            _UpdateDefaultEntries(_selNode, tsk);
                        }
                    }
                }
            }
        }


        private bool Fill_TaskEvents(ArrayList taskEvents,Task _currTask)
        {
            bool hasEvents = false;
            bool hasVisitSpecific = false;

            //Check if "Visit Specific task exists for the selected node in XML
            foreach (XmlNode cnode in _selNode.ChildNodes)
            {
                if (cnode.Attributes.GetNamedItem("type").Value.ToLower().Equals("visit-specific"))
                {
                    hasVisitSpecific = true;
                    break;
                }
            }

            foreach (TaskVisit pe in taskEvents)
            {
                if (pe.getFullDescription().Trim().Length <= 0 && pe.getAssociatedTaskID() == _currTask.getObjID())  //If EMPTY
                {
                    if (lstTaskEvent.Items.ContainsKey(pe.getObjID().ToString() + ">" + _currTask.getObjID()) == false && hasVisitSpecific)
                    {
                        lstTaskEvent.Items.Add(pe.getObjID().ToString() + ">" + _currTask.getObjID(), _currTask.getActualDisplayValue() + " -- " +   pe.getActualDisplayValue(), lstTaskEvent.Items.Count + 1);                        
                        hasEvents = true;
                    }
                }
            }
            return hasEvents;
        }

        private void GetChildNodes(XmlNode _pNode,Task _currTask)
        {
           /* //Parse through all the child nodes for the selected node, and get type of each child node. 
            * [1]type = 'default' --> Go to "default", lib item.
            * [2]type = 'visit-specific' --> Go to "visit specific", lib item.   
            * 
            * taskEvents = get all task Events for selected Period/SubPeriod for the current Task.
            */

            try
            {
                foreach (XmlNode cnode in _pNode.ChildNodes)
                {
                    if (cnode.Attributes.GetNamedItem("type").Value.ToLower().Equals("default"))
                    {
                        //Call Method for handling default types
                        //Get Bucket from Hash Table, based on 

                        if (htBuckets.ContainsKey(cnode.Attributes.GetNamedItem("libraryItem").Value))
                        {
                            LibraryElement libElement = (LibraryElement) htBuckets[cnode.Attributes.GetNamedItem("libraryItem").Value];

                            string filePath = null;
                            try
                            {
                                filePath = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
                            }
                            catch (Exception e)
                            {
                                Log.exception(e, cnode.Attributes.GetNamedItem("libraryItem").Value + " - library item not found!");
                            }
                            _HandleDefault(filePath, _currTask);
                        }
                        else 
                        {
                            Log.trace(cnode.Attributes.GetNamedItem("libraryItem").Value + " - not found.");
                        }
                    }
                    else if (cnode.Attributes.GetNamedItem("type").Value.ToLower().Equals("visit-specific"))
                    {
                        //Call Method for handling visit-specific types
                        //Step1 is get all Task Events for Selected Period/Sub-PEr Vs Task. IF TaskEvent has empty Full Desc, then add it to the
                        //dialog.

                        int cnt = 0;  // This counter is for adding the items to right side List view only once.
                 

                               ListViewItem lvDescItem = new ListViewItem(cnode.Attributes.GetNamedItem("libraryItem").Value, lstTaskDesc.Items.Count + 1);
                               lvDescItem.Tag = _currTask.getObjID();  
                               lstTaskDesc.Items.Add(lvDescItem);
                               //lstTaskDesc.Items[lstTaskDesc.Items.Count]
                          // }
                       }
                    }
                }
            
            catch (Exception ex)
            { }

        }


        private void _UpdateDefaultEntries(XmlNode _pNode, Task _currTask)
        {
            try
            {
                foreach (XmlNode cnode in _pNode.ChildNodes)
                {
                    if (cnode.Attributes.GetNamedItem("type").Value.ToLower().Equals("default"))
                    {
                        //Call Method for handling default types
                        //Get Bucket from Hash Table, based on 

                        if (htBuckets.ContainsKey(cnode.Attributes.GetNamedItem("libraryItem").Value))
                        {
                            LibraryElement libElement = (LibraryElement)htBuckets[cnode.Attributes.GetNamedItem("libraryItem").Value];

                            string filePath = null;
                            try
                            {
                                filePath = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
                            }
                            catch (Exception e)
                            {
                                Log.exception(e, cnode.Attributes.GetNamedItem("libraryItem").Value + " - library item not found!");
                            }
                            _HandleDefault(filePath, _currTask);
                        }
                        else
                        {
                            Log.trace(cnode.Attributes.GetNamedItem("libraryItem").Value + " - not found.");
                        }
                    }
                }
            }
            catch (Exception e)
            {
                Log.exception(e, "Error in Updating Default Lib Item: " + e.Message);
            }
        }


        private void _HandleDefault(string filePath,Task _currTask)
        {
            /*This method will handle the default desc. There will be ONLY one default item,
             * it will get input as filePath (Selected Library Item),& Task.
             * Step1 :Compare : (a) If task desc does not matches sourcedesc (global desc) (b)if its is null 
             *          If "a" or "b" is true -- Goto Step 2.
             * Step 2: "If its doesnt matches or null, then update with with LibraryItem text
            */

            if ((_currTask.getFullDescription().Trim().Length <= 0) || ( _currTask.getFullDescription().Equals(_currTask.getSourceText())))
            {
                if (System.IO.Path.GetExtension(filePath) == ".rtf")
                {
                    //FileStream fs = File.OpenRead(filePath);
                    //StreamReader sr = new StreamReader(fs);
                    try
                    {
                        string myText = File.ReadAllText(filePath);
                        rtBox.Rtf = myText;
                        string plainText = rtBox.Text;

                        //UPDATING FULL DESCRIPTION.
                        
                        if (! _currTask.getFullDescription().Trim().Equals(plainText.Trim()))
                        {
                        //    MessageBox.Show(_currTask.getActualDisplayValue() + " has been updated. Msg1 ");
                            _currTask.setFullDescription(plainText.Trim());
                        }
                        rtBox.Clear();
                    }
                    catch (Exception ex)
                    {
                        Log.exception(ex, filePath + " - This file has some format exception." + ex.Message);
                    }
                }
            }
        }

        private bool TaskExistinXML(string taskName)
        {
            //This method, will go through all the Xmlnodes(in ArrayList), and SETS value in '_selNode';
            //It is being assumed and clarified, that Duplicates wont exists.

            foreach (XmlNode xiNode in customTaskDec)
            {
                if (xiNode.Attributes.GetNamedItem("name").Value.ToLower().Equals(taskName.ToLower()))
                {
                    _selNode = xiNode;
                    return true;
                }
            }

            //If not match, then set _selNode = null, and return
            _selNode = null;
            return false;
        }



        private void lstTaskEvent_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            //Handle the selection 
            lstTaskDesc.Items.Clear();
            
            if (lstTaskEvent.Items.Count > 0)
            {
                if (lstTaskEvent.FocusedItem == null)
                {
                    return;   //When switching between items,there would instance that focused item would be null.
                }
                if (strTaskEvent != lstTaskEvent.FocusedItem.Text)
                {
                    // MessageBox.Show(lstTaskEvent.FocusedItem.Text +">>>>" + lstTaskEvent.FocusedItem.Name);

                    string[] aParms = null;

                    if (lstTaskEvent.FocusedItem.Name.Length > 2)
                    {
                        aParms = lstTaskEvent.FocusedItem.Name.Split('>');
                    }

                    bool isBad = false;

                    if (aParms != null && aParms.Length == 2)
                    {
                        long tskID = PurdueUtil.getNumber(aParms[1], out isBad);
                        Task tsk = _currentSOA.getTaskByID(tskID);
                        // Pass on to fill in the ListView for Task Event Description
                        if (tsk != null)
                        {
                            if (TaskExistinXML(tsk.getBriefDescription()))
                            {
                                GetChildNodes(_selNode, tsk);
                            }
                        }
                    }
                    GetCheckMark(lstTaskEvent.FocusedItem.Index);
                }
                strTaskEvent = lstTaskEvent.FocusedItem.Text;
                
            }
        }

        private void GetCheckMark(int keyIDX)
        {
            //this method, will reset the Checked State for each lvItem. Setting will be stored in a Global HastTable
            if (htCheckedIndices.ContainsKey(keyIDX))
            {
               // ListView.CheckedListViewItemCollection chkItems = (ListView.CheckedListViewItemCollection)htCheckedIndices[keyIDX];
                ArrayList arrIndex = (ArrayList)htCheckedIndices[keyIDX];

                //PLEASE SEE: chkItems.getEnumerator
                //

                foreach (int idx in arrIndex)
                {
                    lstTaskDesc.Items[idx].Checked = true;                     
                }
            }
        }

        private void SetCheckMark(int keyIndex,ListView.CheckedIndexCollection checkedIndices)
        {
            /*This table has 2 important components:
               1. Key: It holds the index of TaskEvent
             * 2. Value: Checked Indices
             * 
             * First check if key is present, if Yes then update the indices.
             * if NO, then Add the indices.
             */
            try
            {
                ArrayList arrIndex = new ArrayList();
                string strTESelection = "";
                foreach (int idx in checkedIndices)
                {
                    strTESelection += lstTaskDesc.Items[idx].Text.Trim() + "|";
                    arrIndex.Add(idx);
                }

                if (htCheckedIndices.ContainsKey(keyIndex))
                {
                    htFinalTE[lstTaskEvent.Items[keyIndex].Name] = strTESelection;
                    htCheckedIndices[keyIndex] = arrIndex;
                }
                else
                {
                    htFinalTE.Add(lstTaskEvent.Items[keyIndex].Name, strTESelection);
                    htCheckedIndices.Add(keyIndex, arrIndex);
                }
            }
            catch (Exception ex)
            {
                Log.exception(ex, "Error in Setting up Check Marks - SetCheckMark() " + ex.ToString());
            }
        }

        private void lstTaskDesc_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            try
            {
                if (e.Item.Checked == true)
                {
                    e.Item.Focused = true;
                    e.Item.Selected = true;
                    e.Item.Tag = "set";  //Using the Tag field, so when this method is called during Lib Item creation which results in overriding "Hash Table"
                   //Idea is to set the tag, when first time an item is checked so its in HastTable, lets say then User Unchecks it, then we have to remove it.
                    SetCheckMark(lstTaskEvent.FocusedItem.Index,lstTaskDesc.CheckedIndices);
                }
                else
                {
                    //If Any item is unselected, update the HashTable.
                    if (e.Item.Tag.ToString() == "set")
                    {
                        SetCheckMark(lstTaskEvent.FocusedItem.Index, lstTaskDesc.CheckedIndices);
                    }
                }
            }
            catch(Exception ex)
            {
                Log.exception(ex, ex.Message + "--> Error in getting Library Item desc");
            }
        }


        private void lstTaskDesc_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
            try
            {

                if (lstTaskDesc.Items.Count > 0  && lstTaskDesc.FocusedItem.Text != null)
                {
                    if (strTaskDesc != lstTaskDesc.FocusedItem.Text)
                    {

                        if (htBuckets.ContainsKey(lstTaskDesc.FocusedItem.Text))
                        {
                            LibraryElement libElement = (LibraryElement)htBuckets[lstTaskDesc.FocusedItem.Text];
                            string filePath = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
                            if (System.IO.Path.GetExtension(filePath) == ".rtf")
                            {
                                //FileStream fs = File.OpenRead(filePath);
                                //StreamReader sr = new StreamReader(fs);
                                try
                                {
                                    richTextBox1.Clear();
                                    richTextBox1.LoadFile(filePath, RichTextBoxStreamType.RichText);
                                }
                                catch (Exception ex)
                                {
                                    Log.exception(ex, filePath + " - This file has some format exception." + ex.Message);
                                }
                            }
                        }
                    }
                    strTaskDesc = lstTaskDesc.FocusedItem.Text;
                }

            }
            catch (Exception ex)
            {
                Log.exception(ex, ex.Message + "--> Error in getting Library Item desc (Item Selection event)");
            }
        }
	}
}
