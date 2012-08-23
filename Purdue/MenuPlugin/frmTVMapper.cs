using System;
using System.Collections;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;

using Tspd.Context;
using Tspd.Tspddoc;
using Tspd.Icp;
using Tspd.Businessobject;
using Tspd.Utilities;
using Tspd.FormBase;
using Tspd.Bridge;
using System.Xml;

namespace TspdCfg.FastTrack.PlugIn
{


    public partial class frmTVMapper : Form
    {
        public LibraryManager lm = null;
        Hashtable htBuckets = new Hashtable();
        public List<TaskvisitMapper> TaskObjects = null;
        BusinessObjectMgr bom = null;
        ArrayList arrsoa = new ArrayList();
        ArrayList customtaskvisitDesc = new ArrayList();
        public XmlNode _selNode = null;
        public SOA _currsoa = null;
        public string strTaskDesc="";
        public string strTaskEvent = "";
        public RichTextBox hiddenrtf = new RichTextBox();
        public TaskvisitMapper currtvMapper = null;
        public string elementPath = "";
        public MacrosConfig mc = null;


        public frmTVMapper(string elePath)
        {
            elementPath = elePath;
            InitializeComponent();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (TaskVisittoSave())
            {
                frmSaveChanges frmSave = new frmSaveChanges(mc.getMessageByName("exception1").Text);
                DialogResult result = frmSave.ShowDialog();
                if (result == DialogResult.Yes)
                {
                    btnSave_Click(sender, e);
                    this.Close();
                }
                else if (result == DialogResult.No)
                {
                    this.Close();
                }
                else if (result == DialogResult.Cancel)
                { }
            }
            else
            {
                this.Close();
            }
        }

        private void frmTVMapper_Load(object sender, EventArgs e)
        {
            DesignerContext cm = DesignerContext.getInstance();
            DesignerDocBase doc = cm.getActiveBaseDocument();
            TaskObjects = new List<TaskvisitMapper>();
            

            bom = doc.getBom();            
            Load_SOA();

            //Customizing Labels
            Log.trace(doc.getTrialProject().getTemplateDirPath());
             
            string configFilePath = doc.getTrialProject().getTemplateDirPath() + "\\plugins\\MacrosConfig.xml";
            if(!System.IO.File.Exists(configFilePath))
            {
                configFilePath = BridgeProxy.getInstance().getSystemTemplatePath() +"\\plugins\\MacrosConfig.xml";
                if (!System.IO.File.Exists(configFilePath))
                {
                    MessageBox.Show("Configuration file is missing! Please contact your configuration adminsitrator.");
                    return;
                }
            }

            Customize_Labels(configFilePath);

            string mappingFilePath = doc.getTrialProject().getTemplateDirPath() + "\\plugins\\ProcDescMapping.xml";
            if (!System.IO.File.Exists(mappingFilePath))
            {
                MessageBox.Show("Procedure Mapping file is missing! Please contact your configuration adminsitrator.");
                return;
            }
           // MessageBox.Show(mappingFilePath);
            customtaskvisitDesc = ReadXML(mappingFilePath);

            Initialize_LibItems();
            btnSave.Enabled = false;
            btnAppend.Enabled = false;

        }

        //customize labels
        private void Customize_Labels(string ConfigFilePath)
        {
            mc= new MacrosConfig(ConfigFilePath,elementPath);

            if (mc != null)
            {
                this.Text = mc.getMessageByName("captiontext").Text;
                this.label1.Text = mc.getMessageByName("label1").Text;
                this.groupBox1.Text = mc.getMessageByName("grpbox1").Text;
                this.label6.Text = mc.getMessageByName("currdetails").Text;
                if (lstTask.Columns.Count > 0)
                {
                    this.lstTask.Columns[0].Text = mc.getMessageByName("coltext").Text;
                }

                if (lstTvDesc.Columns.Count > 0)
                {
                    this.lstTvDesc.Columns[0].Text = mc.getMessageByName("lsttaskeventdesc").Text;
                }
                this.restoretext.Text = mc.getMessageByName("linklabel").Text;
                this.groupBox2.Text = mc.getMessageByName("grpbox2").Text;
                this.label4.Visible = false;
                this.label3.Text = mc.getMessageByName("previewlabel").Text;
            }

        }


        //Load Methods
        private void Load_SOA()
        {
            IList soaList = bom.getAllSchedules().getList();
            foreach (SOA soa in soaList)
            {
                cmbSOA.Items.Add(soa.getActualDisplayValue());
                arrsoa.Add(soa.getObjID());
            }

            if (cmbSOA.Items.Count == 1)
            {
                cmbSOA.SelectedIndex = 0; 
            }
        }

        #region Load_LibraryItems
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
        #endregion

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
                Log.exception(e, e.Message + " - Procedure Mapping Configuration file is missing. Please, contact your configuration administrator");
                return ListTasks;
            }
            return ListTasks;
        }
        
        private void cmbSOA_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cmbSOA.SelectedIndex < 0)
            { return; }

            lstTask.Items.Clear();
            lstTvDesc.Items.Clear();
            TaskObjects.Clear();

            long soaid = (long) arrsoa[cmbSOA.SelectedIndex];
            SOA selSOA = bom.getSchedule(soaid);

            if (selSOA != null)
            {
                _currsoa = selSOA;
                GetTaskvisitList(selSOA);
            }
        }

        private void GetTaskvisitList(SOA currSOA)
        {
            //This method returns array of taskID's with TaskEvents for the Selected PERIOD.

           
            foreach (Period p in currSOA.getPeriodEnumerator().getList())
            {
                if (currSOA.getProtocolEventCount(p) > 0)
                {
                    IList peEnum = currSOA.getProtocolEventEnumerator(p).getList();
                    foreach (ProtocolEvent ev in peEnum)
                    {
                        IList tvEnum = currSOA.getTaskVisitsForVisit(ev).getList();
                        foreach (TaskVisit tv in tvEnum)
                        {
                            Task t = currSOA.getTaskByID(tv.getAssociatedTaskID());
                            if (t != null)
                            {
                                if (TaskExistinXML(t))
                                {
                                    TaskvisitMapper tvmp = new TaskvisitMapper();
                                    tvmp.currentdesc = tv.getFullDescription();
                                    tvmp.newDesc = tv.getFullDescription();
                                    tvmp.taskID = tv.getAssociatedTaskID();
                                    tvmp.taskvisitID = tv.getObjID();
                                    tvmp.visitID = ev.getObjID();
                                    TaskObjects.Add(tvmp);

                                    lstTask.Items.Add(t.getActualDisplayValue() + " - (" + ev.getActualDisplayValue() + ")", tv.getObjID().ToString());
                                }
                            }
                            else
                            {
                              //  MessageBox.Show(tv.getAssociatedTaskID() +"  is null");
                            }
                        }
                    }
                }
            }
            return;
        }

        private bool TaskExistinXML(Task tsk)
        {
            //This method, will go through all the Xmlnodes(in ArrayList), and SETS value in '_selNode';
            //It is being assumed and clarified, that Duplicates wont exists.

            if (tsk == null)
            {
               // MessageBox.Show(" TASKa  is null");
                return false;   
            }


            long tid = tsk.getObjID();
            string tname = tsk.getActualDisplayValue();

            //MessageBox.Show(tid.ToString() + "  - " + tname.ToString());
            //MessageBox.Show(customtaskvisitDesc.Count.ToString());

            foreach (XmlNode xiNode in customtaskvisitDesc)
            {
               
                if (xiNode.Attributes.GetNamedItem("taskid")!=null && xiNode.Attributes.GetNamedItem("taskid").Value.Equals(tid.ToString()))
                {
                    _selNode = xiNode;
                    return true;
                }
                else if (xiNode.Attributes.GetNamedItem("name").Value.ToLower().Equals(tname.ToLower()))
                {
                    _selNode = xiNode;
                    return true;
                }
            }

            //If not match, then set _selNode = null, and return
            _selNode = null;
            return false;
        }
               
        private TaskvisitMapper getTasksVisitObject(long tvID)
        {
            TaskvisitMapper tvmp = new TaskvisitMapper();

            var tvMapper = from tv in TaskObjects
                           //where task.Name..Equals("Pinal") & task.Type[1].Equals("Alpha")
                           where (tv.taskvisitID.Equals(tvID))
                           select new { objTVMP = tv };

            foreach (var item in tvMapper)
            {
                tvmp = item.objTVMP;
                break; 
            }
            return tvmp;
        }

        private void GetChildNodes(XmlNode _pNode, long taskvisitID)
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
                        int cnt = 0;  // This counter is for adding the items to right side List view only once.
                        ListViewItem lvDescItem = new ListViewItem(cnode.Attributes.GetNamedItem("libraryItem").Value, lstTvDesc.Items.Count + 1);
                        lvDescItem.Tag = taskvisitID;
                        lstTvDesc.Items.Add(lvDescItem);                        
                   
                }
            }

            catch (Exception ex)
            { }

        }

        private void lstTvDesc_ItemChecked(object sender, ItemCheckedEventArgs e)
        {
            if (e.Item.Checked == true)
            {
                btnSave.Enabled = true;
                e.Item.Focused = true;
                e.Item.Selected = true;

                if (strTaskDesc != e.Item.Text)
                {
                    if (htBuckets.ContainsKey(e.Item.Text))
                    {
                        LibraryElement libElement = (LibraryElement)htBuckets[e.Item.Text];
                        string filePath = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
                        if (System.IO.Path.GetExtension(filePath) == ".rtf")
                        {
                            try
                            {
                                rtfPreview.Clear();
                                rtfPreview.LoadFile(filePath, RichTextBoxStreamType.RichText);
                            }
                            catch (Exception ex)
                            {
                                Log.exception(ex, filePath + " - This file has some format exception." + ex.Message);
                            }
                        }
                    }
                }
                strTaskDesc = e.Item.Text;
            }

            if (lstTvDesc.CheckedItems.Count > 0)
            {
                btnAppend.Enabled = true;
            }
            else
            {
                btnAppend.Enabled = false;
            }
        }

        private void lstTvDesc_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
              try
            {

                if (lstTvDesc.Items.Count > 0  && lstTvDesc.FocusedItem.Text != null)
                {
                    if (strTaskDesc != lstTvDesc.FocusedItem.Text)
                    {
                        if (htBuckets.ContainsKey(lstTvDesc.FocusedItem.Text))
                        {
                            LibraryElement libElement = (LibraryElement)htBuckets[lstTvDesc.FocusedItem.Text];
                            string filePath = BridgeProxy.getInstance().loadLibraryElement(libElement.getLibraryBucketID(), libElement.getPKValue());
                            if (System.IO.Path.GetExtension(filePath) == ".rtf")
                            {                               
                                try
                                {
                                    rtfPreview.Clear();
                                    rtfPreview.LoadFile(filePath, RichTextBoxStreamType.RichText);
                                }
                                catch (Exception ex)
                                {
                                    Log.exception(ex, filePath + " - This file has some format exception." + ex.Message);
                                }
                            }
                        }
                    }
                    strTaskDesc = lstTvDesc.FocusedItem.Text;
                }

            }
            catch (Exception ex)
            {
                Log.exception(ex, ex.Message + "--> Error in getting Library Item desc (Item Selection event)");
            }
        }

        private void btnAppend_Click(object sender, EventArgs e)
        {
            long tvID =0;
            TaskvisitMapper tvmp = null;
            if (lstTvDesc.CheckedItems.Count > 0)
            {
                tvID = (long) lstTvDesc.CheckedItems[0].Tag;
                tvmp = getTasksVisitObject(tvID);
            }

            foreach (ListViewItem lvt in lstTvDesc.CheckedItems)
            {
                if (lvt.Text != null && lvt.Text.Length > 0) 
                {
                   rtf.Text += ReadfromLibrary(lvt.Text);
                   rtfPreview.Clear();
                }
                lvt.Checked = false;
            }

            if (tvmp != null && tvID != 0)
            {
                tvmp.hasdescChanged = true;
                tvmp.newDesc = rtf.Text;
                btnSave.Enabled = true;
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
                    hiddenrtf.Rtf = myText;
                    plainText = hiddenrtf.Text;
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

        private void lstTask_ItemSelectionChanged(object sender, ListViewItemSelectionChangedEventArgs e)
        {
             ListViewItem lvt=null;
            if (e.Item != null)
            {
                //MessageBox.Show(e.Item.Text + " - "
                //    + e.IsSelected.ToString());

                if(e.IsSelected == false &&   currtvMapper != null)
                    {
                        currtvMapper.newDesc = rtf.Text;
                        btnSave.Enabled = true;
                    }

            }

            if (lstTask.Items.Count > 0 && lstTask.FocusedItem != null)
            {
                if (lstTvDesc.CheckedItems.Count > 0)
                {
                    DialogResult result = MessageBox.Show(mc.getMessageByName("exception2").Text, "", MessageBoxButtons.YesNo);
                    if (result == DialogResult.Yes)
                    {
                        btnAppend_Click(sender, e);
                        btnSave.Enabled = true;
                        //return;
                    }
                    else if (result == DialogResult.No)
                    {
                        if (currtvMapper != null && rtf.TextLength>0)
                        {
                            currtvMapper.newDesc = rtf.Text;
                            btnSave.Enabled = true;
                        }
                        //  return;
                    }
                    else if (result == DialogResult.Cancel)
                    {
                        ////MessageBox.Show(e.ItemIndex.ToString()); 
                        //if (e.Item != null && e.IsSelected == false)
                        //{
                        //     lvt = e.Item;
                        //    lvt.Selected = true;
                        //}
                        //return;
                    }
                }

          
              

                lstTvDesc.Items.Clear();
                rtf.Clear();
                rtfPreview.Clear();

                if (strTaskEvent != lstTask.FocusedItem.Text)
                {
                   
                  //  rtf.Clear();

                    long tvID = Convert.ToInt64(lstTask.FocusedItem.ImageKey);
                    currtvMapper = getTasksVisitObject(tvID);
                    if (_currsoa != null)
                    {
                        Task t = _currsoa.getTaskByID(currtvMapper.taskID);
                        if (TaskExistinXML(t))
                        {
                            GetChildNodes(_selNode, tvID);
                            rtf.Text = currtvMapper.newDesc;
                        }
                    }
                }
                strTaskEvent = lstTask.FocusedItem.Text;
                btnSave.Enabled = true;
            }
            else
            {
                currtvMapper = null;
            }
        }

        private void restoretext_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            if (currtvMapper != null)
            {
                if (rtf.Text.Equals(currtvMapper.currentdesc))
                {
                    MessageBox.Show(mc.getMessageByName("exception3").Text);
                }
                else
                {
                  DialogResult result=  MessageBox.Show(mc.getMessageByName("exception4").Text,"",MessageBoxButtons.YesNo);
                  if (result == DialogResult.Yes)
                  {
                      rtf.Text = currtvMapper.currentdesc;
                      btnSave.Enabled = true;
                  }
                }

 
            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            try
            {
                //foreach (var 
                if (TaskVisittoSave())
                {
                    SaveTaskVisitDesc();
                }
                else
                {
                    btnSave.Enabled = false;
                    
                }
            }
            catch (Exception ex)
            {
                Log.exception(ex, ex.Message + ex.StackTrace);
                MessageBox.Show(ex.Message);
            }
        }

        public void SaveTaskVisitDesc()
        {
            TaskvisitMapper tvmp = new TaskvisitMapper();

            var tvMapper = from tv in TaskObjects
                           //where task.Name..Equals("Pinal") & task.Type[1].Equals("Alpha")
                           where (!tv.currentdesc.Equals(tv.newDesc))
                           select new { objTVMP = tv };

            foreach (var item in tvMapper)
            {
                TaskVisit tv= _currsoa.getTaskVisitById(item.objTVMP.taskvisitID);                
                tv.setFullDescription(item.objTVMP.newDesc);
                item.objTVMP.currentdesc = item.objTVMP.newDesc;
            }
            btnSave.Enabled = false;
        
        }

        public bool TaskVisittoSave()
        {
            TaskvisitMapper tvmp = new TaskvisitMapper();

            var tvMapper = from tv in TaskObjects
                           //where task.Name..Equals("Pinal") & task.Type[1].Equals("Alpha")
                           where (!tv.currentdesc.Equals(tv.newDesc))
                           select new { objTVMP = tv };

            foreach (var item in tvMapper)
            {
                return true;
                break;
            }
           

            return false;
        }

    
}


   
   
    


    public class TaskvisitMapper 
    {
        public long taskvisitID { get; set; }
        public bool hasdescChanged { get; set; }
        public string currentdesc { get; set; }
        public string newDesc { get; set; }
        public long visitID { get; set; }
        public long taskID { get; set; }
    }

}
