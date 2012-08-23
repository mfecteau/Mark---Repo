using System;
using System.Collections;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

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
    public partial class frmTaskSeq : Form
    {

        # region variableDeclaration
        BusinessObjectMgr bom = null;
        ArrayList arrsoa = new ArrayList();
        ArrayList arrVisits = new ArrayList();
        SOA _currsoa = null;
        MacrosConfig mc = null;
        string elementPath = "";
        int currSelIdx = -1;
        int currSelSOAIdx = -1;

        #endregion

        public frmTaskSeq(string elePath)
        {
            elementPath = elePath;
            InitializeComponent();
        }

        private void frmTaskSeq_Load(object sender, EventArgs e)
        {
            DesignerContext cm = DesignerContext.getInstance();
            DesignerDocBase doc = cm.getActiveBaseDocument();
            bom = doc.getBom();


            btnSave.Enabled = false;

            //Customizing Labels
            Log.trace(doc.getTrialProject().getTemplateDirPath());

            string configFilePath = doc.getTrialProject().getTemplateDirPath() + "\\plugins\\MacrosConfig.xml";
            if (!System.IO.File.Exists(configFilePath))
            {
                configFilePath = BridgeProxy.getInstance().getSystemTemplatePath() + "\\plugins\\MacrosConfig.xml";
                if (!System.IO.File.Exists(configFilePath))
                {
                    MessageBox.Show("Configuration file is missing! Please contact your configuration adminsitrator.");
                    return;
                }
            }
                Customize_Labels(configFilePath);


            //Load SOA
                Load_SOA();
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


        //customize labels
        private void Customize_Labels(string ConfigFilePath)
        {
            mc = new MacrosConfig(ConfigFilePath, elementPath);

            if (mc != null)
            {
                this.Text = mc.getMessageByName("captiontext").Text;
                this.label1.Text = mc.getMessageByName("soalabel").Text;
                this.label2.Text = mc.getMessageByName("visitlabel").Text;
                this.groupBox2.Text = mc.getMessageByName("grpbox1").Text;
                this.groupBox3.Text = mc.getMessageByName("grpbox2").Text;
                this.btnReset.Text = mc.getMessageByName("btnreset").Text;
                this.btnSave.Text = mc.getMessageByName("btnsave").Text; 
                
                this.btnClose.Text = mc.getMessageByName("btnclose").Text;

                if (lstTask.Columns.Count > 0)
                {
                    this.lstTask.Columns[0].Text = mc.getMessageByName("col1text").Text;
                    this.lstTask.Columns[1].Text = mc.getMessageByName("col2text").Text;
                }

               
            }

        }

        private void GetVisitforSOA(SOA _soa)
        {
            //_soa.getal

            cmbVisit.Items.Clear();
            arrVisits.Clear();   //Clearing the Visit List; If SOA Changes.

            foreach (Period p in _soa.getPeriodEnumerator().getList())
            {
                foreach (ProtocolEvent pe in _soa.getProtocolEventEnumerator(p).getList())
                {
                    if (_soa.getTaskVisitsWithChildrenForVisit(pe).getList().Count > 0)
                    {
                        cmbVisit.Items.Add(pe.getActualDisplayValue());
                        arrVisits.Add(pe.getObjID());
                    }
                }
            }
        }

        private void cmbVisit_SelectedIndexChanged(object sender, EventArgs e)
        {
        //    MessageBox.Show("2");
            if (cmbVisit.SelectedIndex < 0 || currSelIdx == cmbVisit.SelectedIndex)
            { return; }


        
            if (btnSave.Enabled)
            {
                frmSaveChanges frmSave = new frmSaveChanges(mc.getMessageByName("exception1").Text);
                DialogResult result = frmSave.ShowDialog();
                if (result == DialogResult.Yes)
                {
                    btnSave_Click(sender, e);  //Save and move on.
                    // this.Close();
                }
                else if (result == DialogResult.No)
                {
                    // this.Close(); // Continue
                }
                else if (result == DialogResult.Cancel)
                {
                    cmbVisit.SelectedIndex = currSelIdx;
                    return;   //Return from here, let user decide what to do
                }
            }


            long visitid = (long)arrVisits[cmbVisit.SelectedIndex];
            //SOA selSOA = bom.getSchedule(soaid);
            ProtocolEvent pe = _currsoa.getProtocolEventByID(visitid);
            btnSave.Enabled = false;
            currSelIdx = cmbVisit.SelectedIndex;
            if (pe != null)
            {
                //_currsoa = selSOA;
                FillTasks(pe);
                //  GetTaskvisitList(selSOA);
            }
        }

        private void FillTasks(ProtocolEvent pe)
        {
            try
            {
                lstTask.Items.Clear();
                ListViewItem lvt = null;
                 IList tvEnum = _currsoa.getTaskVisitsForVisit(pe).getList();
                 foreach (TaskVisit tv in tvEnum)
                 {
                     Task t  = _currsoa.getTaskOfTaskVisit(tv);
                     if (t != null)
                         lvt = new ListViewItem();
                         lvt.Text = tv.getSequence().ToString();
                         lvt.Tag = tv.getObjID();
                         lvt.SubItems.Add( t.getActualDisplayValue());
                       //  lvt.SubItems.Add(tv.getObjID().ToString());
                         lstTask.Items.Add(lvt);

                         ////Add the items to the ListView.
                        // listView1.Items.AddRange(new ListViewItem[] { item1, item2, item3 });
                     
                 }
            }
            catch (Exception ex)
            {
 
            }
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            if (btnSave.Enabled)
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
                this.Close();   //nothing to save, just close.
            }
        }

        private void cmbSOA_SelectedIndexChanged_1(object sender, EventArgs e)
        {
            if (cmbSOA.SelectedIndex < 0 || currSelSOAIdx == cmbSOA.SelectedIndex)
            { return; }

            if (btnSave.Enabled)
            {
                frmSaveChanges frmSave = new frmSaveChanges(mc.getMessageByName("exception1").Text);
                DialogResult result = frmSave.ShowDialog();
                if (result == DialogResult.Yes)
                {
                    btnSave_Click(sender, e);  //Save and move on.
                   // this.Close();
                }
                else if (result == DialogResult.No)
                {
                   // this.Close(); // Continue
                }
                else if (result == DialogResult.Cancel)
                {
                    cmbSOA.SelectedIndex = currSelSOAIdx;
                    return;   //Return from here, let user decide what to do
                }
            }

            cmbVisit.Items.Clear();
            lstTask.Items.Clear();
            btnSave.Enabled = false;
            long soaid = (long)arrsoa[cmbSOA.SelectedIndex];
            currSelSOAIdx = cmbSOA.SelectedIndex;
            SOA selSOA = bom.getSchedule(soaid);

            if (selSOA != null)
            {
                _currsoa = selSOA;
                GetVisitforSOA(_currsoa);
                //  GetTaskvisitList(selSOA);
            }
        }

        private void btnReset_Click(object sender, EventArgs e)
        {
            int cnt =1;

            if (lstTask.Items.Count > 0)
            {
                btnSave.Enabled = true;
            }

            foreach (ListViewItem lvt in lstTask.Items)
            {
                lvt.Text = cnt.ToString();
                cnt++;
            }
        }

        private void btnUp_Click(object sender, EventArgs e)
        {
            if (lstTask.SelectedItems.Count> 0)
            {
               ListViewItem item = lstTask.SelectedItems[0];

                int index = lstTask.SelectedItems[0].Index;

                    index--;

                    if (index >=0)
                    {

                        lstTask.Items.Remove(item);
                        

                        lstTask.Items.Insert(index, item);

                        item.Selected = true;

                        lstTask.Focus();
                        btnSave.Enabled = true;

                    } 

                   
            }
        }

        private void btnDown_Click(object sender, EventArgs e)
        {
            if ( lstTask.SelectedItems.Count>0)
            {
                ListViewItem item = lstTask.SelectedItems[0];
                int index = lstTask.SelectedItems[0].Index;
                index++;

                if (index < lstTask.Items.Count)
                {

                    lstTask.Items.Remove(item);

                    lstTask.Items.Insert(index, item);

                    item.Selected = true;

                    lstTask.Focus();
                    btnSave.Enabled = true;
                }

            }
        }

        private void btnSave_Click(object sender, EventArgs e)
        {
            long objId = 0;
            int cnt = 1;
            foreach (ListViewItem lvt in lstTask.Items)
            {
                objId = (long)lvt.Tag;
                TaskVisit tv =   _currsoa.getTaskVisitById(objId);
                tv.setSequence(cnt);
                lvt.Text = cnt.ToString();
                cnt++;
            }
            btnSave.Enabled = false;
        }

       

        //private void cmbVisit_SelectedValueChanged(object sender, EventArgs e)
        //{
        //    MessageBox.Show("1");

        //    if (btnSave.Enabled)
        //    {
        //        frmSaveChanges frmSave = new frmSaveChanges(mc.getMessageByName("exception1").Text);
        //        DialogResult result = frmSave.ShowDialog();
        //        if (result == DialogResult.Yes)
        //        {
        //            btnSave_Click(sender, e);  //Save and move on.
        //            // this.Close();
        //        }
        //        else if (result == DialogResult.No)
        //        {
        //            // this.Close(); // Continue
        //        }
        //        else if (result == DialogResult.Cancel)
        //        {
                    
        //            return;   //Return from here, let user decide what to do
        //        }
        //    }
        //}

        


    }
}
