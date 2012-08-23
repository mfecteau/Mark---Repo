using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Tspd.Icp;
using Tspd.Tspddoc;
using Tspd.Businessobject;

namespace TspdCfg.Purdue.DynTmplts
{
    /// <summary>
    /// Summary description for OutcomeSelection.
    /// </summary>
    public class CriteriaSelection : System.Windows.Forms.Form
    {
        private ArrayList criteria = new ArrayList();
        public string var_OutcomeType, OutcomeLabel;
        public Button btnOK;
        public LinkLabel linkLabel1;
        public Label lblSubtype;
        public CheckedListBox chkLstSubType;
        public ComboBox cmbCriteriaset;
        public Label lblCriteria;
        public bool boolChkAll = true;
        private bool subtype = true;

        public string var_Type = "",varLabel ="";
                    
        public int cntType = 0;
        
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.Container components = null;

        public CriteriaSelection()
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
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                if (components != null)
                {
                    components.Dispose();
                }
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code
        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.btnOK = new System.Windows.Forms.Button();
            this.linkLabel1 = new System.Windows.Forms.LinkLabel();
            this.lblSubtype = new System.Windows.Forms.Label();
            this.chkLstSubType = new System.Windows.Forms.CheckedListBox();
            this.cmbCriteriaset = new System.Windows.Forms.ComboBox();
            this.lblCriteria = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // btnOK
            // 
            this.btnOK.Location = new System.Drawing.Point(246, 198);
            this.btnOK.Name = "btnOK";
            this.btnOK.Size = new System.Drawing.Size(78, 21);
            this.btnOK.TabIndex = 11;
            this.btnOK.Text = "OK";
            this.btnOK.UseVisualStyleBackColor = true;
            this.btnOK.Click += new System.EventHandler(this.btnOK_Click);
            // 
            // linkLabel1
            // 
            this.linkLabel1.AutoSize = true;
            this.linkLabel1.Location = new System.Drawing.Point(3, 77);
            this.linkLabel1.Name = "linkLabel1";
            this.linkLabel1.Size = new System.Drawing.Size(52, 13);
            this.linkLabel1.TabIndex = 10;
            this.linkLabel1.TabStop = true;
            this.linkLabel1.Text = "Check All";
            this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
            // 
            // lblSubtype
            // 
            this.lblSubtype.AutoSize = true;
            this.lblSubtype.Location = new System.Drawing.Point(3, 56);
            this.lblSubtype.Name = "lblSubtype";
            this.lblSubtype.Size = new System.Drawing.Size(159, 13);
            this.lblSubtype.TabIndex = 9;
            this.lblSubtype.Text = "Select Sub-Type to be included:";
            // 
            // chkLstSubType
            // 
            this.chkLstSubType.FormattingEnabled = true;
            this.chkLstSubType.Location = new System.Drawing.Point(3, 98);
            this.chkLstSubType.Name = "chkLstSubType";
            this.chkLstSubType.Size = new System.Drawing.Size(321, 94);
            this.chkLstSubType.TabIndex = 8;
            // 
            // cmbCriteriaset
            // 
            this.cmbCriteriaset.FormattingEnabled = true;
            this.cmbCriteriaset.Location = new System.Drawing.Point(3, 27);
            this.cmbCriteriaset.Name = "cmbCriteriaset";
            this.cmbCriteriaset.Size = new System.Drawing.Size(324, 21);
            this.cmbCriteriaset.TabIndex = 7;
            // 
            // lblCriteria
            // 
            this.lblCriteria.AutoSize = true;
            this.lblCriteria.Location = new System.Drawing.Point(3, 6);
            this.lblCriteria.Name = "lblCriteria";
            this.lblCriteria.Size = new System.Drawing.Size(64, 13);
            this.lblCriteria.TabIndex = 6;
            this.lblCriteria.Text = "Select Type";
            // 
            // CriteriaSelection
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(340, 226);
            this.Controls.Add(this.btnOK);
            this.Controls.Add(this.linkLabel1);
            this.Controls.Add(this.lblSubtype);
            this.Controls.Add(this.chkLstSubType);
            this.Controls.Add(this.cmbCriteriaset);
            this.Controls.Add(this.lblCriteria);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "CriteriaSelection";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Select Criteria";
            this.TopMost = true;
            this.Load += new System.EventHandler(this.CriteriaSelection_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion

        private void CriteriaSelection_Load(object sender, EventArgs e)
        {
        }
        ArrayList ctype,subTypeList;

        public void LoadCriteria(TspdDocument tspdDoc_ ,BusinessObjectMgr bom_,IcpSchemaManager icpschemamgr_, bool disp_subtype)
        {
            subtype = disp_subtype;
            cmbCriteriaset.Items.Clear();


            if (!disp_subtype)
            {
                linkLabel1.Visible = false;
                chkLstSubType.Visible = false;
                lblSubtype.Visible = false;
                btnOK.Top = cmbCriteriaset.Bottom + 10;
                this.Height = btnOK.Bottom + 30 ;
            }

            ctype = icpschemamgr_.getEnumPairs("EntranceCriterionTypes");
          
            //ctype = bom_.getIcpSchemaMgr().getEnumPairs("EntranceCriterionTypes");

            for (int i = 0; i < ctype.Count - 1; i++)
            {
                EnumPair ep = (EnumPair)ctype[i];
                cmbCriteriaset.Items.Add(ep.getUserLabel());                
                cntType++;   //Counter for other types.  So this will cnt will includes all types except "Others".
            }
           
            
            ElementListHelpers elh = new ElementListHelpers(tspdDoc_);
            IList critList = elh.getLiveChooserEntryListForCriteria();
            IEnumerator critIter = critList.GetEnumerator();
            while (critIter.MoveNext())
            {
                Criterion crit = (Criterion)critIter.Current;
                if (crit.getCriterionType().ToLower() == "other")
                {
                   // MessageBox.Show(crit.getOtherCriterion().ToString());
                    cmbCriteriaset.Items.Add(crit.getOtherCriterion());
                }         
            }

            //Get Subtype Filled in.
            subTypeList = bom_.getIcpSchemaMgr().getEnumPairs("EntranceCriterionClassifierTypes");

            for (int i = 0; i < subTypeList.Count ; i++)
            {
                EnumPair ep = (EnumPair)subTypeList[i];
                chkLstSubType.Items.Add(ep.getUserLabel());
            }

            chkLstSubType.Items.Add("unclassified");
            

            //To set the items in the chkList to CHECK ALL ON
            CheckSubTypes();
        
            this.ShowDialog();
        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            //To set the items in the chkList to CHECK ALL/OFF
            CheckSubTypes();
        }

        private void CheckSubTypes()
        {
            //To set the items in the chkList to CHECK ALL/OFF
            for (int i = 0; i < chkLstSubType.Items.Count; i++)
            {
                chkLstSubType.SetItemChecked(i, boolChkAll);
            }

            if (chkLstSubType.Items.Count > 0)
            {
                if (boolChkAll)
                {
                    boolChkAll = false;
                }
                else
                {
                    boolChkAll = true;
                }
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            if (cmbCriteriaset.SelectedIndex != -1)
            {
                if (cmbCriteriaset.SelectedIndex >= cntType)
                {
                    var_Type = "other";
                    varLabel = cmbCriteriaset.SelectedItem.ToString();
                }
                else
                {
                    EnumPair ep = (EnumPair)ctype[cmbCriteriaset.SelectedIndex];  //deduct one as we are not including other
                    var_Type = ep.getSystemName();
                    varLabel = ep.getUserLabel();
                }

                this.DialogResult = DialogResult.OK;
                this.Close();  //Closing form.
            }
            else
            {
                var_Type = "FT_NA";  ///If nothing is selected and HIT OK
                varLabel = "";
                this.DialogResult = DialogResult.OK;
                this.Close();

            }
        }
    }
}