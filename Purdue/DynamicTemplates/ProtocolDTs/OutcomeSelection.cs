using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using Tspd.Businessobject;
using Tspd.Icp;
namespace TspdCfg.Purdue.DynTmplts
{
	/// <summary>
	/// Summary description for OutcomeSelection.
	/// </summary>
	public class OutcomeSelection : System.Windows.Forms.Form
	{
		public string var_Type,varLabel;
        public System.Windows.Forms.ComboBox cmbOutcome;
        public System.Windows.Forms.Button button1;
		public System.Windows.Forms.Label label1;
        public RadioButton rdbyOutcome;
        public RadioButton rdbyObjective;
        public ToolTip outcomeTooltip;
        private IContainer components;

        public OutcomeSelection()
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
            this.components = new System.ComponentModel.Container();
            this.cmbOutcome = new System.Windows.Forms.ComboBox();
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.rdbyOutcome = new System.Windows.Forms.RadioButton();
            this.rdbyObjective = new System.Windows.Forms.RadioButton();
            this.outcomeTooltip = new System.Windows.Forms.ToolTip(this.components);
            this.SuspendLayout();
            // 
            // cmbOutcome
            // 
            this.cmbOutcome.Location = new System.Drawing.Point(24, 70);
            this.cmbOutcome.Name = "cmbOutcome";
            this.cmbOutcome.Size = new System.Drawing.Size(272, 21);
            this.cmbOutcome.TabIndex = 0;
            // 
            // button1
            // 
            this.button1.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button1.Location = new System.Drawing.Point(115, 97);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(68, 26);
            this.button1.TabIndex = 1;
            this.button1.Text = "OK";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.Location = new System.Drawing.Point(24, 54);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(120, 16);
            this.label1.TabIndex = 2;
            this.label1.Text = "Select Outcome type:";
            // 
            // rdbyOutcome
            // 
            this.rdbyOutcome.AutoSize = true;
            this.rdbyOutcome.Location = new System.Drawing.Point(24, 10);
            this.rdbyOutcome.Name = "rdbyOutcome";
            this.rdbyOutcome.Size = new System.Drawing.Size(14, 13);
            this.rdbyOutcome.TabIndex = 3;
            this.rdbyOutcome.TabStop = true;
            this.rdbyOutcome.UseVisualStyleBackColor = true;
            this.rdbyOutcome.CheckedChanged += new System.EventHandler(this.rdbyOutcome_CheckedChanged);
            // 
            // rdbyObjective
            // 
            this.rdbyObjective.AutoSize = true;
            this.rdbyObjective.Location = new System.Drawing.Point(24, 33);
            this.rdbyObjective.Name = "rdbyObjective";
            this.rdbyObjective.Size = new System.Drawing.Size(14, 13);
            this.rdbyObjective.TabIndex = 4;
            this.rdbyObjective.TabStop = true;
            this.rdbyObjective.UseVisualStyleBackColor = true;
            this.rdbyObjective.CheckedChanged += new System.EventHandler(this.rdbyObjective_CheckedChanged);
            // 
            // OutcomeSelection
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(322, 128);
            this.Controls.Add(this.rdbyObjective);
            this.Controls.Add(this.rdbyOutcome);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.cmbOutcome);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedToolWindow;
            this.Name = "OutcomeSelection";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.TopMost = true;
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion
ArrayList arrTypes = new ArrayList();

Tspd.Icp.BusinessObjectMgr bom_ = null;

		public void LoadOutcomes(BusinessObjectMgr thisBom_,String _type)
		{
			//IcpInstanceManager icr= thisBom_.getIcp().getTypedDisplayValue(DesignDefines.OverallStudyOutcomeType, true);


            bom_ = thisBom_;

            if (_type.ToLower() == "optional")
            {
                SetDisplayOptions(true);
                rdbyOutcome.Checked = true;
            }
            else if (_type.ToLower() == "outcometype")
            {
                SetDisplayOptions(false);
                rdbyOutcome.Checked = true;
                
            }
            else if (_type.ToLower() == "objectivetype")
            {
                SetDisplayOptions(false);
                rdbyObjective.Checked = true;
            }

			this.ShowDialog();
		}

        private void SetDisplayOptions(bool flag)
        {
            if (!flag)
            {
                rdbyObjective.Visible = false;
                rdbyOutcome.Visible = false;
                label1.Top = rdbyOutcome.Top;
                cmbOutcome.Top = rdbyObjective.Top;

                button1.Top = cmbOutcome.Bottom + 10;
                this.Height = this.Height - (rdbyObjective.Height * 2 + 10);
            }
        }

        private void FillbyOutcometype()
        {
            cmbOutcome.Items.Clear();
            arrTypes = bom_.getIcpSchemaMgr().getEnumPairs("OutcomeTypes");

            int i = 0;

            for (i = 0; i < arrTypes.Count - 1; i++)
            {
                EnumPair ep = (EnumPair)arrTypes[i];
                cmbOutcome.Items.Add(ep.getUserLabel());
            }

            OutcomeEnumerator ocEnum = bom_.getOutcomes();
            foreach (Outcome oc in ocEnum.getList())
            {
                if (oc.getOutcomeType().ToLower() == "other")
                {
                    if (!cmbOutcome.Items.Contains(oc.getOtherOutcome()))
                    {
                        cmbOutcome.Items.Add(oc.getOtherOutcome());
                    }
                }
            }           
            
        }

        private void FillbyObjectivetype()
        {
            cmbOutcome.Items.Clear();
            arrTypes = bom_.getIcpSchemaMgr().getEnumPairs("ObjectiveTypes");
            int i = 0;

            for (i = 0; i < arrTypes.Count - 1; i++)
            {
                EnumPair ep = (EnumPair)arrTypes[i];
                cmbOutcome.Items.Add(ep.getUserLabel());
            }

            ObjectiveEnumerator objEnum = bom_.getObjectives();

            foreach (Objective obj in objEnum.getList())
            {
                if (obj.getObjectiveType().ToLower() == "other")
                {
                    if (!cmbOutcome.Items.Contains(obj.getOtherObjective()))
                    {
                        cmbOutcome.Items.Add(obj.getOtherObjective());
                    }
                }
            }
        }

		private void button1_Click(object sender, System.EventArgs e)
		{
            if (cmbOutcome.SelectedIndex != -1)
            {
                if (cmbOutcome.SelectedIndex >= arrTypes.Count)
                {
                    var_Type = "other";
                    varLabel = cmbOutcome.SelectedItem.ToString();
                }
                else
                {
                    EnumPair ep = (EnumPair)arrTypes[cmbOutcome.SelectedIndex];
                    var_Type = ep.getSystemName();
                    varLabel = ep.getUserLabel();
                }
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
            else
            {
                var_Type = "FT_NA";  ///If nothing is selected and HIT OK
                varLabel = "";
                this.DialogResult = DialogResult.OK;
                this.Close();
            }
		}

        public void Fill_Assessments(BusinessObjectMgr thisBom_)
        {
            arrTypes = thisBom_.getIcpSchemaMgr().getEnumPairs("PurposeTypes");
            int i = 0;

            for (i = 0; i < arrTypes.Count - 1; i++)
            {
                EnumPair ep = (EnumPair)arrTypes[i];
                cmbOutcome.Items.Add(ep.getUserLabel());
            }
        }

        private void rdbyOutcome_CheckedChanged(object sender, EventArgs e)
        {
            FillbyOutcometype();
        }

        private void rdbyObjective_CheckedChanged(object sender, EventArgs e)
        {
            FillbyObjectivetype();
        }

       
	}
}
